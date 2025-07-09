/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Excel */

// API configuration
import { updateStatus } from '../utils/taskpaneUtils'
import { syncData2 } from './syncData2'
import { timeAsync, startTiming, endTiming } from '../utils/performanceUtils'
import { withExponentialBackoff } from '../utils/chunkingUtils'

const API_BASE_URL = 'http://localhost:3001/api'

interface Product {
  id: number
  name: string
  price: number
}

interface ApiResponse {
  products: Product[]
}

interface CellUpdateRequest {
  id: number
  field: string
  value: string | number
}

interface CellUpdateResponse {
  success: boolean
  updatedProduct?: Product
  error?: string
}

interface ExcelFormula {
  id: string
  name: string
  description: string
  formula: string
  defaultLocation?: string
}

interface FormulasResponse {
  formulas: ExcelFormula[]
}

// Formula cache interface for performance optimization
interface FormulaCache {
  formulas: ExcelFormula[]
  timestamp: number
  dataRowCount: number
  summaryStartRow: number
  processedFormulas: Map<string, any>
}

// Progress tracking for formula operations
interface FormulaProgress {
  total: number
  completed: number
  current: string
  errors: string[]
}

// Range validation result
interface RangeValidationResult {
  isValid: boolean
  error?: string
  normalizedRange?: string
}

// Formula processing options
interface FormulaProcessingOptions {
  enableCache?: boolean
  batchSize?: number
  enableProgressTracking?: boolean
  onProgress?: (progress: FormulaProgress) => void
  validateRanges?: boolean
  enableRetry?: boolean
  maxRetries?: number
}

let data: ApiResponse = { products: [] }

// Formula cache for performance optimization
let formulaCache: FormulaCache | null = null
const FORMULA_CACHE_TTL = 300000 // 5 minutes cache TTL
const DEFAULT_BATCH_SIZE = 10
const MAX_CONCURRENT_OPERATIONS = 3

// Active operation tracking for concurrency control
const activeOperations = new Set<string>()

/**
 * Validates Excel range addresses and normalizes them
 */
function validateRange(range: string): RangeValidationResult {
  try {
    // Basic range validation regex
    const rangePattern = /^[A-Z]+\d+(?::[A-Z]+\d+)?$/i

    if (!rangePattern.test(range)) {
      return {
        isValid: false,
        error: `Invalid range format: ${range}`
      }
    }

    // Normalize to uppercase
    const normalizedRange = range.toUpperCase()

    // Check for circular references (basic check)
    if (
      normalizedRange.includes('CIRCULAR') ||
      normalizedRange.includes('#REF!')
    ) {
      return {
        isValid: false,
        error: `Circular reference detected in range: ${range}`
      }
    }

    return {
      isValid: true,
      normalizedRange
    }
  } catch (error) {
    return {
      isValid: false,
      error: `Range validation error: ${
        error instanceof Error ? error.message : 'Unknown error'
      }`
    }
  }
}

/**
 * Checks if formula cache is valid and up-to-date
 */
function isCacheValid(dataRowCount: number): boolean {
  if (!formulaCache) return false

  const now = Date.now()
  const cacheAge = now - formulaCache.timestamp

  return (
    cacheAge < FORMULA_CACHE_TTL && formulaCache.dataRowCount === dataRowCount
  )
}

/**
 * Creates or updates the formula cache
 */
function updateFormulaCache(
  formulas: ExcelFormula[],
  dataRowCount: number,
  summaryStartRow: number
): void {
  formulaCache = {
    formulas: [...formulas],
    timestamp: Date.now(),
    dataRowCount,
    summaryStartRow,
    processedFormulas: new Map()
  }
}

/**
 * Clears the formula cache
 */
function clearFormulaCache(): void {
  formulaCache = null
}

/**
 * Processes formula placeholders with actual values
 */
function processFormulaPlaceholders(
  formula: string,
  dataRowCount: number
): string {
  return formula.replace(/\{lastRow\}/g, dataRowCount.toString())
}

/**
 * Checks for concurrent operations to prevent conflicts
 */
function checkConcurrentOperations(operationId: string): boolean {
  if (activeOperations.has(operationId)) {
    return false
  }

  if (activeOperations.size >= MAX_CONCURRENT_OPERATIONS) {
    return false
  }

  activeOperations.add(operationId)
  return true
}

/**
 * Removes operation from active tracking
 */
function releaseOperation(operationId: string): void {
  activeOperations.delete(operationId)
}

/**
 * Finds existing summary section in worksheet
 */
async function findSummarySection(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  dataRowCount: number
): Promise<{ found: boolean; startRow: number }> {
  try {
    const usedRange = sheet.getUsedRange()
    usedRange.load(['rowCount', 'values'])
    await context.sync()

    // Search for existing "Summary" header starting from after the data table
    for (let i = dataRowCount + 1; i <= usedRange.rowCount; i++) {
      try {
        const cellRange = sheet.getRange(`A${i}`)
        cellRange.load('values')
        await context.sync()

        if (cellRange.values[0][0] === 'Summary') {
          return { found: true, startRow: i }
        }
      } catch (error) {
        // Continue searching if cell doesn't exist
        continue
      }
    }

    return { found: false, startRow: dataRowCount + 2 }
  } catch (error) {
    console.warn('Error finding summary section:', error)
    return { found: false, startRow: dataRowCount + 2 }
  }
}

/**
 * Batch processes formulas for better performance
 */
async function batchProcessFormulas(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  formulas: ExcelFormula[],
  summaryStartRow: number,
  dataRowCount: number,
  options: FormulaProcessingOptions
): Promise<void> {
  const batchSize = options.batchSize || DEFAULT_BATCH_SIZE
  const chunks: ExcelFormula[][] = []

  // Split formulas into chunks
  for (let i = 0; i < formulas.length; i += batchSize) {
    chunks.push(formulas.slice(i, i + batchSize))
  }

  // Process each chunk
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
    const chunk = chunks[chunkIndex]
    const startRowForChunk = summaryStartRow + 1 + chunkIndex * batchSize

    await timeAsync(
      `processFormulaChunk-${chunkIndex + 1}-of-${chunks.length}`,
      async () => {
        await withExponentialBackoff(async () => {
          // Prepare batch data for this chunk
          const labelValues: string[][] = []
          const formulaValues: string[][] = []

          for (let i = 0; i < chunk.length; i++) {
            const formula = chunk[i]
            const processedFormula = processFormulaPlaceholders(
              formula.formula,
              dataRowCount
            )

            labelValues.push([formula.name])
            formulaValues.push([processedFormula])
          }

          // Set labels in batch
          if (labelValues.length > 0) {
            const labelRange = sheet.getRange(
              `A${startRowForChunk}:A${
                startRowForChunk + labelValues.length - 1
              }`
            )
            labelRange.values = labelValues
          }

          // Set formulas in batch
          if (formulaValues.length > 0) {
            const formulaRange = sheet.getRange(
              `C${startRowForChunk}:C${
                startRowForChunk + formulaValues.length - 1
              }`
            )
            formulaRange.formulas = formulaValues
          }

          // Sync once per chunk
          await context.sync()

          // Load calculated values in batch
          if (formulaValues.length > 0) {
            const calculatedRange = sheet.getRange(
              `C${startRowForChunk}:C${
                startRowForChunk + formulaValues.length - 1
              }`
            )
            calculatedRange.load('values')
            await context.sync()

            // Replace formulas with calculated values in batch
            calculatedRange.values = calculatedRange.values
          }
        }, options.maxRetries || 3)
      },
      { chunkSize: chunk.length }
    )

    // Update progress if callback provided
    if (options.onProgress) {
      const completed = Math.min((chunkIndex + 1) * batchSize, formulas.length)
      options.onProgress({
        total: formulas.length,
        completed,
        current: `Processing chunk ${chunkIndex + 1} of ${chunks.length}`,
        errors: []
      })
    }
  }
}

/**
 * Lazy loading implementation for large formula datasets
 */
async function applyFormulasLazy(
  formulas: ExcelFormula[],
  options: FormulaProcessingOptions = {}
): Promise<void> {
  const LAZY_LOAD_THRESHOLD = 20
  const LAZY_LOAD_CHUNK_SIZE = 5

  if (formulas.length <= LAZY_LOAD_THRESHOLD) {
    // Use regular batch processing for smaller datasets
    return applyFormulas({
      ...options,
      batchSize: options.batchSize || DEFAULT_BATCH_SIZE
    })
  }

  // For large datasets, use progressive loading with user feedback
  updateStatus(
    `Large dataset detected (${formulas.length} formulas). Using progressive loading...`
  )

  const chunks: ExcelFormula[][] = []
  for (let i = 0; i < formulas.length; i += LAZY_LOAD_CHUNK_SIZE) {
    chunks.push(formulas.slice(i, i + LAZY_LOAD_CHUNK_SIZE))
  }

  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i]
    updateStatus(
      `Processing chunk ${i + 1}/${chunks.length} (${chunk.length} formulas)...`
    )

    // Process chunk with delay to prevent UI blocking
    await new Promise(resolve => setTimeout(resolve, 100))

    // Apply formulas for this chunk
    await applyFormulas({
      ...options,
      batchSize: chunk.length,
      onProgress: progress => {
        const overallProgress =
          (i * LAZY_LOAD_CHUNK_SIZE + progress.completed) / formulas.length
        updateStatus(`Overall progress: ${Math.round(overallProgress * 100)}%`)
        if (options.onProgress) {
          options.onProgress({
            ...progress,
            total: formulas.length,
            completed: i * LAZY_LOAD_CHUNK_SIZE + progress.completed
          })
        }
      }
    })
  }
}

/**
 * Memory management utility to dispose of Excel objects properly
 */
function disposeExcelObjects(
  ...objects: (Excel.Range | Excel.Worksheet | any)[]
): void {
  try {
    objects.forEach(obj => {
      if (obj && typeof obj.untrack === 'function') {
        obj.untrack()
      }
    })
  } catch (error) {
    console.warn('Error disposing Excel objects:', error)
  }
}

/**
 * Advanced error handling with retry logic and fallback strategies
 */
async function executeWithFallback<T>(
  primaryOperation: () => Promise<T>,
  fallbackOperation: () => Promise<T>,
  operationName: string,
  maxRetries: number = 3
): Promise<T> {
  try {
    return await withExponentialBackoff(primaryOperation, maxRetries)
  } catch (primaryError) {
    console.warn(
      `Primary operation '${operationName}' failed, attempting fallback:`,
      primaryError
    )

    try {
      return await withExponentialBackoff(
        fallbackOperation,
        Math.max(1, maxRetries - 1)
      )
    } catch (fallbackError) {
      console.error(
        `Both primary and fallback operations failed for '${operationName}':`,
        {
          primaryError,
          fallbackError
        }
      )
      throw new Error(
        `Operation '${operationName}' failed: ${primaryError.message}. Fallback also failed: ${fallbackError.message}`
      )
    }
  }
}

/**
 * Detects and handles circular references in formulas
 */
function detectCircularReferences(formulas: ExcelFormula[]): string[] {
  const circularRefs: string[] = []
  const formulaMap = new Map<string, string>()

  // Build formula map
  formulas.forEach(f => {
    formulaMap.set(f.name, f.formula)
  })

  // Check for circular references
  formulas.forEach(formula => {
    const visited = new Set<string>()
    const stack = new Set<string>()

    function hasCircularRef(name: string): boolean {
      if (stack.has(name)) {
        return true
      }
      if (visited.has(name)) {
        return false
      }

      visited.add(name)
      stack.add(name)

      const formulaText = formulaMap.get(name)
      if (formulaText) {
        // Simple check for references to other formulas
        formulas.forEach(otherFormula => {
          if (
            otherFormula.name !== name &&
            formulaText.includes(otherFormula.name)
          ) {
            if (hasCircularRef(otherFormula.name)) {
              return true
            }
          }
        })
      }

      stack.delete(name)
      return false
    }

    if (hasCircularRef(formula.name)) {
      circularRefs.push(formula.name)
    }
  })

  return circularRefs
}

/**
 * Validates worksheet state before applying formulas
 */
async function validateWorksheetState(context: Excel.RequestContext): Promise<{
  isValid: boolean
  errors: string[]
  warnings: string[]
}> {
  const errors: string[] = []
  const warnings: string[] = []

  try {
    const workbook = context.workbook
    const sheet = workbook.worksheets.getActiveWorksheet()

    // Check if worksheet is protected
    sheet.load('protection')
    await context.sync()

    if (sheet.protection.protected) {
      errors.push('Worksheet is protected and cannot be modified')
    }

    // Check for calculation mode (using application instead of workbook)
    try {
      const application = context.application
      application.load('calculationMode')
      await context.sync()

      if (application.calculationMode !== Excel.CalculationMode.automatic) {
        warnings.push(
          'Application calculation mode is not automatic. Formulas may not update immediately.'
        )
      }
    } catch (error) {
      // Calculation mode check is not critical, just log warning
      console.warn('Could not check calculation mode:', error)
    }

    // Check used range size
    const usedRange = sheet.getUsedRange()
    if (usedRange) {
      usedRange.load(['rowCount', 'columnCount'])
      await context.sync()

      const totalCells = usedRange.rowCount * usedRange.columnCount
      if (totalCells > 100000) {
        warnings.push(
          `Large worksheet detected (${totalCells} cells). Performance may be impacted.`
        )
      }
    }
  } catch (error) {
    errors.push(
      `Worksheet validation error: ${
        error instanceof Error ? error.message : 'Unknown error'
      }`
    )
  }

  return {
    isValid: errors.length === 0,
    errors,
    warnings
  }
}

/**
 * Enhanced applyFormulas with all optimizations
 */
async function applyFormulasEnhanced(
  options: FormulaProcessingOptions = {}
): Promise<void> {
  const operationId = `applyFormulasEnhanced_${Date.now()}`

  // Check for concurrent operations
  if (!checkConcurrentOperations(operationId)) {
    updateStatus('Formula operation already in progress. Please wait...', true)
    return
  }

  const timerId = startTiming('applyFormulasEnhanced', options)

  try {
    // Pre-validation
    await Excel.run(async context => {
      const validation = await validateWorksheetState(context)

      if (!validation.isValid) {
        throw new Error(
          `Worksheet validation failed: ${validation.errors.join(', ')}`
        )
      }

      if (validation.warnings.length > 0) {
        console.warn('Worksheet warnings:', validation.warnings)
        updateStatus(`Warnings: ${validation.warnings.join(', ')}`)
      }
    })

    // Use enhanced error handling with fallback
    await executeWithFallback(
      // Primary operation: optimized batch processing
      async () => {
        await applyFormulas(options)
      },
      // Fallback operation: legacy single-formula processing
      async () => {
        updateStatus('Using fallback processing method...')
        await applyFormulasLegacy()
      },
      'applyFormulasEnhanced',
      options.maxRetries || 3
    )
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Enhanced formula application failed: ${errorMessage}`, true)
    console.error('Enhanced formula application error:', error)
    throw error
  } finally {
    releaseOperation(operationId)
    endTiming(timerId)
  }
}

// Initialize Office.js
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Register event handlers
    document.getElementById('fetchData')?.addEventListener('click', syncData)
    document.getElementById('fetchData2')?.addEventListener('click', syncData2)

    // Initial data load
    syncData2()
    syncData()
  }
})

/**
 * Optimized function to apply formulas from the API to the worksheet
 * Features: batch operations, caching, progress tracking, error handling, concurrency control
 */
async function applyFormulas(
  options: FormulaProcessingOptions = {}
): Promise<void> {
  const operationId = `applyFormulas_${Date.now()}`

  // Check for concurrent operations
  if (!checkConcurrentOperations(operationId)) {
    updateStatus('Formula operation already in progress. Please wait...', true)
    return
  }

  const timerId = startTiming('applyFormulas', {
    enableCache: options.enableCache !== false,
    batchSize: options.batchSize || DEFAULT_BATCH_SIZE
  })

  try {
    // Calculate where the data table ends (products + header row)
    const dataRowCount = data.products ? data.products.length + 1 : 1 // +1 for header

    // Check cache first if enabled
    if (options.enableCache !== false && isCacheValid(dataRowCount)) {
      updateStatus('Using cached formulas for better performance...')

      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()

        // Use cached summary start row and process cached formulas
        await batchProcessFormulas(
          context,
          sheet,
          formulaCache!.formulas,
          formulaCache!.summaryStartRow,
          dataRowCount,
          options
        )

        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns()
        await context.sync()
      })

      updateStatus(`Applied ${formulaCache!.formulas.length} cached formulas`)
      return
    }

    updateStatus('Fetching formulas from API...')

    // Fetch formulas from the API with retry logic
    const formulasData = await timeAsync('fetchFormulas', async () => {
      return await withExponentialBackoff(async () => {
        const response = await fetch(`${API_BASE_URL}/formulas`)

        if (!response.ok) {
          throw new Error(
            `API error: ${response.status} ${response.statusText}`
          )
        }

        return (await response.json()) as FormulasResponse
      }, options.maxRetries || 3)
    })

    if (!formulasData.formulas || formulasData.formulas.length === 0) {
      updateStatus('No formulas available from API')
      return
    }

    // Validate formulas before processing
    const validationErrors: string[] = []
    for (const formula of formulasData.formulas) {
      if (!formula.formula || !formula.name) {
        validationErrors.push(`Invalid formula: ${formula.name || 'unnamed'}`)
        continue
      }

      // Basic formula validation
      const processedFormula = processFormulaPlaceholders(
        formula.formula,
        dataRowCount
      )
      if (options.validateRanges !== false) {
        // Extract range references and validate them
        const rangeMatches = processedFormula.match(
          /[A-Z]+\d+(?::[A-Z]+\d+)?/gi
        )
        if (rangeMatches) {
          for (const range of rangeMatches) {
            const validation = validateRange(range)
            if (!validation.isValid) {
              validationErrors.push(
                `Formula '${formula.name}': ${validation.error}`
              )
            }
          }
        }
      }
    }

    if (validationErrors.length > 0) {
      throw new Error(
        `Formula validation failed:\n${validationErrors.join('\n')}`
      )
    }

    updateStatus(`Processing ${formulasData.formulas.length} formulas...`)

    // Apply the formulas to the worksheet with optimized batch processing
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Find existing summary section efficiently
      const summaryInfo = await timeAsync('findSummarySection', async () => {
        return await findSummarySection(context, sheet, dataRowCount)
      })

      let summaryStartRow = summaryInfo.startRow

      // Clear existing summary section if found
      if (summaryInfo.found) {
        await timeAsync('clearExistingSummary', async () => {
          const summaryRowsToDelete = 1 + formulasData.formulas.length
          const clearRange = sheet.getRange(
            `A${summaryStartRow}:C${summaryStartRow + summaryRowsToDelete - 1}`
          )
          clearRange.clear()
          await context.sync()
        })
      }

      // Add the summary header
      const summaryHeaderRange = sheet.getRange(`A${summaryStartRow}`)
      summaryHeaderRange.values = [['Summary']]
      summaryHeaderRange.format.font.bold = true
      await context.sync()

      // Process formulas in batches for optimal performance
      await timeAsync('batchProcessFormulas', async () => {
        await batchProcessFormulas(
          context,
          sheet,
          formulasData.formulas,
          summaryStartRow,
          dataRowCount,
          {
            ...options,
            onProgress: progress => {
              updateStatus(
                `Processing formulas: ${progress.completed}/${progress.total} (${progress.current})`
              )
              if (options.onProgress) {
                options.onProgress(progress)
              }
            }
          }
        )
      })

      // Auto-fit columns
      await timeAsync('autofitColumns', async () => {
        sheet.getUsedRange().format.autofitColumns()
        await context.sync()
      })

      // Update cache with successful results
      if (options.enableCache !== false) {
        updateFormulaCache(formulasData.formulas, dataRowCount, summaryStartRow)
      }

      updateStatus(
        `Successfully applied ${formulasData.formulas.length} formulas`
      )
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Formula error: ${errorMessage}`, true)
    console.error('Formula application error:', error)

    // Clear cache on error to prevent stale data
    if (options.enableCache !== false) {
      clearFormulaCache()
    }

    throw error
  } finally {
    // Always release the operation lock and end timing
    releaseOperation(operationId)
    endTiming(timerId)
  }
}

/**
 * Legacy wrapper for backward compatibility
 * @deprecated Use applyFormulas(options) instead
 */
async function applyFormulasLegacy(): Promise<void> {
  return applyFormulas({
    enableCache: true,
    batchSize: 1, // Use original behavior of processing one at a time
    validateRanges: false, // Disable validation for backward compatibility
    enableRetry: true,
    maxRetries: 1
  })
}

/**
 * Synchronize data from API to Excel
 */
async function syncData(): Promise<void> {
  try {
    updateStatus('Fetching data from API...')

    console.time('api')

    // Call your API endpoint
    const response = await fetch(`${API_BASE_URL}/data`)

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }

    data = await response.json()

    console.timeEnd('api')

    console.log('API response:', data)

    updateStatus('Writing data to Excel...')

    // Write data to Excel
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Clear previous event handlers to avoid duplicates
      sheet.onChanged.remove(handleCellChange)

      // Create a header row
      const headerRange = sheet.getRange('A1:C1')
      headerRange.values = [['ID', 'Product Name', 'Price']]
      headerRange.format.font.bold = true

      console.log(
        'Writing to Excel:',
        data.products.map(p => [p.id, p.name, p.price])
      )

      // Write data rows
      if (data.products && data.products.length > 0) {
        const dataRange = sheet.getRange(`A2:C${data.products.length + 1}`)
        dataRange.values = data.products.map(product => [
          product.id,
          product.name,
          product.price
        ])

        // Format price column as currency
        const priceRange = sheet.getRange(`C2:C${data.products.length + 1}`)
        priceRange.numberFormat = [['$#,##0.00']]
      }

      // Format table for readability
      sheet.getUsedRange().format.autofitColumns()

      // Register a cell change event handler
      sheet.onChanged.add(handleCellChange)

      await context.sync()

      updateStatus(`Data synchronized! Now applying formulas...`)
      // After syncing data, apply formulas from API with optimized settings
      await applyFormulas({
        enableCache: true,
        batchSize: 5,
        enableProgressTracking: true,
        validateRanges: true,
        enableRetry: true,
        maxRetries: 3,
        onProgress: progress => {
          updateStatus(
            `Applying formulas: ${progress.completed}/${progress.total}`
          )
        }
      })

      updateStatus(
        `Data synchronized successfully! (${data.products.length} products)`
      )
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Error: ${errorMessage}`, true)
    console.error('Sync error:', error)
  }
}

/**
 * Synchronize complex data from API to Excel
 */

/**
 * Handle cell changes and send updates to API
 */
async function handleCellChange(
  event: Excel.WorksheetChangedEventArgs
): Promise<void> {
  await Excel.run(async context => {
    try {
      // Get details about the changed range
      const changedRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(event.address)
      changedRange.load(['values', 'rowIndex', 'columnIndex'])

      await context.sync()

      // Skip if not a single cell change
      if (
        changedRange.values?.length !== 1 ||
        changedRange.values[0].length !== 1
      ) {
        return
      }

      const rowIndex = changedRange.rowIndex
      const colIndex = changedRange.columnIndex
      const newValue = changedRange.values[0][0]

      // Skip header row
      if (rowIndex === 0) {
        return
      }

      // Get the used range to determine where the data table ends
      const usedRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getUsedRange()
      usedRange.load('rowCount')
      await context.sync()

      // Skip if the cell is in the formula/summary section (below the data table)
      // First, find where our data products end
      const dataRowCount = data.products ? data.products.length + 1 : 0 // +1 for header

      // If we're in a row beyond the data table, skip processing
      // We compare rowIndex to dataRowCount-1 because rowIndex is 0-based (0 is header, 1 is first data row)
      if (rowIndex > dataRowCount - 1) {
        console.log(`Skipping change in formula section - row ${rowIndex + 1}`) // +1 for display
        return
      }

      // Continue with normal processing for data table cells

      // Skip ID column changes (column A)
      if (colIndex === 0) {
        updateStatus('Product ID cannot be modified directly', true)
        return
      }

      // Skip if not in our data columns (A-C)
      if (colIndex > 2) {
        return
      }

      // Get the product ID from column A
      const idCell = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(`A${rowIndex + 1}`)
      idCell.load('values')
      await context.sync()

      const productIdValue = idCell.values[0][0]

      // Safely convert to number, validating we have a proper ID
      let productId: number

      if (typeof productIdValue === 'number') {
        productId = productIdValue
      } else if (
        typeof productIdValue === 'string' &&
        !isNaN(parseInt(productIdValue))
      ) {
        productId = parseInt(productIdValue)
      } else {
        console.log(`Invalid product ID: ${productIdValue}`)
        updateStatus('Cannot identify product ID', true)
        return
      }

      // Determine which field was updated
      let field: string
      switch (colIndex) {
        case 1:
          field = 'name'
          break
        case 2:
          field = 'price'
          break
        default:
          return
      }

      // Prepare update data for API
      const updateData: CellUpdateRequest = {
        id: productId,
        field: field,
        value: field === 'price' ? parseFloat(newValue) : newValue
      }

      // Send update to API
      updateStatus(`Sending ${field} update to API...`)
      await sendUpdateToApi(updateData)
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error'
      updateStatus(`Cell change error: ${errorMessage}`, true)
      console.error('Cell change handling error:', error)
    }
  })
}

/**
 * Send cell update to API
 */
async function sendUpdateToApi(updateData: CellUpdateRequest): Promise<void> {
  try {
    const response = await fetch(`${API_BASE_URL}/update-cell`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updateData)
    })

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }

    const result: CellUpdateResponse = await response.json()

    if (result.success) {
      updateStatus(`Updated ${updateData.field} successfully!`)
    } else {
      updateStatus(`Update failed: ${result.error || 'Unknown error'}`, true)
    }
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`API Error: ${errorMessage}`, true)
    console.error('API update error:', error)
  }
}
