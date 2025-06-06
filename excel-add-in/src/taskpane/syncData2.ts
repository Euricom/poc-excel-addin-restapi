import { handleComplexCellChange } from './handleComplexCellChange'
import { updateStatus } from '../utils/taskpaneUtils'
import { ComplexApiResponse } from '../types/complexProduct'
import {
  COMPLEX_PRODUCT_HEADERS,
  productsToRows
} from '../utils/complexProductUtils'
import { timeAsync, logPerformanceReport } from '../utils/performanceUtils'
import {
  processRangeInChunks,
  processItemsInChunks,
  ChunkingProgress,
  CHUNKING_CONFIG
} from '../utils/chunkingUtils'

const API_BASE_URL = 'http://localhost:3001/api'

// Cache for complex data to avoid unnecessary API calls
let complexDataCache: ComplexApiResponse = { products: [] }
let lastFetchTime = 0
const CACHE_TTL = 60000 // 1 minute cache TTL

/**
 * Synchronizes complex product data from the API to the active Excel worksheet.
 * Optimized version with:
 * - Data caching
 * - Batched Excel operations
 * - Reduced context.sync() calls
 * - Progressive loading for large datasets
 * - Optimized formatting
 */
export async function syncData2(): Promise<void> {
  try {
    updateStatus('Preparing to sync complex data...')

    // Use performance monitoring for the entire operation
    await timeAsync('syncData2-total', async () => {
      // Fetch data with caching
      await timeAsync('fetchComplexData', fetchComplexDataWithCaching, {
        cacheSize: complexDataCache.products.length
      })

      // Process data in Excel
      await timeAsync('writeDataToExcel', writeDataToExcel, {
        rowCount: complexDataCache.products.length
      })
    })

    // Log performance metrics to console
    logPerformanceReport()

    updateStatus(
      `Complex data synchronized successfully! (${complexDataCache.products.length} products)`
    )
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Error: ${errorMessage}`, true)
    console.error('Complex data sync error:', error)
  }
}

/**
 * Fetches complex data from API with caching to reduce unnecessary API calls
 */
async function fetchComplexDataWithCaching(): Promise<void> {
  const now = Date.now()

  // Use cached data if it's still fresh
  if (complexDataCache.products.length > 0 && now - lastFetchTime < CACHE_TTL) {
    console.log('Using cached complex data')
    updateStatus('Using cached complex data...')
    return
  }

  updateStatus('Fetching complex data from API...')

  try {
    const response = await fetch(`${API_BASE_URL}/data2`)
    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }

    complexDataCache = await response.json()
    lastFetchTime = now

    console.log('API response (complex data):', complexDataCache)
  } catch (error) {
    console.error('Error fetching complex data:', error)
    throw error
  }
}

/**
 * Writes data to Excel with optimized batching and formatting
 * Uses chunking strategy for large datasets to prevent 5MB payload limit
 * Writes to a second sheet, creating it if it doesn't exist
 */
async function writeDataToExcel(): Promise<void> {
  if (!complexDataCache.products || complexDataCache.products.length === 0) {
    updateStatus('No products to synchronize')
    return
  }

  updateStatus('Writing complex data to Excel...')

  await Excel.run(async context => {
    // Get or create a second worksheet
    let sheet: Excel.Worksheet
    const worksheets = context.workbook.worksheets
    worksheets.load('items')
    await context.sync()

    // Check if there's at least a second sheet
    if (worksheets.items.length > 1) {
      // Use the second sheet
      sheet = worksheets.items[1]
    } else {
      // Create a new sheet named "Data Sync 2"
      sheet = worksheets.add('Data Sync 2')
      updateStatus('Created new worksheet "Data Sync 2"')
    }

    // Remove previous event handler to avoid duplicates
    sheet.onChanged.remove(handleComplexCellChange)

    // Prepare data for batch operations
    const products = complexDataCache.products
    const rowCount = products.length

    // Convert all products to rows in a single optimized batch operation
    const dataConversionStart = performance.now()
    const allRows = [COMPLEX_PRODUCT_HEADERS]
    const productRows = productsToRows(products)
    allRows.push(...productRows)
    console.log(
      `Data conversion took ${(performance.now() - dataConversionStart).toFixed(
        2
      )}ms for ${rowCount} rows`
    )

    // For small datasets (under threshold), use the original approach
    if (rowCount <= 100) {
      // Write all data at once (headers + data)
      const writeStart = performance.now()
      const dataRange = sheet.getRange(`A1:AX${rowCount + 1}`)
      dataRange.values = allRows
      console.log(
        `Range values assignment took ${(
          performance.now() - writeStart
        ).toFixed(2)}ms`
      )
    } else {
      // For larger datasets, use chunking strategy
      updateStatus(`Using chunking strategy for ${rowCount} rows...`)

      // Write header row first
      const headerRange = sheet.getRange('A1:AX1')
      headerRange.values = [COMPLEX_PRODUCT_HEADERS]
      await context.sync()

      // Write data rows using chunking
      await processRangeInChunks(
        context,
        sheet,
        productRows,
        'A2', // Start at row 2 (after headers)
        {
          operationId: `sync_data_${Date.now()}`,
          onProgress: progress => {
            const percent = Math.round(
              (progress.completedChunks / progress.totalChunks) * 100
            )
            updateStatus(`Writing data: ${percent}% complete`)
          }
        }
      )
    }

    // Apply all formatting in a single batch
    await timeAsync(
      'applyFormatting',
      async () => {
        await applyFormattingInBatch(context, sheet, rowCount)
      },
      { rowCount }
    )

    // Autofit columns for readability (can be expensive for large datasets)
    // Only autofit visible columns for better performance
    const autofitStart = performance.now()
    sheet.getRange('A:AX').format.autofitColumns()
    console.log(
      `Autofit columns assignment took ${(
        performance.now() - autofitStart
      ).toFixed(2)}ms`
    )

    // Register the cell change event handler
    sheet.onChanged.add(handleComplexCellChange)

    // Execute all operations in a single batch
    const syncStart = performance.now()
    await context.sync()
    console.log(
      `Final context.sync() took ${(performance.now() - syncStart).toFixed(
        2
      )}ms`
    )
  })
}

/**
 * Applies all formatting in a single batch operation
 */
async function applyFormattingInBatch(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowCount: number
): Promise<void> {
  // Format header row
  const headerRange = sheet.getRange('A1:AX1')
  headerRange.format.font.bold = true

  // Only proceed with data formatting if we have rows
  if (rowCount === 0) return

  // Group formatting by type to reduce API calls

  // Currency formatting
  sheet.getRange(`D2:D${rowCount + 1}`).numberFormat = [['$#,##0.00']]

  // Date formatting - group all date columns
  const dateRanges = [
    sheet.getRange(`H2:H${rowCount + 1}`), // Date Added
    sheet.getRange(`AX2:AX${rowCount + 1}`) // Last Updated
  ]
  dateRanges.forEach(range => {
    range.numberFormat = [['yyyy-mm-dd']]
  })

  // Boolean columns - format all in one operation
  const booleanColumns = [
    'G',
    'AA',
    'AB',
    'AC',
    'AD',
    'AE',
    'AF',
    'AG',
    'AH',
    'AI',
    'AJ',
    'AK',
    'AL',
    'AM',
    'AN',
    'AO',
    'AP',
    'AQ',
    'AR'
  ]

  // Format each boolean column individually
  booleanColumns.forEach(col => {
    sheet.getRange(`${col}2:${col}${rowCount + 1}`).numberFormat = [['@']]
  })

  // Decimal number formatting - group similar formats
  const decimalRanges = [
    sheet.getRange(`L2:L${rowCount + 1}`), // Weight
    sheet.getRange(`Y2:Y${rowCount + 1}`), // Shipping Weight
    sheet.getRange(`AW2:AW${rowCount + 1}`) // Power Consumption
  ]
  decimalRanges.forEach(range => {
    range.numberFormat = [['0.00']]
  })

  // Percentage formatting
  const percentRanges = [
    sheet.getRange(`W2:W${rowCount + 1}`), // Discount Percentage
    sheet.getRange(`X2:X${rowCount + 1}`) // Tax Rate
  ]
  percentRanges.forEach(range => {
    range.numberFormat = [['0.00%']]
  })

  // Special formatting for Power Consumption
  sheet.getRange(`AW2:AW${rowCount + 1}`).numberFormat = [['0.00 W']]

  // Popularity rating
  sheet.getRange(`AX2:AX${rowCount + 1}`).numberFormat = [['0.00']]
}

/**
 * Optimized function to handle large datasets by loading them progressively
 * with advanced chunking, retry logic, and progress tracking.
 * This can be used for very large datasets instead of the standard syncData2
 * Writes to a second sheet, creating it if it doesn't exist
 */
export async function syncLargeDataset(
  options: {
    operationId?: string
    chunkSize?: number
    resumeFromChunk?: number
  } = {}
): Promise<void> {
  try {
    updateStatus('Preparing to sync large dataset...')

    // Generate a unique operation ID if not provided
    const operationId =
      options.operationId || `sync_large_dataset_${Date.now()}`

    // Fetch data
    await fetchComplexDataWithCaching()

    const products = complexDataCache.products
    const totalProducts = products.length

    if (totalProducts === 0) {
      updateStatus('No products to synchronize')
      return
    }

    // Determine optimal chunk size if not specified
    const chunkSize =
      options.chunkSize ||
      Math.min(
        Math.max(
          Math.floor(
            CHUNKING_CONFIG.MAX_PAYLOAD_SIZE_BYTES /
              (COMPLEX_PRODUCT_HEADERS.length *
                CHUNKING_CONFIG.ESTIMATED_BYTES_PER_CELL)
          ),
          CHUNKING_CONFIG.MIN_CHUNK_SIZE
        ),
        CHUNKING_CONFIG.MAX_CHUNK_SIZE
      )

    updateStatus(
      `Syncing ${totalProducts} products with chunk size: ${chunkSize}`
    )

    await Excel.run(async context => {
      // Get or create a second worksheet
      let sheet: Excel.Worksheet
      const worksheets = context.workbook.worksheets
      worksheets.load('items')
      await context.sync()

      // Check if there's at least a second sheet
      if (worksheets.items.length > 1) {
        // Use the second sheet
        sheet = worksheets.items[1]
      } else {
        // Create a new sheet named "Data Sync 2"
        sheet = worksheets.add('Data Sync 2')
        updateStatus('Created new worksheet "Data Sync 2"')
      }

      // Remove previous event handler
      sheet.onChanged.remove(handleComplexCellChange)

      // Write headers
      const headerRange = sheet.getRange('A1:AX1')
      headerRange.values = [COMPLEX_PRODUCT_HEADERS]
      headerRange.format.font.bold = true
      await context.sync()

      // Progress tracking callback
      const onProgress = (progress: ChunkingProgress) => {
        const percent = Math.round(
          (progress.completedChunks / progress.totalChunks) * 100
        )
        updateStatus(
          `Processing products: ${percent}% complete (chunk ${progress.completedChunks}/${progress.totalChunks})`
        )
      }

      // Process products in chunks
      await processItemsInChunks(
        async (productChunk, chunkIndex, totalChunks) => {
          // Convert products to rows
          const rowsChunk = productsToRows(productChunk)

          // Calculate starting row (accounting for header row)
          const startRow = chunkIndex * chunkSize + 2 // +2 because row 1 is header

          // Write data for this chunk
          const dataRange = `A${startRow}:AX${startRow + rowsChunk.length - 1}`

          // Use processRangeInChunks for very large chunks
          if (rowsChunk.length > 100) {
            await processRangeInChunks(
              context,
              sheet,
              rowsChunk,
              `A${startRow}`,
              { operationId: `${operationId}_subchunk_${chunkIndex}` }
            )
          } else {
            // For smaller chunks, write directly
            const range = sheet.getRange(dataRange)
            range.values = rowsChunk
            await context.sync()
          }

          // Apply minimal formatting for this chunk
          sheet.getRange(
            `D${startRow}:D${startRow + rowsChunk.length - 1}`
          ).numberFormat = [['$#,##0.00']]
          await context.sync()
        },
        products,
        {
          chunkSize,
          operationId,
          onProgress,
          resumeFromChunk: options.resumeFromChunk || 0
        }
      )

      // Apply full formatting after all data is written
      await applyFormattingInBatch(context, sheet, totalProducts)

      // Register event handler
      sheet.onChanged.add(handleComplexCellChange)

      await context.sync()

      updateStatus(
        `Large dataset synchronized successfully! (${totalProducts} products)`
      )
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Error: ${errorMessage}`, true)
    console.error('Large dataset sync error:', error)
  }
}

/**
 * Resumes a previously interrupted large dataset sync operation
 * @param operationId The ID of the operation to resume
 */
export async function resumeLargeDatasetSync(
  operationId: string
): Promise<void> {
  try {
    // Load progress from session storage
    const progress = sessionStorage.getItem(
      `${CHUNKING_CONFIG.PROGRESS_KEY_PREFIX}${operationId}`
    )

    if (!progress) {
      updateStatus(
        `No saved progress found for operation: ${operationId}`,
        true
      )
      return
    }

    const progressData = JSON.parse(progress) as ChunkingProgress

    if (progressData.status === 'completed') {
      updateStatus(`Operation ${operationId} was already completed`)
      return
    }

    updateStatus(
      `Resuming operation from chunk ${progressData.completedChunks + 1}/${
        progressData.totalChunks
      }`
    )

    // Resume the operation
    await syncLargeDataset({
      operationId,
      resumeFromChunk: progressData.completedChunks
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Resume error: ${errorMessage}`, true)
    console.error('Resume operation error:', error)
  }
}
