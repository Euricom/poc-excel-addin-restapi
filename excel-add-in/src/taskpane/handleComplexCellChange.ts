import { updateStatus } from '../utils/taskpaneUtils'
import {
  COMPLEX_PRODUCT_FIELD_MAP,
  getFieldNameByColumnIndex,
  parseBoolean,
  parseNumber,
  parseStringArray,
  parseDate
} from '../utils/complexProductUtils'
import { ComplexProduct } from '../types/complexProduct'
import { startTiming, endTiming, timeAsync } from '../utils/performanceUtils'
import { withExponentialBackoff } from '../utils/chunkingUtils'

const API_BASE_URL = 'http://localhost:3001/api'

// Cache for pending updates to batch API calls
const pendingUpdates: Array<{
  id: number
  field: string
  value: any
  timestamp: number
}> = []
const UPDATE_BATCH_DELAY = 500 // ms to wait before sending batch updates
let updateTimeoutId: number | null = null

// Field type mappings for efficient validation
const numericFields: (keyof ComplexProduct)[] = [
  'price',
  'quantity',
  'rating',
  'weight',
  'minOrderQuantity',
  'maxOrderQuantity',
  'discountPercentage',
  'taxRate',
  'shippingWeight',
  'powerConsumption',
  'popularity'
]

const booleanFields: (keyof ComplexProduct)[] = [
  'inStock',
  'assemblyRequired',
  'batteryRequired',
  'batteriesIncluded',
  'waterproof',
  'heatResistant',
  'coldResistant',
  'uvResistant',
  'windResistant',
  'shockResistant',
  'dustResistant',
  'scratchResistant',
  'stainResistant',
  'fadeResistant',
  'rustResistant',
  'moldResistant',
  'fireResistant',
  'recyclable',
  'biodegradable'
]

const dateFields: (keyof ComplexProduct)[] = ['dateAdded', 'lastUpdated']
const arrayFields: (keyof ComplexProduct)[] = ['tags', 'certifications']

/**
 * Optimized handler for cell changes in the complex data table.
 * - Validates and converts the changed value.
 * - Batches updates to reduce API calls.
 * - Uses pre-computed field type mappings.
 * - Minimizes Excel API calls.
 */
export async function handleComplexCellChange(
  event: Excel.WorksheetChangedEventArgs
): Promise<void> {
  const overallTimerId = startTiming('handleComplexCellChange', {
    address: event.address
  })

  await Excel.run(async context => {
    try {
      // Load all required properties in a single operation
      const loadTimerId = startTiming('load-cell-data')
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const changedRange = sheet.getRange(event.address)
      changedRange.load(['values', 'rowIndex', 'columnIndex'])

      // Get the used range to determine data boundaries
      const usedRange = sheet.getUsedRange()
      usedRange.load('rowCount')

      // Execute all load operations in a single batch
      await context.sync()
      endTiming(loadTimerId)

      // Skip processing for non-single cell changes
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
      if (rowIndex === 0) return

      const dataRowCount = usedRange.rowCount

      // Skip if beyond data table
      if (rowIndex > dataRowCount - 1) return

      // Skip ID column edits
      if (colIndex === 0) {
        updateStatus('Product ID cannot be modified directly', true)
        return
      }

      // Skip undefined columns
      if (colIndex < 0 || colIndex >= COMPLEX_PRODUCT_FIELD_MAP.length) return

      // Get the product ID efficiently
      const idCell = sheet.getRange(`A${rowIndex + 1}`)
      idCell.load('values')
      await context.sync()

      const productIdValue = idCell.values[0][0]
      let productId: number

      if (typeof productIdValue === 'number') {
        productId = productIdValue
      } else if (
        typeof productIdValue === 'string' &&
        !isNaN(parseInt(productIdValue))
      ) {
        productId = parseInt(productIdValue)
      } else {
        updateStatus('Cannot identify product ID', true)
        return
      }

      // Get field name
      const field = getFieldNameByColumnIndex(colIndex)
      if (!field) return

      // Validate and convert value using pre-computed field type mappings
      let value: any = newValue

      if (numericFields.includes(field as keyof ComplexProduct)) {
        value = parseNumber(newValue)
        if (value === undefined) {
          updateStatus(`Invalid number for ${field}`, true)
          return
        }
      } else if (booleanFields.includes(field as keyof ComplexProduct)) {
        value = parseBoolean(newValue)
        if (value === undefined) {
          updateStatus(`Invalid boolean for ${field}`, true)
          return
        }
      } else if (arrayFields.includes(field as keyof ComplexProduct)) {
        value = parseStringArray(newValue)
      } else if (dateFields.includes(field as keyof ComplexProduct)) {
        value = parseDate(newValue)
        if (value === undefined) {
          updateStatus(`Invalid date for ${field}`, true)
          return
        }
      }

      // Add to pending updates for batching
      queueUpdate(productId, field, value)

      updateStatus(`Change queued: ${field} (will be sent shortly)`)
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error'
      updateStatus(`Cell change error: ${errorMessage}`, true)
      console.error('Complex cell change handling error:', error)
    } finally {
      endTiming(overallTimerId)
    }
  })
}

/**
 * Queues an update to be sent in a batch
 */
function queueUpdate(id: number, field: string, value: any): void {
  // Add to pending updates
  pendingUpdates.push({
    id,
    field,
    value,
    timestamp: Date.now()
  })

  // Clear existing timeout if any
  if (updateTimeoutId !== null) {
    clearTimeout(updateTimeoutId)
  }

  // Set new timeout to process batch
  updateTimeoutId = setTimeout(
    processPendingUpdates,
    UPDATE_BATCH_DELAY
  ) as unknown as number
}

/**
 * Processes all pending updates in a single batch API call
 * Uses chunking strategy for large batches and implements exponential backoff for retries
 */
async function processPendingUpdates(): Promise<void> {
  if (pendingUpdates.length === 0) return

  const batchTimerId = startTiming('processPendingUpdates', {
    updateCount: pendingUpdates.length
  })

  updateStatus(`Sending ${pendingUpdates.length} updates to API...`)

  try {
    // For single updates, use the original API endpoint with retry logic
    if (pendingUpdates.length === 1) {
      const update = pendingUpdates[0]

      await timeAsync('singleUpdate', async () => {
        await withExponentialBackoff(async () => {
          const response = await fetch(`${API_BASE_URL}/update-cell2`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              id: update.id,
              field: update.field,
              value: update.value
            })
          })

          if (!response.ok) {
            throw new Error(
              `API error: ${response.status} ${response.statusText}`
            )
          }

          const result = await response.json()
          if (result.success) {
            updateStatus(`Updated ${update.field} successfully!`)
          } else {
            throw new Error(`Update failed: ${result.error || 'Unknown error'}`)
          }
        })
      })
    }
    // For multiple updates, use chunking strategy and batch endpoint
    else {
      // Determine if we need to chunk the updates (if batch is very large)
      const CHUNK_SIZE = 50 // Maximum updates per batch

      if (pendingUpdates.length <= CHUNK_SIZE) {
        // Process as a single batch with retry logic
        await timeAsync('batchUpdate', async () => {
          await processBatchWithRetry(pendingUpdates)
        })
      } else {
        // Split into smaller chunks
        const chunks = []
        for (let i = 0; i < pendingUpdates.length; i += CHUNK_SIZE) {
          chunks.push(pendingUpdates.slice(i, i + CHUNK_SIZE))
        }

        let successCount = 0
        let totalProcessed = 0

        // Process each chunk with progress updates
        for (let i = 0; i < chunks.length; i++) {
          const chunk = chunks[i]
          updateStatus(
            `Processing batch ${i + 1}/${chunks.length} (${
              chunk.length
            } updates)...`
          )

          try {
            await timeAsync(`batchChunk-${i + 1}`, async () => {
              const result = await processBatchWithRetry(chunk)
              successCount += result.successCount
              totalProcessed += chunk.length
            })
          } catch (error) {
            console.error(`Error processing batch ${i + 1}:`, error)
            totalProcessed += chunk.length
          }

          // Update progress
          const percent = Math.round(
            (totalProcessed / pendingUpdates.length) * 100
          )
          updateStatus(
            `Batch progress: ${percent}% (${successCount} successful)`
          )
        }

        updateStatus(
          `Batch update completed: ${successCount} of ${pendingUpdates.length} changes applied`
        )
      }
    }
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Update error: ${errorMessage}`, true)
    console.error('Batch update error:', error)
  } finally {
    // Clear pending updates
    pendingUpdates.length = 0
    updateTimeoutId = null

    // End timing for batch processing
    endTiming(batchTimerId)
  }
}

/**
 * Process a batch of updates with retry logic
 * @param batch Array of updates to process
 * @returns Result with success count
 */
async function processBatchWithRetry(
  batch: Array<{
    id: number
    field: string
    value: any
    timestamp: number
  }>
): Promise<{ successCount: number }> {
  return await withExponentialBackoff(async () => {
    // Try batch endpoint first
    try {
      const response = await fetch(`${API_BASE_URL}/batch-update-cells2`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          updates: batch.map(u => ({
            id: u.id,
            field: u.field,
            value: u.value
          }))
        })
      })

      // If batch endpoint exists and request was successful
      if (response.ok) {
        const result = await response.json()
        return { successCount: result.successCount || batch.length }
      }

      // If batch endpoint doesn't exist (404), fall back to sequential updates
      if (response.status === 404) {
        return await processSequentialUpdatesWithRetry(batch)
      }

      // For other errors, throw to trigger retry
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    } catch (error) {
      // If there's a network error or other issue, try sequential updates
      if (
        error instanceof TypeError ||
        (error instanceof Error && error.message.includes('Failed to fetch'))
      ) {
        return await processSequentialUpdatesWithRetry(batch)
      }

      // Re-throw other errors to trigger retry
      throw error
    }
  })
}

/**
 * Process updates sequentially with individual retries
 * @param batch Array of updates to process
 * @returns Result with success count
 */
async function processSequentialUpdatesWithRetry(
  batch: Array<{
    id: number
    field: string
    value: any
    timestamp: number
  }>
): Promise<{ successCount: number }> {
  let successCount = 0

  for (const update of batch) {
    try {
      await withExponentialBackoff(async () => {
        const response = await fetch(`${API_BASE_URL}/update-cell2`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            id: update.id,
            field: update.field,
            value: update.value
          })
        })

        if (!response.ok) {
          throw new Error(
            `API error: ${response.status} ${response.statusText}`
          )
        }

        const result = await response.json()
        if (result.success) {
          successCount++
        }
      })
    } catch (error) {
      console.warn(
        `Failed to update ${update.field} for product ${update.id}:`,
        error
      )
    }
  }

  return { successCount }
}
