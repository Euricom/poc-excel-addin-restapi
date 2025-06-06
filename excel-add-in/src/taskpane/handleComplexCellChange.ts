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
import { startTiming, endTiming } from '../utils/performanceUtils'

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
 */
async function processPendingUpdates(): Promise<void> {
  if (pendingUpdates.length === 0) return

  const batchTimerId = startTiming('processPendingUpdates', {
    updateCount: pendingUpdates.length
  })

  updateStatus(`Sending ${pendingUpdates.length} updates to API...`)

  try {
    // For single updates, use the original API endpoint
    if (pendingUpdates.length === 1) {
      const update = pendingUpdates[0]
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
        throw new Error(`API error: ${response.status} ${response.statusText}`)
      }

      const result = await response.json()
      if (result.success) {
        updateStatus(`Updated ${update.field} successfully!`)
      } else {
        updateStatus(`Update failed: ${result.error || 'Unknown error'}`, true)
      }
    }
    // For multiple updates, use a batch endpoint (assuming it exists)
    // If not, fall back to sequential updates
    else {
      // Assuming a batch update endpoint exists
      // If not, this would need to be replaced with sequential calls
      const response = await fetch(`${API_BASE_URL}/batch-update-cells2`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          updates: pendingUpdates.map(u => ({
            id: u.id,
            field: u.field,
            value: u.value
          }))
        })
      })

      // If batch endpoint doesn't exist, fall back to sequential updates
      if (response.status === 404) {
        let successCount = 0

        // Process updates sequentially
        for (const update of pendingUpdates) {
          const singleResponse = await fetch(`${API_BASE_URL}/update-cell2`, {
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

          if (singleResponse.ok) {
            const result = await singleResponse.json()
            if (result.success) {
              successCount++
            }
          }
        }

        updateStatus(
          `Updated ${successCount} of ${pendingUpdates.length} changes successfully`
        )
      }
      // Process batch response
      else if (response.ok) {
        const result = await response.json()
        updateStatus(
          `Batch update completed: ${
            result.successCount || 'all'
          } changes applied`
        )
      } else {
        throw new Error(`API error: ${response.status} ${response.statusText}`)
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
