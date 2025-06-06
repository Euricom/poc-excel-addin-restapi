import { updateStatus } from '../utils/taskpaneUtils'
import {
  COMPLEX_PRODUCT_FIELD_MAP,
  getFieldNameByColumnIndex,
  parseBoolean,
  parseNumber,
  parseStringArray,
  parseDate
} from '../utils/complexProductUtils'

const API_BASE_URL = 'http://localhost:3001/api'

/**
 * Specialized handler for cell changes in the complex data table.
 * - Validates and converts the changed value.
 * - Sends the update to the API.
 * - Handles errors and updates status.
 */
export async function handleComplexCellChange(
  event: Excel.WorksheetChangedEventArgs
): Promise<void> {
  await Excel.run(async context => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const changedRange = sheet.getRange(event.address)
      changedRange.load(['values', 'rowIndex', 'columnIndex'])
      await context.sync()

      // Only handle single cell changes
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

      // Get the used range to determine where the data table ends
      const usedRange = sheet.getUsedRange()
      usedRange.load('rowCount')
      await context.sync()
      const dataRowCount = usedRange.rowCount

      // If we're in a row beyond the data table, skip processing
      if (rowIndex > dataRowCount - 1) return

      // Prevent editing ID column
      if (colIndex === 0) {
        updateStatus('Product ID cannot be modified directly', true)
        return
      }

      // Only allow editing within defined columns
      if (colIndex < 0 || colIndex >= COMPLEX_PRODUCT_FIELD_MAP.length) return

      // Get the product ID from column A
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

      // Determine which field was updated
      const field = getFieldNameByColumnIndex(colIndex)
      if (!field) return

      // --- Value validation/conversion ---
      let value: any = newValue
      // Numeric fields
      const numericFields: (keyof import('../types/complexProduct').ComplexProduct)[] =
        [
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
      // Boolean fields
      const booleanFields: (keyof import('../types/complexProduct').ComplexProduct)[] =
        [
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
      if (
        numericFields.includes(
          field as keyof import('../types/complexProduct').ComplexProduct
        )
      ) {
        value = parseNumber(newValue)
        if (value === undefined) {
          updateStatus(`Invalid number for ${field}`, true)
          return
        }
      } else if (
        booleanFields.includes(
          field as keyof import('../types/complexProduct').ComplexProduct
        )
      ) {
        value = parseBoolean(newValue)
        if (value === undefined) {
          updateStatus(`Invalid boolean for ${field}`, true)
          return
        }
      } else if (field === 'tags' || field === 'certifications') {
        value = parseStringArray(newValue)
      } else if (field === 'dateAdded' || field === 'lastUpdated') {
        value = parseDate(newValue)
        if (value === undefined) {
          updateStatus(`Invalid date for ${field}`, true)
          return
        }
      }

      // Prepare update data for API
      const updateData = {
        id: productId,
        field,
        value
      }

      updateStatus(`Sending ${field} update to API...`)
      const response = await fetch(`${API_BASE_URL}/update-cell2`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(updateData)
      })

      if (!response.ok) {
        throw new Error(`API error: ${response.status} ${response.statusText}`)
      }

      const result = await response.json()
      if (result.success) {
        updateStatus(`Updated ${field} successfully!`)
      } else {
        updateStatus(`Update failed: ${result.error || 'Unknown error'}`, true)
      }
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : 'Unknown error'
      updateStatus(`Cell change error: ${errorMessage}`, true)
      console.error('Complex cell change handling error:', error)
    }
  })
}
