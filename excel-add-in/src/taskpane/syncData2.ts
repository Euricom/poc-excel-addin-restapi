import { handleComplexCellChange } from './handleComplexCellChange'
import { updateStatus } from '../utils/taskpaneUtils'
import { ComplexApiResponse } from '../types/complexProduct'
import {
  COMPLEX_PRODUCT_HEADERS,
  productToRow
} from '../utils/complexProductUtils'

const API_BASE_URL = 'http://localhost:3001/api'

let complexData: ComplexApiResponse = { products: [] }

/**
 * Synchronizes complex product data from the API to the active Excel worksheet.
 * - Fetches data from the API.
 * - Writes headers and data rows.
 * - Applies formatting.
 * - Registers cell change event handler.
 * - Handles errors and updates status.
 */
export async function syncData2(): Promise<void> {
  try {
    updateStatus('Fetching complex data from API...')

    // Fetch complex data from API
    const response = await fetch(`${API_BASE_URL}/data2`)
    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }
    complexData = await response.json()
    console.log('API response (complex data):', complexData)

    updateStatus('Writing complex data to Excel...')

    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Remove previous event handler to avoid duplicates
      sheet.onChanged.remove(handleComplexCellChange)

      // Write header row
      const headerRange = sheet.getRange('A1:AX1')
      headerRange.values = [COMPLEX_PRODUCT_HEADERS]
      headerRange.format.font.bold = true

      // Write data rows if available
      if (complexData.products && complexData.products.length > 0) {
        const rowCount = complexData.products.length
        const dataRange = sheet.getRange(`A2:AX${rowCount + 1}`)
        dataRange.values = complexData.products.map(productToRow)

        // --- Column formatting ---
        // Price as currency (D)
        sheet.getRange(`D2:D${rowCount + 1}`).numberFormat = [['$#,##0.00']]
        // Date Added (H)
        sheet.getRange(`H2:H${rowCount + 1}`).numberFormat = [['yyyy-mm-dd']]
        // Boolean columns (G, AA:AR)
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
        booleanColumns.forEach(
          col =>
            (sheet.getRange(`${col}2:${col}${rowCount + 1}`).numberFormat = [
              ['@']
            ])
        )
        // Weight (L)
        sheet.getRange(`L2:L${rowCount + 1}`).numberFormat = [['0.00']]
        // Discount Percentage (W)
        sheet.getRange(`W2:W${rowCount + 1}`).numberFormat = [['0.00%']]
        // Tax Rate (X)
        sheet.getRange(`X2:X${rowCount + 1}`).numberFormat = [['0.00%']]
        // Shipping Weight (Y)
        sheet.getRange(`Y2:Y${rowCount + 1}`).numberFormat = [['0.00']]
        // Power Consumption (AW)
        sheet.getRange(`AW2:AW${rowCount + 1}`).numberFormat = [['0.00 W']]
        // Last Updated (AW)
        sheet.getRange(`AW2:AW${rowCount + 1}`).numberFormat = [['yyyy-mm-dd']]
        // Popularity (AX)
        sheet.getRange(`AX2:AX${rowCount + 1}`).numberFormat = [['0.00']]
      }

      // Autofit columns for readability
      sheet.getUsedRange().format.autofitColumns()

      // Register the cell change event handler
      sheet.onChanged.add(handleComplexCellChange)

      await context.sync()

      updateStatus(
        `Complex data synchronized successfully! (${complexData.products.length} products)`
      )
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Error: ${errorMessage}`, true)
    console.error('Complex data sync error:', error)
  }
}
