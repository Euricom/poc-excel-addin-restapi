import { handleComplexCellChange } from './handleComplexCellChange'
import { updateStatus } from '../utils/taskpaneUtils'
import { ComplexApiResponse } from '../types/complexProduct'
import {
  COMPLEX_PRODUCT_HEADERS,
  productToRow,
  productsToRows
} from '../utils/complexProductUtils'
import { timeAsync, logPerformanceReport } from '../utils/performanceUtils'

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
 */
async function writeDataToExcel(): Promise<void> {
  if (!complexDataCache.products || complexDataCache.products.length === 0) {
    updateStatus('No products to synchronize')
    return
  }

  updateStatus('Writing complex data to Excel...')

  await Excel.run(async context => {
    // Get worksheet and clear previous data
    const sheet = context.workbook.worksheets.getActiveWorksheet()

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

    // Write all data at once (headers + data)
    const writeStart = performance.now()
    const dataRange = sheet.getRange(`A1:AX${rowCount + 1}`)
    dataRange.values = allRows
    console.log(
      `Range values assignment took ${(performance.now() - writeStart).toFixed(
        2
      )}ms`
    )

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
 * This can be used for very large datasets instead of the standard syncData2
 */
export async function syncLargeDataset(): Promise<void> {
  try {
    updateStatus('Preparing to sync large dataset...')

    // Fetch data
    await fetchComplexDataWithCaching()

    const BATCH_SIZE = 100 // Process 100 rows at a time
    const products = complexDataCache.products
    const totalProducts = products.length

    if (totalProducts === 0) {
      updateStatus('No products to synchronize')
      return
    }

    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Remove previous event handler
      sheet.onChanged.remove(handleComplexCellChange)

      // Write headers
      const headerRange = sheet.getRange('A1:AX1')
      headerRange.values = [COMPLEX_PRODUCT_HEADERS]
      headerRange.format.font.bold = true

      // Process data in batches
      for (let i = 0; i < totalProducts; i += BATCH_SIZE) {
        const batchEnd = Math.min(i + BATCH_SIZE, totalProducts)
        const batchProducts = products.slice(i, batchEnd)

        updateStatus(
          `Processing products ${i + 1} to ${batchEnd} of ${totalProducts}...`
        )

        // Write batch data
        const startRow = i + 2 // +2 because row 1 is header
        const endRow = batchEnd + 1
        const batchRange = sheet.getRange(`A${startRow}:AX${endRow}`)
        batchRange.values = productsToRows(batchProducts)

        // Apply minimal formatting for this batch
        sheet.getRange(`D${startRow}:D${endRow}`).numberFormat = [['$#,##0.00']]

        // Sync after each batch to avoid timeout
        await context.sync()
      }

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
