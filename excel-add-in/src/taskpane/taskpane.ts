/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Excel */

// API configuration
const API_BASE_URL = 'http://localhost:3001/api'

interface Product {
  id: number
  name: string
  price: number
}

interface ComplexProduct {
  id: number
  name: string
  category: string
  price: number
  quantity: number
  rating: number
  inStock: boolean
  dateAdded: string
  description: string
  tags: string[]
  // Additional columns to reach 50 total
  color: string
  weight: number
  dimensions: string
  manufacturer: string
  countryOfOrigin: string
  warranty: string
  material: string
  sku: string
  barcode: string
  minOrderQuantity: number
  maxOrderQuantity: number
  discountPercentage: number
  taxRate: number
  shippingWeight: number
  shippingDimensions: string
  returnPolicy: string
  assemblyRequired: boolean
  batteryRequired: boolean
  batteriesIncluded: boolean
  waterproof: boolean
  heatResistant: boolean
  coldResistant: boolean
  uvResistant: boolean
  windResistant: boolean
  shockResistant: boolean
  dustResistant: boolean
  scratchResistant: boolean
  stainResistant: boolean
  fadeResistant: boolean
  rustResistant: boolean
  moldResistant: boolean
  fireResistant: boolean
  recyclable: boolean
  biodegradable: boolean
  energyEfficiencyRating: string
  noiseLevel: string
  powerConsumption: number
  certifications: string[]
  // Adding 2 more columns to reach exactly 50
  lastUpdated: string
  popularity: number
}

interface ApiResponse {
  products: Product[]
}

interface ComplexApiResponse {
  products: ComplexProduct[]
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

let data: ApiResponse = { products: [] }
let complexData: ComplexApiResponse = { products: [] }

// Initialize Office.js
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Register event handlers
    document.getElementById('fetchData')?.addEventListener('click', syncData)
    document.getElementById('fetchData2')?.addEventListener('click', syncData2)

    // Initial data load
    syncData2()
  }
})

/**
 * Applies formulas from the API to the worksheet
 */
async function applyFormulas(): Promise<void> {
  try {
    updateStatus('Fetching formulas from API...')

    // Fetch formulas from the API
    const response = await fetch(`${API_BASE_URL}/formulas`)

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }

    const data: FormulasResponse = await response.json()

    if (!data.formulas || data.formulas.length === 0) {
      updateStatus('No formulas available from API')
      return
    }

    // Apply the formulas to the worksheet
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Get the data range to know where our data ends
      const usedRange = sheet.getUsedRange()
      usedRange.load('rowCount')
      await context.sync()

      const lastRow = usedRange.rowCount

      // Add a header for the summary section
      const summaryHeaderRange = sheet.getRange(`A${lastRow + 2}`)
      summaryHeaderRange.values = [['Summary']]
      summaryHeaderRange.format.font.bold = true

      // Process and apply each formula from the API
      for (let i = 0; i < data.formulas.length; i++) {
        const formula = data.formulas[i]
        const rowIndex = lastRow + 3 + i

        // Add label in column A
        const labelCell = sheet.getRange(`A${rowIndex}`)
        labelCell.values = [[formula.name]]

        // Process the formula by replacing placeholders
        const processedFormula = formula.formula.replace(
          /\{lastRow\}/g,
          lastRow.toString()
        )

        // Add formula in column C
        const formulaCell = sheet.getRange(`C${rowIndex}`)
        formulaCell.formulas = [[processedFormula]]
      }

      // Auto-fit columns
      sheet.getUsedRange().format.autofitColumns()

      await context.sync()
      updateStatus(`Applied ${data.formulas.length} formulas from API`)
    })
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'Unknown error'
    updateStatus(`Formula error: ${errorMessage}`, true)
    console.error('Formula application error:', error)
  }
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
      // After syncing data, apply formulas from API
      await applyFormulas()

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
async function syncData2(): Promise<void> {
  try {
    updateStatus('Fetching complex data from API...')

    // Call the complex data API endpoint
    const response = await fetch(`${API_BASE_URL}/data2`)

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`)
    }

    complexData = await response.json()

    console.log('API response (complex data):', complexData)

    updateStatus('Writing complex data to Excel...')

    // Write data to Excel
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()

      // Clear previous event handlers to avoid duplicates
      sheet.onChanged.remove(handleCellChange)

      // Create a header row with all 50 columns
      const headerRange = sheet.getRange('A1:AX1')
      headerRange.values = [
        [
          'ID',
          'Product Name',
          'Category',
          'Price',
          'Quantity',
          'Rating',
          'In Stock',
          'Date Added',
          'Description',
          'Tags',
          // Additional columns
          'Color',
          'Weight',
          'Dimensions',
          'Manufacturer',
          'Country of Origin',
          'Warranty',
          'Material',
          'SKU',
          'Barcode',
          'Min Order Quantity',
          'Max Order Quantity',
          'Discount Percentage',
          'Tax Rate',
          'Shipping Weight',
          'Shipping Dimensions',
          'Return Policy',
          'Assembly Required',
          'Battery Required',
          'Batteries Included',
          'Waterproof',
          'Heat Resistant',
          'Cold Resistant',
          'UV Resistant',
          'Wind Resistant',
          'Shock Resistant',
          'Dust Resistant',
          'Scratch Resistant',
          'Stain Resistant',
          'Fade Resistant',
          'Rust Resistant',
          'Mold Resistant',
          'Fire Resistant',
          'Recyclable',
          'Biodegradable',
          'Energy Efficiency Rating',
          'Noise Level',
          'Power Consumption',
          'Certifications',
          // Adding 2 more columns to reach exactly 50
          'Last Updated',
          'Popularity'
        ]
      ]
      headerRange.format.font.bold = true

      // Write data rows
      if (complexData.products && complexData.products.length > 0) {
        const dataRange = sheet.getRange(
          `A2:AX${complexData.products.length + 1}`
        )
        dataRange.values = complexData.products.map(product => [
          product.id,
          product.name,
          product.category,
          product.price,
          product.quantity,
          product.rating,
          product.inStock,
          product.dateAdded,
          product.description,
          product.tags.join(', '),
          // Additional columns
          product.color,
          product.weight,
          product.dimensions,
          product.manufacturer,
          product.countryOfOrigin,
          product.warranty,
          product.material,
          product.sku,
          product.barcode,
          product.minOrderQuantity,
          product.maxOrderQuantity,
          product.discountPercentage,
          product.taxRate,
          product.shippingWeight,
          product.shippingDimensions,
          product.returnPolicy,
          product.assemblyRequired,
          product.batteryRequired,
          product.batteriesIncluded,
          product.waterproof,
          product.heatResistant,
          product.coldResistant,
          product.uvResistant,
          product.windResistant,
          product.shockResistant,
          product.dustResistant,
          product.scratchResistant,
          product.stainResistant,
          product.fadeResistant,
          product.rustResistant,
          product.moldResistant,
          product.fireResistant,
          product.recyclable,
          product.biodegradable,
          product.energyEfficiencyRating,
          product.noiseLevel,
          product.powerConsumption,
          product.certifications.join(', '),
          // Adding 2 more columns to reach exactly 50
          product.lastUpdated,
          product.popularity
        ])

        // Format price column as currency
        const priceRange = sheet.getRange(
          `D2:D${complexData.products.length + 1}`
        )
        priceRange.numberFormat = [['$#,##0.00']]

        // Format date column
        const dateRange = sheet.getRange(
          `H2:H${complexData.products.length + 1}`
        )
        dateRange.numberFormat = [['yyyy-mm-dd']]

        // Format boolean columns (In Stock and other boolean properties)
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
        booleanColumns.forEach(column => {
          const booleanRange = sheet.getRange(
            `${column}2:${column}${complexData.products.length + 1}`
          )
          booleanRange.numberFormat = [['@']]
        })

        // Format additional numeric columns
        const weightRange = sheet.getRange(
          `L2:L${complexData.products.length + 1}`
        )
        weightRange.numberFormat = [['0.00']]

        const percentageRange = sheet.getRange(
          `W2:W${complexData.products.length + 1}`
        )
        percentageRange.numberFormat = [['0.00%']]

        const taxRateRange = sheet.getRange(
          `X2:X${complexData.products.length + 1}`
        )
        taxRateRange.numberFormat = [['0.00%']]

        const shippingWeightRange = sheet.getRange(
          `Y2:Y${complexData.products.length + 1}`
        )
        shippingWeightRange.numberFormat = [['0.00']]

        const powerConsumptionRange = sheet.getRange(
          `AW2:AW${complexData.products.length + 1}`
        )
        powerConsumptionRange.numberFormat = [['0.00 W']]

        // Format the two new columns
        const lastUpdatedRange = sheet.getRange(
          `AW2:AW${complexData.products.length + 1}`
        )
        lastUpdatedRange.numberFormat = [['yyyy-mm-dd']]

        const popularityRange = sheet.getRange(
          `AX2:AX${complexData.products.length + 1}`
        )
        popularityRange.numberFormat = [['0.00']]
      }

      // Format table for readability
      sheet.getUsedRange().format.autofitColumns()

      // Register a cell change event handler
      // Note: We're using the same handler as syncData for now
      sheet.onChanged.add(handleCellChange)

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

/**
 * Update status message in the UI
 */
function updateStatus(message: string, isError: boolean = false): void {
  const statusElement = document.getElementById('status')
  if (statusElement) {
    statusElement.textContent = message
    statusElement.className = isError ? 'status error' : 'status success'
  }
}
