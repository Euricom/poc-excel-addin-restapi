// Utility functions and constants for ComplexProduct table handling in Excel Add-in

import { ComplexProduct } from '../types/complexProduct'

/**
 * Array of column headers for the ComplexProduct table (50 columns).
 * The order must match the field mapping below.
 */
export const COMPLEX_PRODUCT_HEADERS: string[] = [
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
  'Last Updated',
  'Popularity'
]

/**
 * Array mapping column index to ComplexProduct field name.
 * The order must match COMPLEX_PRODUCT_HEADERS.
 */
export const COMPLEX_PRODUCT_FIELD_MAP: (keyof ComplexProduct)[] = [
  'id',
  'name',
  'category',
  'price',
  'quantity',
  'rating',
  'inStock',
  'dateAdded',
  'description',
  'tags',
  'color',
  'weight',
  'dimensions',
  'manufacturer',
  'countryOfOrigin',
  'warranty',
  'material',
  'sku',
  'barcode',
  'minOrderQuantity',
  'maxOrderQuantity',
  'discountPercentage',
  'taxRate',
  'shippingWeight',
  'shippingDimensions',
  'returnPolicy',
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
  'biodegradable',
  'energyEfficiencyRating',
  'noiseLevel',
  'powerConsumption',
  'certifications',
  'lastUpdated',
  'popularity'
]

// Lookup maps for faster field/column access
const fieldToColumnMap = new Map<keyof ComplexProduct, number>()
const columnToFieldMap = new Map<number, keyof ComplexProduct>()

// Initialize lookup maps
COMPLEX_PRODUCT_FIELD_MAP.forEach((field, index) => {
  fieldToColumnMap.set(field, index)
  columnToFieldMap.set(index, field)
})

/**
 * Returns the field name for a given column index.
 * Optimized with Map lookup instead of array search.
 */
export function getFieldNameByColumnIndex(
  index: number
): keyof ComplexProduct | undefined {
  return columnToFieldMap.get(index)
}

/**
 * Returns the column index for a given field name.
 * Optimized with Map lookup instead of array search.
 */
export function getColumnIndexByFieldName(field: keyof ComplexProduct): number {
  const index = fieldToColumnMap.get(field)
  return index !== undefined ? index : -1
}

// Optimized boolean parsing with Set for faster lookups
const trueBooleanValues = new Set([true, 'true', 'yes', 1, '1'])
const falseBooleanValues = new Set([false, 'false', 'no', 0, '0'])

/**
 * Converts a value to a boolean, accepting various representations.
 * Returns undefined if invalid.
 * Optimized with Set for faster lookups.
 */
export function parseBoolean(value: any): boolean | undefined {
  if (trueBooleanValues.has(value)) {
    return true
  }
  if (falseBooleanValues.has(value)) {
    return false
  }
  return undefined
}

/**
 * Converts a value to a number, returns undefined if invalid.
 * Optimized with type checking.
 */
export function parseNumber(value: any): number | undefined {
  if (typeof value === 'number') return isNaN(value) ? undefined : value
  if (typeof value === 'string') {
    const num = parseFloat(value)
    return isNaN(num) ? undefined : num
  }
  return undefined
}

/**
 * Converts a value to a string array, splitting by comma.
 * Optimized for different input types.
 */
export function parseStringArray(value: any): string[] {
  if (Array.isArray(value)) {
    return value.map(item => String(item).trim()).filter(Boolean)
  }
  if (typeof value === 'string') {
    return value
      .split(',')
      .map(s => s.trim())
      .filter(Boolean)
  }
  return []
}

// Date validation regex for common formats
const dateRegex = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d{3})?Z?)?$/

/**
 * Validates a date string (ISO or yyyy-mm-dd).
 * Returns the string if valid, otherwise undefined.
 * Optimized with regex pre-check before expensive Date.parse.
 */
export function parseDate(value: any): string | undefined {
  if (typeof value === 'string') {
    // Quick check with regex before expensive Date.parse
    if (dateRegex.test(value) || !isNaN(Date.parse(value))) {
      return value
    }
  }
  return undefined
}

// Pre-allocate array for row conversion to avoid repeated allocations
const rowBuffer: any[] = new Array(COMPLEX_PRODUCT_FIELD_MAP.length).fill(null)

/**
 * Returns a row array for Excel from a ComplexProduct object.
 * The order matches COMPLEX_PRODUCT_FIELD_MAP.
 * Optimized to reuse array buffer and minimize allocations.
 */
export function productToRow(product: ComplexProduct): any[] {
  for (let i = 0; i < COMPLEX_PRODUCT_FIELD_MAP.length; i++) {
    const field = COMPLEX_PRODUCT_FIELD_MAP[i]
    const val = product[field]

    // Handle array values
    if (Array.isArray(val)) {
      rowBuffer[i] = val.join(', ')
    } else {
      rowBuffer[i] = val
    }
  }

  // Return a copy of the buffer to avoid reference issues
  return [...rowBuffer]
}

/**
 * Batch converts multiple products to rows in a single operation.
 * More efficient than mapping each product individually.
 */
export function productsToRows(products: ComplexProduct[]): any[][] {
  const rows: any[][] = new Array(products.length)

  for (let i = 0; i < products.length; i++) {
    rows[i] = productToRow(products[i])
  }

  return rows
}

/**
 * Optimized function to format a cell value based on its field type
 * Used for consistent formatting across the application
 */
export function formatCellValue(
  field: keyof ComplexProduct,
  value: any
): string {
  if (value === null || value === undefined) return ''

  // Format based on field type
  const fieldIndex = getColumnIndexByFieldName(field)

  // Numeric fields with currency
  if (field === 'price') {
    return `$${Number(value).toFixed(2)}`
  }

  // Percentage fields
  if (field === 'discountPercentage' || field === 'taxRate') {
    return `${(Number(value) * 100).toFixed(2)}%`
  }

  // Boolean fields
  if (
    field === 'inStock' ||
    field.includes('Resistant') ||
    field === 'assemblyRequired' ||
    field === 'batteryRequired' ||
    field === 'batteriesIncluded' ||
    field === 'recyclable' ||
    field === 'biodegradable'
  ) {
    return value ? 'Yes' : 'No'
  }

  // Array fields
  if (field === 'tags' || field === 'certifications') {
    if (Array.isArray(value)) {
      return value.join(', ')
    }
    return String(value)
  }

  // Default formatting
  return String(value)
}
