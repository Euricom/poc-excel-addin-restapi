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

/**
 * Returns the field name for a given column index.
 */
export function getFieldNameByColumnIndex(
  index: number
): keyof ComplexProduct | undefined {
  return COMPLEX_PRODUCT_FIELD_MAP[index]
}

/**
 * Returns the column index for a given field name.
 */
export function getColumnIndexByFieldName(field: keyof ComplexProduct): number {
  return COMPLEX_PRODUCT_FIELD_MAP.indexOf(field)
}

/**
 * Converts a value to a boolean, accepting various representations.
 * Returns undefined if invalid.
 */
export function parseBoolean(value: any): boolean | undefined {
  if (
    value === true ||
    value === 'true' ||
    value === 'yes' ||
    value === 1 ||
    value === '1'
  ) {
    return true
  }
  if (
    value === false ||
    value === 'false' ||
    value === 'no' ||
    value === 0 ||
    value === '0'
  ) {
    return false
  }
  return undefined
}

/**
 * Converts a value to a number, returns undefined if invalid.
 */
export function parseNumber(value: any): number | undefined {
  const num = typeof value === 'number' ? value : parseFloat(value)
  return isNaN(num) ? undefined : num
}

/**
 * Converts a value to a string array, splitting by comma.
 */
export function parseStringArray(value: any): string[] {
  if (Array.isArray(value)) return value.map(String)
  if (typeof value === 'string') {
    return value
      .split(',')
      .map(s => s.trim())
      .filter(Boolean)
  }
  return []
}

/**
 * Validates a date string (ISO or yyyy-mm-dd).
 * Returns the string if valid, otherwise undefined.
 */
export function parseDate(value: any): string | undefined {
  if (typeof value === 'string' && !isNaN(Date.parse(value))) {
    return value
  }
  return undefined
}

/**
 * Returns a row array for Excel from a ComplexProduct object.
 * The order matches COMPLEX_PRODUCT_FIELD_MAP.
 */
export function productToRow(product: ComplexProduct): any[] {
  return COMPLEX_PRODUCT_FIELD_MAP.map(field => {
    const val = product[field]
    if (Array.isArray(val)) {
      return val.join(', ')
    }
    return val
  })
}
