export interface ComplexProduct {
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

export interface ComplexApiResponse {
  products: ComplexProduct[]
}
