import express from 'express'
import cors from 'cors'
import bodyParser from 'body-parser'
import { formulas } from './formulas'

const app = express()
const PORT = process.env.PORT || 3001

app.use(cors())
app.use(bodyParser.json())

// Sample database (replace with real DB in production)
let data: Record<string, any> = {
  products: [
    { id: 1, name: 'Kelder', price: 500 },
    { id: 2, name: 'Zolder', price: 250 },
    { id: 3, name: 'Handelspand', price: 100 }
  ]
}

// GET endpoint to fetch all data
app.get('/api/data', (req, res) => {
  res.json(data)
})

// GET endpoint to fetch a specific item
app.get('/api/data/:id', (req, res) => {
  const id = parseInt(req.params.id)
  const product = data.products.find((p: { id: number }) => p.id === id)

  if (product) {
    res.json(product)
  } else {
    res.status(404).json({ error: 'Product not found' })
  }
})

// POST endpoint to update a value
app.post('/api/update-cell', (req, res) => {
  console.log(req.body)
  const { id, field, value } = req.body

  const index = data.products.findIndex((p: { id: any }) => p.id === id)
  if (index !== -1) {
    data.products[index][field] = value
    res.json({ success: true, updatedProduct: data.products[index] })
  } else {
    res.status(404).json({ error: 'Product not found' })
  }
})

// GET endpoint to fetch available formulas
app.get('/api/formulas', (req, res) => {
  res.json({ formulas })
})

// POST endpoint to get a specific formula with row information
app.post('/api/formula/:id', (req, res) => {
  const { id } = req.params
  const { lastRow } = req.body

  const formula = formulas.find(f => f.id === id)

  if (!formula) {
    return res.status(404).json({ error: 'Formula not found' })
  }

  // Replace {lastRow} placeholder with actual row number
  const processedFormula = {
    ...formula,
    formula: formula.formula.replace(/\{lastRow\}/g, lastRow.toString()),
    defaultLocation: formula.defaultLocation?.replace(
      /\{lastRow([+-]\d+)?\}/g,
      (match, offset) => {
        if (offset) {
          return (lastRow + parseInt(offset)).toString()
        }
        return lastRow.toString()
      }
    )
  }

  res.json({ formula: processedFormula })
})

// Function to generate complex random data with 50 columns and 1000 rows
function generateComplexData() {
  const categories = [
    'Electronics',
    'Clothing',
    'Food',
    'Books',
    'Home',
    'Garden',
    'Sports',
    'Toys',
    'Beauty',
    'Automotive'
  ]
  const adjectives = [
    'Amazing',
    'Incredible',
    'Fantastic',
    'Excellent',
    'Superior',
    'Premium',
    'Deluxe',
    'Ultimate',
    'Essential',
    'Professional'
  ]
  const nouns = [
    'Product',
    'Item',
    'Solution',
    'Tool',
    'Device',
    'System',
    'Kit',
    'Package',
    'Set',
    'Collection'
  ]
  const tagOptions = [
    'New',
    'Sale',
    'Popular',
    'Limited',
    'Exclusive',
    'Trending',
    'Seasonal',
    'Organic',
    'Eco-friendly',
    'Handmade'
  ]

  const result: Array<{
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
  }> = []

  for (let i = 1; i <= 10000; i++) {
    // Generate random name
    const randomAdjective =
      adjectives[Math.floor(Math.random() * adjectives.length)]
    const randomNoun = nouns[Math.floor(Math.random() * nouns.length)]
    const name = `${randomAdjective} ${randomNoun} ${i}`

    // Generate random category
    const category = categories[Math.floor(Math.random() * categories.length)]

    // Generate random price (between 10 and 1000)
    const price = parseFloat((Math.random() * 990 + 10).toFixed(2))

    // Generate random quantity (between 0 and 100)
    const quantity = Math.floor(Math.random() * 101)

    // Generate random rating (between 1 and 5)
    const rating = parseFloat((Math.random() * 4 + 1).toFixed(1))

    // Generate random inStock status
    const inStock = quantity > 0

    // Generate random date within the last year
    const today = new Date()
    const oneYearAgo = new Date()
    oneYearAgo.setFullYear(today.getFullYear() - 1)
    const randomDate = new Date(
      oneYearAgo.getTime() +
        Math.random() * (today.getTime() - oneYearAgo.getTime())
    )
    const dateAdded = randomDate.toISOString().split('T')[0]

    // Generate random description
    const description = `This is a ${randomAdjective.toLowerCase()} ${randomNoun.toLowerCase()} suitable for various purposes. It comes with exceptional quality and reliability.`

    // Generate random tags (1-3 tags)
    const numTags = Math.floor(Math.random() * 3) + 1
    const tags: string[] = []
    for (let j = 0; j < numTags; j++) {
      const randomTag =
        tagOptions[Math.floor(Math.random() * tagOptions.length)]
      if (!tags.includes(randomTag)) {
        tags.push(randomTag)
      }
    }

    // Generate additional random data for the new columns
    const colors = [
      'Red',
      'Blue',
      'Green',
      'Yellow',
      'Black',
      'White',
      'Purple',
      'Orange',
      'Pink',
      'Brown'
    ]
    const manufacturers = [
      'TechCorp',
      'GlobalIndustries',
      'InnovateInc',
      'PrimeSolutions',
      'EliteBrands'
    ]
    const countries = [
      'USA',
      'China',
      'Germany',
      'Japan',
      'Canada',
      'UK',
      'France',
      'Italy',
      'South Korea',
      'Taiwan'
    ]
    const materials = [
      'Plastic',
      'Metal',
      'Wood',
      'Glass',
      'Ceramic',
      'Fabric',
      'Leather',
      'Silicone',
      'Rubber',
      'Composite'
    ]
    const warranties = [
      '1 Year',
      '2 Years',
      '3 Years',
      '5 Years',
      'Lifetime',
      '30 Days',
      '90 Days',
      '6 Months',
      'None'
    ]
    const energyRatings = ['A+++', 'A++', 'A+', 'A', 'B', 'C', 'D', 'E']
    const noiseLevels = [
      'Silent',
      'Very Quiet',
      'Quiet',
      'Moderate',
      'Loud',
      'Very Loud'
    ]
    const certifications = [
      'CE',
      'ISO9001',
      'UL',
      'RoHS',
      'Energy Star',
      'FSC',
      'Fair Trade',
      'Organic',
      'Vegan'
    ]

    // Generate random values for additional columns
    const color = colors[Math.floor(Math.random() * colors.length)]
    const weight = parseFloat((Math.random() * 20 + 0.1).toFixed(2))
    const dimensions = `${Math.floor(Math.random() * 50 + 5)}x${Math.floor(
      Math.random() * 50 + 5
    )}x${Math.floor(Math.random() * 20 + 1)}`
    const manufacturer =
      manufacturers[Math.floor(Math.random() * manufacturers.length)]
    const countryOfOrigin =
      countries[Math.floor(Math.random() * countries.length)]
    const warranty = warranties[Math.floor(Math.random() * warranties.length)]
    const material = materials[Math.floor(Math.random() * materials.length)]
    const sku = `SKU-${Math.floor(Math.random() * 10000)
      .toString()
      .padStart(5, '0')}`
    const barcode = `${Math.floor(Math.random() * 10000000000000)
      .toString()
      .padStart(13, '0')}`
    const minOrderQuantity = Math.floor(Math.random() * 5 + 1)
    const maxOrderQuantity = Math.floor(Math.random() * 100 + minOrderQuantity)
    const discountPercentage = parseFloat((Math.random() * 50).toFixed(2))
    const taxRate = parseFloat((Math.random() * 25).toFixed(2))
    const shippingWeight = parseFloat((weight * 1.2).toFixed(2))
    const shippingDimensions = `${parseInt(dimensions.split('x')[0]) + 2}x${
      parseInt(dimensions.split('x')[1]) + 2
    }x${parseInt(dimensions.split('x')[2]) + 2}`
    const returnPolicy = Math.random() > 0.5 ? '30 Days' : '14 Days'
    const assemblyRequired = Math.random() > 0.7
    const batteryRequired = Math.random() > 0.7
    const batteriesIncluded = batteryRequired && Math.random() > 0.5
    const waterproof = Math.random() > 0.7
    const heatResistant = Math.random() > 0.7
    const coldResistant = Math.random() > 0.7
    const uvResistant = Math.random() > 0.7
    const windResistant = Math.random() > 0.7
    const shockResistant = Math.random() > 0.7
    const dustResistant = Math.random() > 0.7
    const scratchResistant = Math.random() > 0.7
    const stainResistant = Math.random() > 0.7
    const fadeResistant = Math.random() > 0.7
    const rustResistant = Math.random() > 0.7
    const moldResistant = Math.random() > 0.7
    const fireResistant = Math.random() > 0.7
    const recyclable = Math.random() > 0.3
    const biodegradable = Math.random() > 0.7
    const energyEfficiencyRating =
      energyRatings[Math.floor(Math.random() * energyRatings.length)]
    const noiseLevel =
      noiseLevels[Math.floor(Math.random() * noiseLevels.length)]
    const powerConsumption = parseFloat((Math.random() * 1000).toFixed(2))

    // Generate random certifications (1-3)
    const numCerts = Math.floor(Math.random() * 3) + 1
    const productCertifications: string[] = []

    // Generate a random date within the last month for lastUpdated
    const oneMonthAgo = new Date()
    oneMonthAgo.setMonth(today.getMonth() - 1)
    const randomUpdateDate = new Date(
      oneMonthAgo.getTime() +
        Math.random() * (today.getTime() - oneMonthAgo.getTime())
    )
    const lastUpdated = randomUpdateDate.toISOString().split('T')[0]

    // Generate random popularity score between 1 and 100
    const popularity = Math.floor(Math.random() * 100) + 1
    for (let j = 0; j < numCerts; j++) {
      const randomCert =
        certifications[Math.floor(Math.random() * certifications.length)]
      if (!productCertifications.includes(randomCert)) {
        productCertifications.push(randomCert)
      }
    }

    result.push({
      id: i,
      name,
      category,
      price,
      quantity,
      rating,
      inStock,
      dateAdded,
      description,
      tags,
      // Additional columns
      color,
      weight,
      dimensions,
      manufacturer,
      countryOfOrigin,
      warranty,
      material,
      sku,
      barcode,
      minOrderQuantity,
      maxOrderQuantity,
      discountPercentage,
      taxRate,
      shippingWeight,
      shippingDimensions,
      returnPolicy,
      assemblyRequired,
      batteryRequired,
      batteriesIncluded,
      waterproof,
      heatResistant,
      coldResistant,
      uvResistant,
      windResistant,
      shockResistant,
      dustResistant,
      scratchResistant,
      stainResistant,
      fadeResistant,
      rustResistant,
      moldResistant,
      fireResistant,
      recyclable,
      biodegradable,
      energyEfficiencyRating,
      noiseLevel,
      powerConsumption,
      certifications: productCertifications,
      lastUpdated,
      popularity
    })
  }

  return result
}

// GET endpoint to fetch complex data
app.get('/api/data2', (req, res) => {
  const products = generateComplexData()
  res.json({ products })
})

app.listen(PORT, () => {
  console.log(`API server running on port ${PORT}`)
})
