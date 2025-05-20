import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import {formulas} from "./formulas";

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(bodyParser.json());

// Sample database (replace with real DB in production)
let data: Record<string, any> = {
    products: [
        { id: 1, name: 'Kelder', price: 500 },
        { id: 2, name: 'Zolder', price: 250 },
        { id: 3, name: 'Handelspand', price: 100 },
    ]
};

// GET endpoint to fetch all data
app.get('/api/data', (req, res) => {
    res.json(data);
});

// GET endpoint to fetch a specific item
app.get('/api/data/:id', (req, res) => {
    const id = parseInt(req.params.id);
    const product = data.products.find((p: { id: number; }) => p.id === id);

    if (product) {
        res.json(product);
    } else {
        res.status(404).json({ error: 'Product not found' });
    }
});

// POST endpoint to update a value
app.post('/api/update-cell', (req, res) => {
    console.log(req.body);
    const { id, field, value } = req.body;

    const index = data.products.findIndex((p: { id: any; }) => p.id === id);
    if (index !== -1) {
        data.products[index][field] = value;
        res.json({ success: true, updatedProduct: data.products[index] });
    } else {
        res.status(404).json({ error: 'Product not found' });
    }
});

// GET endpoint to fetch available formulas
app.get('/api/formulas', (req, res) => {
    res.json({ formulas });
});

// POST endpoint to get a specific formula with row information
app.post('/api/formula/:id', (req, res) => {
    const { id } = req.params;
    const { lastRow } = req.body;

    const formula = formulas.find(f => f.id === id);

    if (!formula) {
        return res.status(404).json({ error: 'Formula not found' });
    }

    // Replace {lastRow} placeholder with actual row number
    const processedFormula = {
        ...formula,
        formula: formula.formula.replace(/\{lastRow\}/g, lastRow.toString()),
        defaultLocation: formula.defaultLocation?.replace(/\{lastRow([+-]\d+)?\}/g, (match, offset) => {
            if (offset) {
                return (lastRow + parseInt(offset)).toString();
            }
            return lastRow.toString();
        })
    };

    res.json({ formula: processedFormula });
});

app.listen(PORT, () => {
    console.log(`API server running on port ${PORT}`);
});