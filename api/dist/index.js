"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const body_parser_1 = __importDefault(require("body-parser"));
const app = (0, express_1.default)();
const PORT = process.env.PORT || 3001;
app.use((0, cors_1.default)());
app.use(body_parser_1.default.json());
// Sample database (replace with real DB in production)
let data = {
    products: [
        { id: 1, name: 'Product A', price: 100 },
        { id: 2, name: 'Product B', price: 200 },
        { id: 3, name: 'Product C', price: 300 },
    ]
};
// GET endpoint to fetch data
app.get('/api/data', (req, res) => {
    res.json(data);
});
// GET endpoint to fetch a specific item
app.get('/api/data/:id', (req, res) => {
    const id = parseInt(req.params.id);
    const product = data.products.find((p) => p.id === id);
    if (product) {
        res.json(product);
    }
    else {
        res.status(404).json({ error: 'Product not found' });
    }
});
// POST endpoint to update a value
app.post('/api/update-cell', (req, res) => {
    const { id, field, value } = req.body;
    const index = data.products.findIndex((p) => p.id === id);
    if (index !== -1) {
        data.products[index][field] = value;
        res.json({ success: true, updatedProduct: data.products[index] });
    }
    else {
        res.status(404).json({ error: 'Product not found' });
    }
});
app.listen(PORT, () => {
    console.log(`API server running on port ${PORT}`);
});
