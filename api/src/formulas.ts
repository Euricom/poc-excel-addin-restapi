// Define formula structure
export interface ExcelFormula {
    id: string;
    name: string;
    description: string;
    formula: string;
    defaultLocation?: string; // Optional default cell reference
}

// Add a collection of pre-defined formulas
export const formulas: ExcelFormula[] = [
    {
        id: "total_price",
        name: "Total Price",
        description: "Sum of all product prices",
        formula: "=SUM(C2:C{lastRow})",
        defaultLocation: "C{lastRow+2}"
    },
    {
        id: "average_price",
        name: "Average Price",
        description: "Average price of all products",
        formula: "=AVERAGE(C2:C{lastRow})",
        defaultLocation: "C{lastRow+3}"
    },
    {
        id: "product_count",
        name: "Product Count",
        description: "Total number of products",
        formula: "=COUNTA(A2:A{lastRow})",
        defaultLocation: "C{lastRow+4}"
    },
    {
        id: "highest_price",
        name: "Highest Price",
        description: "Most expensive product",
        formula: "=MAX(C2:C{lastRow})",
        defaultLocation: "C{lastRow+5}"
    }
];