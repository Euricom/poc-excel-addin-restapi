/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Excel */

// API configuration
const API_BASE_URL = "http://localhost:3001/api";

interface Product {
  id: number;
  name: string;
  price: number;
}

interface ApiResponse {
  products: Product[];
}

interface CellUpdateRequest {
  id: number;
  field: string;
  value: string | number;
}

interface CellUpdateResponse {
  success: boolean;
  updatedProduct?: Product;
  error?: string;
}

interface ExcelFormula {
  id: string;
  name: string;
  description: string;
  formula: string;
  defaultLocation?: string;
}

interface FormulasResponse {
  formulas: ExcelFormula[];
}

let data: ApiResponse = { products: [] };

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Register event handlers
    document.getElementById("fetchData")?.addEventListener("click", syncData);

    // Initial data load
    syncData();
  }
});


/**
 * Applies formulas from the API to the worksheet
 */
async function applyFormulas(): Promise<void> {
  try {
    updateStatus("Fetching formulas from API...");

    // Fetch formulas from the API
    const response = await fetch(`${API_BASE_URL}/formulas`);

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    const data: FormulasResponse = await response.json();

    if (!data.formulas || data.formulas.length === 0) {
      updateStatus("No formulas available from API");
      return;
    }

    // Apply the formulas to the worksheet
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the data range to know where our data ends
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      const lastRow = usedRange.rowCount;

      // Add a header for the summary section
      const summaryHeaderRange = sheet.getRange(`A${lastRow + 2}`);
      summaryHeaderRange.values = [["Summary"]];
      summaryHeaderRange.format.font.bold = true;

      // Process and apply each formula from the API
      for (let i = 0; i < data.formulas.length; i++) {
        const formula = data.formulas[i];
        const rowIndex = lastRow + 3 + i;

        // Add label in column A
        const labelCell = sheet.getRange(`A${rowIndex}`);
        labelCell.values = [[formula.name]];

        // Process the formula by replacing placeholders
        const processedFormula = formula.formula.replace(/\{lastRow\}/g, lastRow.toString());

        // Add formula in column C
        const formulaCell = sheet.getRange(`C${rowIndex}`);
        formulaCell.formulas = [[processedFormula]];
      }

      // Auto-fit columns
      sheet.getUsedRange().format.autofitColumns();

      await context.sync();
      updateStatus(`Applied ${data.formulas.length} formulas from API`);
    });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    updateStatus(`Formula error: ${errorMessage}`, true);
    console.error("Formula application error:", error);
  }
}

/**
 * Synchronize data from API to Excel
 */
async function syncData(): Promise<void> {
  try {
    updateStatus("Fetching data from API...");

    // Call your API endpoint
    const response = await fetch(`${API_BASE_URL}/data`);

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    data =  await response.json();
    
    console.log("API response:", data);

    updateStatus("Writing data to Excel...");

    // Write data to Excel
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Clear previous event handlers to avoid duplicates
      sheet.onChanged.remove(handleCellChange);

      // Create a header row
      const headerRange = sheet.getRange("A1:C1");
      headerRange.values = [["ID", "Product Name", "Price"]];
      headerRange.format.font.bold = true;

      console.log("Writing to Excel:", data.products.map(p => [p.id, p.name, p.price]));

      // Write data rows
      if (data.products && data.products.length > 0) {
        const dataRange = sheet.getRange(`A2:C${data.products.length + 1}`);
        dataRange.values = data.products.map(product => [
          product.id,
          product.name,
          product.price
        ]);

        // Format price column as currency
        const priceRange = sheet.getRange(`C2:C${data.products.length + 1}`);
        priceRange.numberFormat = [["$#,##0.00"]];
      }

      // Format table for readability
      sheet.getUsedRange().format.autofitColumns();

      // Register a cell change event handler
      sheet.onChanged.add(handleCellChange);

      await context.sync();

      updateStatus(`Data synchronized! Now applying formulas...`);
      // After syncing data, apply formulas from API
      await applyFormulas();
      
      updateStatus(`Data synchronized successfully! (${data.products.length} products)`);
    });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    updateStatus(`Error: ${errorMessage}`, true);
    console.error("Sync error:", error);
  }
}

/**
 * Handle cell changes and send updates to API
 */
async function handleCellChange(event: Excel.WorksheetChangedEventArgs): Promise<void> {
  await Excel.run(async (context) => {
    try {
      // Get details about the changed range
      const changedRange = context.workbook.worksheets.getActiveWorksheet().getRange(event.address);
      changedRange.load(["values", "rowIndex", "columnIndex"]);

      await context.sync();

      // Skip if not a single cell change
      if (changedRange.values?.length !== 1 || changedRange.values[0].length !== 1) {
        return;
      }

      const rowIndex = changedRange.rowIndex;
      const colIndex = changedRange.columnIndex;
      const newValue = changedRange.values[0][0];

      // Skip header row
      if (rowIndex === 0) {
        return;
      }

      // Get the used range to determine where the data table ends
      const usedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      // Skip if the cell is in the formula/summary section (below the data table)
      // First, find where our data products end
      const dataRowCount = data.products ? data.products.length + 1 : 0; // +1 for header

      // If we're in a row beyond the data table, skip processing
      // We compare rowIndex to dataRowCount-1 because rowIndex is 0-based (0 is header, 1 is first data row)
      if (rowIndex > dataRowCount - 1) {
        console.log(`Skipping change in formula section - row ${rowIndex + 1}`); // +1 for display
        return;
      }

      // Continue with normal processing for data table cells

      // Skip ID column changes (column A)
      if (colIndex === 0) {
        updateStatus("Product ID cannot be modified directly", true);
        return;
      }

      // Skip if not in our data columns (A-C)
      if (colIndex > 2) {
        return;
      }

      // Get the product ID from column A
      const idCell = context.workbook.worksheets.getActiveWorksheet().getRange(`A${rowIndex + 1}`);
      idCell.load("values");
      await context.sync();

      const productIdValue = idCell.values[0][0];

      // Safely convert to number, validating we have a proper ID
      let productId: number;

      if (typeof productIdValue === 'number') {
        productId = productIdValue;
      } else if (typeof productIdValue === 'string' && !isNaN(parseInt(productIdValue))) {
        productId = parseInt(productIdValue);
      } else {
        console.log(`Invalid product ID: ${productIdValue}`);
        updateStatus("Cannot identify product ID", true);
        return;
      }

      // Determine which field was updated
      let field: string;
      switch (colIndex) {
        case 1: field = "name"; break;
        case 2: field = "price"; break;
        default: return;
      }

      // Prepare update data for API
      const updateData: CellUpdateRequest = {
        id: productId,
        field: field,
        value: field === "price" ? parseFloat(newValue) : newValue
      };

      // Send update to API
      updateStatus(`Sending ${field} update to API...`);
      await sendUpdateToApi(updateData);

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      updateStatus(`Cell change error: ${errorMessage}`, true);
      console.error("Cell change handling error:", error);
    }
  });
}

/**
 * Send cell update to API
 */
async function sendUpdateToApi(updateData: CellUpdateRequest): Promise<void> {
  try {
    const response = await fetch(`${API_BASE_URL}/update-cell`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(updateData)
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    const result: CellUpdateResponse = await response.json();

    if (result.success) {
      updateStatus(`Updated ${updateData.field} successfully!`);
    } else {
      updateStatus(`Update failed: ${result.error || "Unknown error"}`, true);
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    updateStatus(`API Error: ${errorMessage}`, true);
    console.error("API update error:", error);
  }
}

/**
 * Update status message in the UI
 */
function updateStatus(message: string, isError: boolean = false): void {
  const statusElement = document.getElementById("status");
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = isError ? "status error" : "status success";
  }
}