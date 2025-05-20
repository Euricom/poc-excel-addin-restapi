# Excel API Integration

A comprehensive solution for integrating Excel with a custom REST API backend. This project consists of an Excel add-in and a Node.js API server, allowing seamless data synchronization between Excel and your data source.

## Project Overview

This project provides a complete system for:
- Loading data from a REST API into Excel spreadsheets
- Synchronizing changes made in Excel back to the API
- Applying dynamic formulas to analyze the data
- Real-time notification of sync status

## Project Structure

```
excel-api-integration/
├── api/                      # Backend REST API
│   ├── src/                  # API source code
│   │   ├── index.ts          # Main API entry point
│   ├── package.json          # API dependencies
│   └── tsconfig.json         # TypeScript configuration for API
│
├── excel-add-in/             # Excel add-in frontend
│   ├── assets/               # Images and static assets
│   ├── src/                  # Add-in source code
│   │   ├── commands/         # Add-in command handlers
│   │   └── taskpane/         # Taskpane UI implementation
│   │       ├── taskpane.html # Main UI layout
│   │       ├── taskpane.css  # Styling for the add-in
│   │       └── taskpane.ts   # Main add-in logic
│   ├── manifest.xml          # Add-in manifest for Office
│   ├── package.json          # Add-in dependencies
│   └── tsconfig.json         # TypeScript configuration for add-in
│
├── package.json              # Root scripts for running both projects
└── README.md                 # Project documentation
```

## Features

- **Data Synchronization**: Load data from API to Excel and push changes back
- **Real-time Updates**: Instantly send cell changes to the backend
- **Automatic Formulas**: Apply API-defined formulas to create summaries and analytics
- **TypeScript Support**: Fully typed codebase for better development experience
- **Development Mode**: Hot-reloading for both API and add-in during development

## Prerequisites

- [Node.js](https://nodejs.org/) (v14 or newer)
- [npm](https://www.npmjs.com/)
- Microsoft Excel (Desktop or Online)

## Getting Started

### Installation

1. Clone the repository
   ```bash
   git clone https://github.com/yourusername/excel-api-integration.git
   cd excel-api-integration
   ```

2. Install dependencies (root, API, and add-in)
   ```bash
   npm install
    cd api && npm install
    cd ../excel-add-in && npm install
    cd ..
   ```

### Running in development mode
Start both the API server and Excel add-in in development mode:
```bash
npm start
```

This will:

- Start the API server at http://localhost:3001
- Launch the Excel add-in development server
- Open Excel with the add-in loaded (if using desktop mode)

### Running Components Separately
Start just the API with hot-reloading:
```bash
npm run start:api
```

Start just the Excel add-in with hot-reloading:
```bash
npm run start:addin
```

### Building for production
Build the API and add-in for production:
```bash
npm run build
```

## API Server
The Node.js REST API server provides endpoints for:
- GET /api/data: Fetch all product data
- GET /api/data/:id: Fetch a specific product by ID
- POST /api/update-cell: Update a specific field of a product
- GET /api/formulas: Get available Excel formulas for data analysis

## Excel Add-in - Detailed Overview

The Excel add-in is a critical component that provides the user interface and functionality for interacting with your API directly from Excel. Here's a comprehensive breakdown of how it works:

### Key Components

#### Manifest.xml

The manifest file defines the add-in's properties, permissions, and entry points. It tells Excel:
- How to load the add-in
- What permissions it needs
- Where to position it in the document
- What icons and UI elements to use

#### UI Components (taskpane.html/css)

The taskpane interface includes:
- A sync button that triggers data retrieval from the API
- A status indicator showing the current operation state
- Visual feedback for successful/failed operations

#### Core Functionality (taskpane.ts)

The TypeScript code handles several critical tasks:

1. **Data Synchronization**
   - Fetches product data from the API
   - Maps API data to Excel cells in a structured format
   - Formats cells appropriately (e.g., currency formatting for prices)

2. **Change Detection**
   - Listens for cell changes using Excel's event system
   - Identifies which field was changed and for which product
   - Validates changes before sending to the API
   - Prevents direct editing of product IDs for data integrity

3. **Formula Application**
   - Retrieves formula definitions from the API
   - Places formulas in appropriate cells below the data table
   - Handles dynamic placement based on the data size
   - Formats formula results based on their type (e.g., currency formatting)

4. **Error Handling**
   - Gracefully manages API connection issues
   - Provides clear user feedback for all operations
   - Prevents common user errors

### Data Flow

1. **Initial Load**:
   - Add-in initializes when Excel starts
   - `Office.onReady()` triggers initial data sync
   - API data is fetched and displayed in Excel

2. **User-Initiated Sync**:
   - User clicks the sync button
   - Latest data is fetched from API
   - Excel table is refreshed with current data
   - Formulas are reapplied based on the new data

3. **Cell Edits**:
   - User modifies a cell value
   - Change event is detected
   - Modified value is sent to API
   - Success/failure feedback is displayed

4. **Formula Application**:
   - Either triggered automatically after data sync
   - Or manually initiated by the user
   - Formula definitions come from the API
   - Applied dynamically based on the current data table size

### Technical Implementation

The add-in uses several advanced Excel and Office.js concepts:

- **Excel.run() Pattern**: Efficient batching of Excel operations
- **Event Handlers**: For detecting cell changes with `onChanged` events
- **Range Manipulation**: Precise control over cell ranges, values, and formatting
- **Formula Application**: Dynamic formula creation and application
- **Context Synchronization**: Proper management of Excel's asynchronous context model

### Customizing the Add-in

The add-in can be customized in several ways:

1. **UI Customization**:
   - Modify `taskpane.html` and `taskpane.css` to change appearance
   - Add or remove UI elements as needed
   - Change colors, fonts, and layouts

2. **Data Handling**:
   - Update TypeScript interfaces to match your API data structure
   - Modify column mappings in the data sync function
   - Change cell formatting for different data types

3. **Advanced Features**:
   - Add filtering capabilities
   - Implement sorting options
   - Create custom visualization features

4. **Error Handling**:
   - Customize error messages
   - Add retry mechanisms
   - Implement validation for user inputs

### How It Interconnects

The add-in works with the API in a tightly coupled way:

1. **API Endpoints**: The add-in expects specific API endpoints to be available
2. **Data Format**: The TypeScript interfaces must match the API's JSON structure
3. **Update Protocol**: The cell change handler sends updates in a format expected by the API
4. **Formulas**: The formula system assumes the API provides formula definitions

This integration allows for a seamless experience where Excel effectively becomes a frontend for your data service, with bidirectional synchronization ensuring data consistency.
