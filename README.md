# pl2xl

[![JSR](https://jsr.io/badges/@jackfiszr/pl2xl)](https://jsr.io/@jackfiszr/pl2xl)
[![JSR Score](https://jsr.io/badges/@jackfiszr/pl2xl/score)](https://jsr.io/@jackfiszr/pl2xl)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/pulse)
[![GitHub last commit](https://img.shields.io/github/last-commit/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/commits/main)
[![GitHub](https://img.shields.io/github/license/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/blob/main/LICENSE)

A lightweight library for reading and writing Excel files as Polars DataFrames.\
`pl2xl` enables seamless integration between Polars and Excel, allowing you to:

- Import data from Excel files directly into a Polars DataFrame.
- Export Polars DataFrames back to Excel files, with optional Excel formatting.

## Installation

This library can be imported using the `jsr` import specifier and relies on the
`nodejs-polars` package.

### Importing the library in Deno

```typescript
import { readExcel, writeExcel } from "jsr:@jackfiszr/pl2xl@0.0.11";
import pl from "npm:nodejs-polars";
```

### Using the library in Node.js

Install the library in your Node project using `npx jsr`:

```bash
npx jsr add @jackfiszr/pl2xl
```

Then import and use it as follows:

```typescript
import { readExcel, writeExcel } from "@jackfiszr/pl2xl";
import pl from "nodejs-polars";

// Create a sample DataFrame
const inputDf = pl.DataFrame({
  Name: ["Alice", "Bob", "Charlie"],
  Age: [25, 30, 35],
  City: ["New York", "Los Angeles", "Chicago"],
});

// Write the DataFrame to an Excel file
await writeExcel(inputDf, "input.xlsx");

// Read the DataFrame back from the Excel file
const df = await readExcel("input.xlsx");
console.log("Read DataFrame:", df.toString());

// Modify the DataFrame by increasing the "Age" column by 1
const modifiedDf = df.withColumn(pl.col("Age").add(1).alias("Age"));

console.log("Modified DataFrame:", modifiedDf.toString());

// Write the modified DataFrame to a new Excel file
await writeExcel(modifiedDf, "output.xlsx");
console.log("Modified DataFrame written to output.xlsx");

// Create multiple DataFrames, one of which is empty
const emptyDf = pl.DataFrame([]);
await writeExcel([inputDf, modifiedDf, emptyDf], "multiple_sheets.xlsx", {
  sheetName: ["Input", "Modified", "Empty"],
});
console.log("Multiple DataFrames written to multiple_sheets.xlsx");
```

## API

### `readExcel(filePath: string, options?: { sheetName?: string[], inferSchemaLength?: number }): Promise<pl.DataFrame>`

Reads data from an Excel file and returns it as a Polars DataFrame.

- **`filePath`**: The path to the Excel file to be read.
- **`options`** _(optional)_:
  - **`sheetName`**: The name(s) of the sheet(s) to read. If not provided, the
    first sheet will be read.
  - **`inferSchemaLength`**: The number of rows to infer the schema from.
    Defaults to `100`.

**Key Behavior**:

- Empty cells in Excel are interpreted by ExcelJS as empty strings. Since `null`
  is the appropriate representation for missing values in DataFrames,
  `readExcel` automatically converts empty strings returned by ExcelJS into
  `null` values.

**Returns**: A `Promise` that resolves to a `pl.DataFrame` containing the data
from the Excel sheet.

**Throws**: Will throw an error if the specified worksheet is not found.

---

### `writeExcel(df: pl.DataFrame | pl.DataFrame[], filePath: string, options?: { sheetName?: string | string[]; includeHeader?: boolean; autofitColumns?: boolean; tableStyle?: TableStyle, header?: string, footer?: string, withWorkbook?: (workbook: ExcelJS.Workbook) => void }): Promise<void>`

Writes one or more Polars DataFrames to an Excel file, with optional styling and
formatting.

- **`df`**: A Polars DataFrame or an array of DataFrames to write to the file.
- **`filePath`**: The path to save the Excel file.
- **`options`** _(optional)_:
  - **`sheetName`**: The name(s) of the sheets to write to. Can be a string (for
    a single DataFrame) or an array of strings (for multiple DataFrames).
    Defaults to sequential names like `"Sheet1"`, `"Sheet2"`, etc.
  - **`includeHeader`**: Whether to include column headers. Defaults to `true`.
  - **`autofitColumns`**: Whether to auto-fit columns based on their content.
    Defaults to `true`.
  - **`tableStyle`**: A style theme for formatting the table in the Excel sheet.
  - **`header`**: - The header to add to the top of each page in the Excel file.
  - **`footer`**: - The footer to add to the bottom of each page in the Excel
    file.
  - **`withWorkbook`**: A function that receives the ExcelJS `workbook` object
    for further customization.

**Key Behavior**:

- If all DataFrames are empty, the function throws an error to prevent writing
  an Excel file with no meaningful content.
- If some DataFrames are empty, they are skipped, and a warning is logged for
  each skipped DataFrame. Non-empty DataFrames are written as expected.
- The `options.withWorkbook` parameter allows further customization of the
  workbook using the ExcelJS API.

**Returns**: A `Promise` that resolves when the file is successfully written.

---

## Requirements

- **Deno** (for Deno usage) or **Node.js** (for Node usage).
- `nodejs-polars` for Polars DataFrame support.
- `@tinkie101/exceljs-wrapper` as a wrapper for `ExcelJS`.

## Key Features

- Support for reading specific sheets from Excel files.
- Automatically converts empty cells interpreted by ExcelJS as empty strings to
  `null` for compatibility with Polars DataFrames.
- Optional column auto-fitting when writing DataFrames to Excel.
- Ability to write multiple DataFrames into separate worksheets.
- Skips empty DataFrames when writing, with warnings for each skipped sheet.
- Full compatibility with Polars DataFrames for seamless data manipulation.
- Options for customizing headers, autofit, and table styling when writing Excel
  files.
- Exposes the `workbook` object for further customization using the ExcelJS API
  through the `options.withWorkbook` parameter.
- Ensured strict type safety with TypeScript best practices.

---

## Key Changes

- Added support for writing multiple DataFrames into separate worksheets.
- Introduced dynamic handling of sheet names for multiple DataFrames.
- Empty cells in Excel are now automatically converted to `null` values for
  consistency in DataFrame processing.
- Skips empty DataFrames during writing, throwing an error only if all provided
  DataFrames are empty.
- Options for customizing headers, autofit, and table styling when writing Excel
  files.
- Exposes the `workbook` object for further customization using the ExcelJS API
  through the `options.withWorkbook` parameter.
- Ensured strict type safety with TypeScript best practices.

---

## License

GNU GENERAL PUBLIC LICENSE 3.0
