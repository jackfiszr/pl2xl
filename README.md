# pl2xl

A lightweight library for reading and writing Excel files as Polars DataFrames.\
`pl2xl` enables seamless integration between Polars and Excel, allowing you to:

- Import data from Excel files directly into a Polars DataFrame.
- Export Polars DataFrames back to Excel files, with optional Excel formatting.

## Installation

This library can be imported using the `jsr` import specifier and relies on the
`nodejs-polars` package.

### Importing the library in Deno

```typescript
import { readExcel, writeExcel } from "jsr:@jackfiszr/pl2xl@0.0.5";
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
console.log("Read DataFrame:", df);

// Modify the DataFrame by increasing the "Age" column by 1
const modifiedDf = df.withColumn(pl.col("Age").add(1).alias("Age"));

console.log("Modified DataFrame:", modifiedDf);

// Write the modified DataFrame to a new Excel file
await writeExcel(modifiedDf, "output.xlsx");
console.log("Modified DataFrame written to output.xlsx");
```

## API

### `readExcel(filePath: string, sheetName?: string): Promise<pl.DataFrame>`

Reads data from an Excel file and returns it as a Polars DataFrame.

- **`filePath`**: The path to the Excel file to be read.
- **`sheetName`** _(optional)_: The name of the sheet to read. If not provided,
  the first sheet will be read.

**Returns**: A `Promise` that resolves to a `pl.DataFrame` containing the data
from the Excel sheet.

**Throws**: Will throw an error if the specified worksheet is not found.

---

### `writeExcel(df: pl.DataFrame, filePath: string, options?: { sheetName?: string; includeHeader?: boolean; autofitColumns?: boolean; tableStyle?: TableStyle }): Promise<void>`

Writes a Polars DataFrame to an Excel file, with optional styling and
formatting.

- **`df`**: The Polars DataFrame to write to the file.
- **`filePath`**: The path to save the Excel file.
- **`options`** _(optional)_:
  - **`sheetName`**: The name of the sheet to write to. Defaults to `"Sheet1"`.
  - **`includeHeader`**: Whether to include column headers. Defaults to `true`.
  - **`autofitColumns`**: Whether to auto-fit columns based on their content.
    Defaults to `true`.
  - **`tableStyle`**: A style theme for formatting the table in the Excel sheet.

**Returns**: A `Promise` that resolves when the file is successfully written.

**Throws**: Will throw an error if the DataFrame is empty.

---

## Requirements

- **Deno** (for Deno usage) or **Node.js** (for Node usage).
- `nodejs-polars` for Polars DataFrame support.
- `@tinkie101/exceljs-wrapper` as a wrapper for `ExcelJS`.

## Key Features

- Support for reading specific sheets from Excel files.
- Optional column auto-fitting when writing DataFrames to Excel.
- Ability to style tables with Excel themes for enhanced readability.
- Full compatibility with Polars DataFrames for seamless data manipulation.

---

## Key Changes

- Added support for specifying a sheet name when reading Excel files.
- Introduced options for customizing headers, autofit, and table styling when
  writing Excel files.
- Leveraged ExcelJS's formatting capabilities to support future enhancements.
- Ensured strict type safety with TypeScript best practices.

---

## License

GNU GENERAL PUBLIC LICENSE 3.0
