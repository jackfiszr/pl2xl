# pl2xl

A lightweight library for reading and writing Excel files as Polars DataFrames.
`pl2xl` enables seamless integration between Polars and Excel, allowing you to
import data from Excel files directly into a Polars DataFrame and export
DataFrames back to Excel.

## Installation

This library can be imported using the `jsr` import specifier and relies on the
`nodejs-polars` package.

### Importing the library in Deno

```typescript
import { readExcel, writeExcel } from "jsr:@jackfiszr/pl2xl@0.0.3";
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
  "Name": ["Alice", "Bob", "Charlie"],
  "Age": [25, 30, 35],
  "City": ["New York", "Los Angeles", "Chicago"],
});

// Write the DataFrame to an Excel file
writeExcel(inputDf, "input.xlsx");

// Read the DataFrame back from the Excel file
const df = readExcel("input.xlsx");
console.log("Read DataFrame:", df);

// Modify the DataFrame by increasing the "Age" column by 1
const modifiedDf = df.withColumn(
  pl.col("Age").add(1).alias("Age"),
);

console.log("Modified DataFrame:", modifiedDf);

// Write the modified DataFrame to a new Excel file
writeExcel(modifiedDf, "output.xlsx");
console.log("Modified DataFrame written to output.xlsx");
```

## API

### `readExcel(filePath: string): pl.DataFrame`

Reads data from the first sheet of an Excel file and returns it as a Polars
DataFrame.

- `filePath`: The path to the Excel file to be read.

**Returns**: A `pl.DataFrame` containing the data from the Excel file.

### `writeExcel(df: pl.DataFrame, filePath: string): void`

Writes a Polars DataFrame to an Excel file.

- `df`: The Polars DataFrame to write to the file.
- `filePath`: The path to save the Excel file.

## Requirements

- Deno (for Deno usage) or Node.js (for Node usage)
- `nodejs-polars` and `xlsx` packages

## License

MIT License
