# pl2xl <small>Extended Polars Library with Excel Support</small>

[![JSR](https://jsr.io/badges/@jackfiszr/pl2xl)](https://jsr.io/@jackfiszr/pl2xl)
[![JSR Score](https://jsr.io/badges/@jackfiszr/pl2xl/score)](https://jsr.io/@jackfiszr/pl2xl)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/pulse)
[![GitHub last commit](https://img.shields.io/github/last-commit/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/commits/main)
[![GitHub](https://img.shields.io/github/license/jackfiszr/pl2xl)](https://github.com/jackfiszr/pl2xl/blob/main/LICENSE)

## `pl.readExcel`

```typescript
pl.readExcel(
    filePath: string,
    options?: {
        sheetName?: string | null,
        inferSchemaLength?: number,
    },
) → ExtendedDataFrame
```

Reads an Excel file and converts the specified worksheet into an
`ExtendedDataFrame` (a `DataFrame` that has the `writeExcel` method).

### Parameters

- **filePath** (string):\
  The path to the Excel file to be read.

- **options** (object, optional):\
  A dictionary containing additional options:
  - **sheetName** (string | null, optional):\
    The name of the worksheet to read. Defaults to the first worksheet if not
    specified.
  - **inferSchemaLength** (number, optional):\
    The number of rows to infer the schema from. Defaults to 100.

### Returns

- **ExtendedDataFrame**:\
  A DataFrame containing the data from the specified worksheet.

### Example

```typescript
import pl from "jsr:@jackfiszr/pl2xl@0.1.0";

const df = await pl.readExcel("data.xlsx", { sheetName: "Sheet1" });
console.log(df.toString());
```

---

## `pl.DataFrame.writeExcel`

```typescript
pl.DataFrame.writeExcel(
    filePath: string,
    options?: {
        sheetName?: string | string[],
        includeHeader?: boolean,
        autofitColumns?: boolean,
        tableStyle?: string,
        header?: string,
        footer?: string,
        withWorkbook?: (workbook: ExcelJS.Workbook) => void,
    },
) → Promise<void>
```

Writes the dataframe to an Excel `xlsx` file.

### Parameters

- **filePath** (string):\
  The path where the Excel file will be saved.

- **options** (object, optional):\
  A dictionary containing additional options:
  - **sheetName** (string | string[], optional):\
    Name(s) of the worksheet(s). Defaults to `Sheet1`, `Sheet2`, etc.
  - **includeHeader** (boolean, optional):\
    Whether to include column headers in the Excel file. Defaults to `true`.
  - **autofitColumns** (boolean, optional):\
    Whether to auto-fit the columns based on content. Defaults to `true`.
  - **tableStyle** (string, optional):\
    The style to apply to the table(s) in the Excel file.
  - **header** (string, optional):\
    The header text to add at the top of each worksheet.
  - **footer** (string, optional):\
    The footer text to add at the bottom of each worksheet.
  - **withWorkbook** (function, optional):\
    A callback function that receives the `ExcelJS.Workbook` instance for
    additional customization.

### Returns

- **Promise<void>**:\
  A promise that resolves when the Excel file has been written.

### Example

```typescript
import pl from "jsr:@jackfiszr/pl2xl@0.1.0";

const df = pl.DataFrame({
  Name: ["Alice", "Bob"],
  Age: [25, 30],
});

await df.writeExcel("output.xlsx", {
  sheetName: "People",
  includeHeader: true,
  autofitColumns: true,
});
```

## `pl.writeExcel` <small>(for writing multiple dataframes to separate worksheets)</small>

```typescript
pl.writeExcel(
    df: ExtendedDataFrame | ExtendedDataFrame[],
    filePath: string,
    options?: {
        sheetName?: string | string[],
        includeHeader?: boolean,
        autofitColumns?: boolean,
        tableStyle?: string,
        header?: string,
        footer?: string,
        withWorkbook?: (workbook: ExcelJS.Workbook) => void,
    },
) → Promise<void>
```

Writes one or more Polars `ExtendedDataFrame` objects to an Excel file.

### Parameters

Has one additional parameter that is the first parameter:

- **df** (ExtendedDataFrame | ExtendedDataFrame[]):\
  The DataFrame(s) to write to the Excel file.

### Returns

- **Promise<void>**:\
  A promise that resolves when the Excel file has been written.

### Example

```typescript
import pl from "jsr:@jackfiszr/pl2xl@0.1.0";

const df1 = pl.DataFrame({
  Name: ["Alice", "Bob"],
  Age: [25, 30],
});

const df2 = pl.DataFrame({
  Name: ["Cat", "Dog"],
  Age: [14, 10],
});

await pl.writeExcel([df1, df2], "output.xlsx", {
  sheetName: ["People", "Animals"],
});
```

---

## nodejs-polars

For the core functionality of the library, please refer to the official
[nodejs-polars](https://pola-rs.github.io/nodejs-polars/index.html)
documentation.

## License

This library is open-source and distributed under the GNU GENERAL PUBLIC LICENSE
3.0.
