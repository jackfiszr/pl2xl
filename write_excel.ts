import type pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import type { RowData, TableStyle } from "./types.ts";

/**
 * Writes a DataFrame to an Excel file.
 *
 * @param df - The DataFrame to write to the Excel file.
 * @param filePath - The path where the Excel file will be saved.
 * @param options - Optional settings for writing the Excel file.
 * @param options.sheetName - The name of the sheet in the Excel file. Defaults to "Sheet1".
 * @param options.includeHeader - Whether to include the DataFrame's column headers in the Excel file. Defaults to true.
 * @param options.autofitColumns - Whether to auto-fit the columns based on their content. Defaults to true.
 * @param options.tableStyle - The style to apply to the table in the Excel file.
 * @throws Will throw an error if the DataFrame is empty.
 * @returns A promise that resolves when the Excel file has been written.
 */
export async function writeExcel(
  df: pl.DataFrame,
  filePath: string,
  options: {
    sheetName?: string;
    includeHeader?: boolean;
    autofitColumns?: boolean;
    tableStyle?: TableStyle;
  } = {},
): Promise<void> {
  const {
    sheetName = "Sheet1",
    includeHeader = true,
    autofitColumns = true,
    tableStyle,
  } = options;

  const rows: RowData[] = df.toRecords();

  if (rows.length === 0) {
    throw new Error("The DataFrame is empty. Nothing to write.");
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);

  // Add headers if needed
  const headers = includeHeader ? Object.keys(rows[0]) : [];
  if (includeHeader) worksheet.addRow(headers);

  // Add data rows
  rows.forEach((row) => {
    const values = headers.map((header) => row[header] ?? null);
    worksheet.addRow(values);
  });

  // Apply table style if provided
  if (tableStyle && includeHeader) {
    const tableRange = {
      topLeft: worksheet.getCell(1, 1),
      bottomRight: worksheet.getCell(rows.length + 1, headers.length),
    };

    worksheet.addTable({
      name: `Table_${sheetName}`,
      ref: tableRange.topLeft.address,
      headerRow: true,
      style: { theme: tableStyle },
      columns: headers.map((header) => ({ name: header })),
      rows: rows.map((row) => headers.map((header) => row[header] ?? null)),
    });
  }

  // Auto-fit columns
  if (autofitColumns) {
    worksheet.columns.forEach((column) => {
      if (column.values) {
        column.width = Math.max(
          ...column.values.map((
            value,
          ) => (value ? value.toString().length : 10)),
        );
      } else {
        column.width = 10; // Default width
      }
    });
  }

  await workbook.xlsx.writeFile(filePath);
}
