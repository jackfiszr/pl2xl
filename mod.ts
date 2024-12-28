import type pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import { readExcel } from "./read_excel.ts";
import type { RowData } from "./types.ts";

/**
 * Writes a Polars DataFrame to an Excel file.
 * This function converts the DataFrame to JSON records and writes them to the specified file,
 * storing the data in a sheet named "Sheet1".
 *
 * @param df - The Polars DataFrame to write to an Excel file.
 * @param filePath - The path to save the Excel file.
 */
async function writeExcel(df: pl.DataFrame, filePath: string): Promise<void> {
  const rows: RowData[] = df.toRecords();

  if (rows.length === 0) {
    throw new Error("The DataFrame is empty. Nothing to write.");
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  const headers = Object.keys(rows[0]);
  worksheet.addRow(headers);

  rows.forEach((row) => {
    const values = headers.map((header) => row[header] ?? null);
    worksheet.addRow(values);
  });

  await workbook.xlsx.writeFile(filePath);
}

export { readExcel, writeExcel };
