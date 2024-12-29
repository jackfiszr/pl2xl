import type pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import type { RowData, TableStyle } from "./types.ts";

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
