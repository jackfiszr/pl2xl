import ExcelJS from "@tinkie101/exceljs-wrapper";
import { exists } from "@std/fs";

export async function createTestExcelFile(
  filePath: string,
  data: { headers: string[]; rows: (string | number | boolean)[][] },
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  worksheet.addRow(data.headers);
  data.rows.forEach((row) => worksheet.addRow(row));

  await workbook.xlsx.writeFile(filePath);
}

export async function removeTestFile(filePath: string): Promise<void> {
  if (await exists(filePath)) {
    await Deno.remove(filePath);
  }
}

/**
 * Extracts rows from an ExcelJS worksheet and removes the first row and column.
 *
 * @param worksheet - The ExcelJS worksheet to extract rows from.
 * @returns A 2D array with the first row and first column removed.
 * @throws If the worksheet is undefined.
 */
export function getRows(
  worksheet: ExcelJS.Worksheet | undefined,
): ExcelJS.CellValue[][] {
  if (!worksheet) {
    throw new Error("The worksheet is undefined.");
  }

  return worksheet.getSheetValues()
    .slice(1) // Exclude the header row
    .map((row) => (Array.isArray(row) ? row.slice(1) : [])); // Exclude the empty first column
}
