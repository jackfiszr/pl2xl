import ExcelJS from "jsr:@tinkie101/exceljs-wrapper@^1.0.2";
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
