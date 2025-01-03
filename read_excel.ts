import pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import type { RowData } from "./types.ts";

/**
 * Reads an Excel file and returns its content as a DataFrame.
 *
 * @param filePath - The path to the Excel file to be read.
 * @param sheetName - Optional. The name of the sheet to read. If not provided, the first sheet will be read.
 * @returns A promise that resolves to a DataFrame containing the data from the specified Excel sheet.
 * @throws Will throw an error if the specified worksheet is not found in the Excel file.
 */
export async function readExcel(
  filePath: string,
  sheetName?: string,
): Promise<pl.DataFrame> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = sheetName
    ? workbook.getWorksheet(sheetName)
    : workbook.worksheets[0];

  if (!worksheet) {
    throw new Error(
      `Worksheet ${sheetName || "Sheet1"} not found in the Excel file.`,
    );
  }

  const jsonData: RowData[] = [];
  const headers: string[] = [];

  worksheet.eachRow((row, rowNumber) => {
    const rowData: RowData = {};

    row.eachCell((cell, colNumber) => {
      let cellValue = cell.value as string | number | boolean | null;

      if (typeof cellValue === "string" && cellValue.trim() === "") {
        // Replace empty string with null
        cellValue = null;
      }

      if (rowNumber === 1) {
        headers[colNumber - 1] = cellValue?.toString() || `Column${colNumber}`;
      } else {
        rowData[headers[colNumber - 1]] = cellValue;
      }
    });

    if (rowNumber > 1) {
      jsonData.push(rowData);
    }
  });

  return pl.DataFrame(jsonData);
}
