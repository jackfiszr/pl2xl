import pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import type { ReadExcelOptions, RowData } from "./types.ts";

/**
 * Reads an Excel file and converts the specified worksheet to a DataFrame.
 *
 * @param filePath - The path to the Excel file to be read.
 * @param options - Optional settings for reading the Excel file.
 * @param options.sheetName - The name of the worksheet to read. Defaults to "Sheet1".
 * @param options.inferSchemaLength - The number of rows to infer the schema from. Defaults to 100.
 * @returns A promise that resolves to a DataFrame containing the data from the specified worksheet.
 * @throws Will throw an error if the specified worksheet is not found in the Excel file.
 */
export async function readExcel(
  filePath: string,
  options: ReadExcelOptions = {},
): Promise<pl.DataFrame> {
  const {
    sheetName = ["Sheet1"],
    inferSchemaLength = 100,
  } = options;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = sheetName
    ? workbook.getWorksheet(sheetName[0])
    : workbook.worksheets[0];

  if (!worksheet) {
    throw new Error(
      `Worksheet ${sheetName || "Sheet1"} not found in the Excel file.`,
    );
  }

  const jsonData = worksheetToJson(worksheet);

  return pl.DataFrame(jsonData, { inferSchemaLength });
}

/**
 * Converts an Excel worksheet to a JSON array of row data.
 *
 * @param {ExcelJS.Worksheet} worksheet - The worksheet to convert.
 * @returns {RowData[]} An array of row data objects, where each object represents a row in the worksheet.
 *
 * @remarks
 * - The first row of the worksheet is assumed to be the header row.
 * - Each cell value in the header row is used as the key for the corresponding column in the row data objects.
 * - Empty string cell values are replaced with `null`.
 * - If a header cell is empty, a default column name in the format `Column{colNumber}` is used.
 */
function worksheetToJson(worksheet: ExcelJS.Worksheet): RowData[] {
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

  return jsonData;
}
