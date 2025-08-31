import ExcelJS from "@tinkie101/exceljs-wrapper";
import extenedPl from "./mod.ts";
import type {
  ExtendedDataFrame,
  ReadExcelOptions,
  RowData,
} from "./types.d.ts";

/**
 * Reads an Excel file and converts the specified worksheet to a DataFrame.
 *
 * @param filePath - The path to the Excel file to be read.
 * @param options - Optional settings for reading the Excel file.
 * @param options.sheetName - The name of the worksheet to read. Defaults to `null`.
 * @param options.inferSchemaLength - The number of rows to infer the schema from. Defaults to 100.
 * @returns A promise that resolves to a DataFrame containing the data from the specified worksheet.
 * @throws Will throw an error if the specified worksheet is not found in the Excel file.
 */
export async function readExcel(
  filePath: string,
  options: ReadExcelOptions = {},
): Promise<ExtendedDataFrame<any>> {
  const {
    sheetName = null,
    inferSchemaLength = 100,
  } = options;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = sheetName
    ? workbook.getWorksheet(sheetName[0])
    : workbook.worksheets[0];

  if (!worksheet) {
    throw new Error(
      `Worksheet ${sheetName} not found in the Excel file.`,
    );
  }

  const jsonData = worksheetToJson(worksheet);

  return extenedPl.DataFrame(jsonData, {
    inferSchemaLength,
  }) as ExtendedDataFrame<any>;
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
export function worksheetToJson(worksheet: ExcelJS.Worksheet): RowData[] {
  const jsonData: RowData[] = [];
  const headers: string[] = [];

  worksheet.eachRow((row, rowNumber) => {
    const rowData: RowData = {};

    if (rowNumber === 1) {
      // first row is header
      row.eachCell((cell, colNumber) => {
        let cellValue = cell.value as string | number | boolean | null;

        if (typeof cellValue === "string" && cellValue.trim() === "") {
          cellValue = null; // replace empty string with null
        }

        headers[colNumber - 1] = cellValue?.toString() || `Column${colNumber}`;
      });
    } else {
      // data rows
      headers.forEach((header, idx) => {
        // keep original column order
        const cellValue = row.getCell(idx + 1).value as
          | string
          | number
          | boolean
          | null
          | undefined;
        rowData[header] =
          (typeof cellValue === "string" && cellValue.trim() === "")
            ? null
            : (cellValue ?? null);
      });

      jsonData.push(rowData);
    }
  });

  return jsonData;
}
