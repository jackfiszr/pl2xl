import pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";

type RowData = Record<string, string | number | boolean | null | undefined>;

/**
 * Reads an Excel file and returns its content as a Polars DataFrame.
 * This function loads the first sheet of the workbook and converts it into a DataFrame.
 *
 * @param filePath - The path to the Excel file to be read.
 * @returns A Promise resolving to a Polars DataFrame containing the data from the first sheet.
 */
async function readExcel(filePath: string): Promise<pl.DataFrame> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("No sheets found in the Excel file.");
  }

  const jsonData: RowData[] = [];
  const headers: string[] = [];

  worksheet.eachRow((row, rowNumber) => {
    const rowData: RowData = {};

    row.eachCell((cell, colNumber) => {
      const cellValue = cell.value as string | number | boolean | null;

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
