import pl from "polars";
import xlsx from "xlsx";

/**
 * Reads an Excel file and returns its content as a Polars DataFrame.
 * This function takes the first sheet of the workbook and converts it to JSON
 * before loading it into a DataFrame.
 *
 * @param filePath - The path to the Excel file to be read.
 * @returns A Polars DataFrame containing the data from the first sheet of the Excel file.
 */
function readExcel(filePath: string): pl.DataFrame {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(sheet);
  return pl.DataFrame(jsonData);
}

/**
 * Writes a Polars DataFrame to an Excel file.
 * This function converts the DataFrame to JSON records and writes them to the
 * specified file, with the data stored in a sheet named "Sheet1".
 *
 * @param df - The Polars DataFrame to write to an Excel file.
 * @param filePath - The path to save the Excel file.
 */
function writeExcel(df: pl.DataFrame, filePath: string): void {
  const rows = df.toRecords();
  const newWorkbook = xlsx.utils.book_new();
  const newSheet = xlsx.utils.json_to_sheet(rows);
  xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");
  xlsx.writeFile(newWorkbook, filePath);
}

export { readExcel, writeExcel };
