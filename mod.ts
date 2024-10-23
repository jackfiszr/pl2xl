import pl from "polars";
import xlsx from "xlsx";

function readExcel(filePath: string): pl.DataFrame {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(sheet);
  return pl.DataFrame(jsonData);
}

function writeExcel(df: pl.DataFrame, filePath: string): void {
  const rows = df.toRecords();
  const newWorkbook = xlsx.utils.book_new();
  const newSheet = xlsx.utils.json_to_sheet(rows);
  xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");
  xlsx.writeFile(newWorkbook, filePath);
}

export { readExcel, writeExcel };
