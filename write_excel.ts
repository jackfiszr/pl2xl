import type pl from "polars";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import type { RowData, TableStyle } from "./types.ts";

/**
 * Writes one or more DataFrames to an Excel file, each in its own worksheet.
 *
 * @param df - The DataFrame or array of DataFrames to write to the Excel file.
 * @param filePath - The path where the Excel file will be saved.
 * @param options - Optional settings for writing the Excel file.
 * @param options.sheetName - The name(s) of the sheets in the Excel file. Defaults to ["Sheet1", "Sheet2", ...].
 * @param options.includeHeader - Whether to include the DataFrame's column headers in the Excel file. Defaults to true.
 * @param options.autofitColumns - Whether to auto-fit the columns based on their content. Defaults to true.
 * @param options.tableStyle - The style to apply to the tables in the Excel file.
 * @param options.header - The header to add to the top of each page in the Excel file.
 * @param options.footer - The footer to add to the bottom of each page in the Excel file.
 * @param options.withWorkbook - A callback function that receives the workbook instance for further customization.
 * @throws Will throw an error if all the DataFrames are empty.
 * @returns A promise that resolves when the Excel file has been written.
 */
export async function writeExcel(
  df: pl.DataFrame | pl.DataFrame[],
  filePath: string,
  options: {
    sheetName?: string | string[];
    includeHeader?: boolean;
    autofitColumns?: boolean;
    tableStyle?: TableStyle;
    header?: string;
    footer?: string;
    withWorkbook?: (workbook: ExcelJS.Workbook) => void;
  } = {},
): Promise<void> {
  const {
    sheetName = "Sheet1",
    includeHeader = true,
    autofitColumns = true,
    tableStyle,
    header,
    footer,
    withWorkbook,
  } = options;

  const dataframes = Array.isArray(df) ? df : [df];
  const sheetNames = Array.isArray(sheetName) ? sheetName : [sheetName];

  if (sheetNames.length < dataframes.length) {
    throw new Error("Not enough sheet names provided for the DataFrames.");
  }

  // Check if all DataFrames are empty
  const allEmpty = dataframes.every((df) => df.height === 0);
  if (allEmpty) {
    if (dataframes.length === 1) {
      throw new Error("The DataFrame is empty. Nothing to write.");
    } else {
      throw new Error("All provided DataFrames are empty. Nothing to write.");
    }
  }

  const workbook = new ExcelJS.Workbook();

  for (let i = 0; i < dataframes.length; i++) {
    const currentDf = dataframes[i];
    const currentSheetName = sheetNames[i] || `Sheet${i + 1}`;

    const rows: RowData[] = currentDf.toRecords();

    // Skip writing empty DataFrames but don't throw
    if (rows.length === 0) {
      console.warn(
        `DataFrame at index ${i} is empty. Skipping this worksheet.`,
      );
      continue;
    }

    const worksheet = workbook.addWorksheet(currentSheetName);

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
        name: `Table_${currentSheetName}`,
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
            ...column.values
              .slice(1) // Skip the metadata slot
              .map((value) => (value ? value.toString().length : 10)),
          );
        } else {
          column.width = 10; // Default width
        }
      });
    }

    if (header) {
      worksheet.headerFooter.oddHeader = header;
      worksheet.headerFooter.evenHeader = header;
    }
    if (footer) {
      worksheet.headerFooter.oddFooter = footer;
      worksheet.headerFooter.evenFooter = footer;
    }
  }

  if (withWorkbook) {
    withWorkbook(workbook);
  }

  await workbook.xlsx.writeFile(filePath);
}
