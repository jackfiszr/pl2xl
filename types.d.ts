import type ExcelJS from "@tinkie101/exceljs-wrapper";
import type originalPl from "polars";

export type RowData = Record<
  string,
  string | number | boolean | null | undefined
>;

export type TableStyle = ExcelJS.TableStyleProperties["theme"];

type ExcelSpreadsheetEngine = "exceljs" | "xslx";
type SchemaDict = Record<string, unknown>;

export interface ReadExcelOptions {
  sheetId?: number | null;
  sheetName?: string[] | [string] | null;
  engine?: ExcelSpreadsheetEngine;
  engineOptions?: Record<string, unknown>;
  readOptions?: Record<string, unknown>;
  hasHeader?: boolean;
  columns?: number[] | string[] | null;
  schemaOverrides?: SchemaDict | null;
  inferSchemaLength?: number;
  includeFilePaths?: string | null;
  dropEmptyRows?: boolean;
  dropEmptyCols?: boolean;
  raiseIfEmpty?: boolean;
}

export interface WriteExcelOptions {
  sheetName?: string | string[];
  includeHeader?: boolean;
  autofitColumns?: boolean;
  tableStyle?: TableStyle;
  header?: string;
  footer?: string;
  withWorkbook?: (workbook: ExcelJS.Workbook) => void;
}

// deno-fmt-ignore
export interface ExtendedDataFrame extends originalPl.DataFrame {
  writeExcel: (filePath: string, options?: WriteExcelOptions) => Promise<void>;

  clone: () => ExtendedDataFrame;
  describe: () => ExtendedDataFrame;

  withColumn: (columns: originalPl.Series | originalPl.Expr) => ExtendedDataFrame;
  withColumns: (...columns: (originalPl.Expr | originalPl.Series)[]) => ExtendedDataFrame;
}
