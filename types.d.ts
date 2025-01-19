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

type ReplaceDataFrameWithExtended<T> = T extends originalPl.DataFrame<infer U>
  ? ExtendedDataFrame<U>
  : T extends originalPl.DataFrame<any> ? ExtendedDataFrame<any>
  : T extends (...args: infer Args) => infer Return ? (
      ...args: ReplaceDataFrameWithExtended<Args>
    ) => ReplaceDataFrameWithExtended<Return>
  : T extends object ? { [K in keyof T]: ReplaceDataFrameWithExtended<T[K]> }
  : T;

// @ts-ignore: Ignore recursive base type reference error
export interface ExtendedDataFrame<
  T extends Record<string, originalPl.Series> = any,
> extends ReplaceDataFrameWithExtended<originalPl.DataFrame<T>> {
  // Add your custom method here
  writeExcel(filePath: string, options?: WriteExcelOptions): Promise<void>;
}
