import type ExcelJS from "@tinkie101/exceljs-wrapper";

export type RowData = Record<
  string,
  string | number | boolean | null | undefined
>;

export type TableStyle = ExcelJS.TableStyleProperties["theme"];
