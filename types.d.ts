import type ExcelJS from "@tinkie101/exceljs-wrapper";
import type originalPl from "polars";
import type {
  ColumnSelection,
  ColumnsOrExpr,
  ExprOrString,
  ValueOrArray,
} from "polars/utils.ts";

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

export interface ExtendedDataFrame extends originalPl.DataFrame {
  writeExcel: (filePath: string, options?: WriteExcelOptions) => Promise<void>;

  clone: () => this;

  describe: () => this;

  drop<U extends string>(name: U): this;
  drop<const U extends string[]>(names: U): this;
  drop<U extends string, const V extends string[]>(name: U, ...names: V): any;

  dropNulls(column: string): this;
  dropNulls(columns: string[]): this;
  dropNulls(...columns: string[]): this;

  explode(column: ExprOrString): this;
  explode(columns: ExprOrString[]): this;
  explode(column: ExprOrString, ...columns: ExprOrString[]): this;

  extend(other: this): this;

  fillNull(strategy: originalPl.FillNullStrategy): this;

  filter(predicate: any): this;

  frameEqual(other: this): boolean;
  frameEqual(other: this, nullEqual: boolean): boolean;

  head(length?: number): this;

  hstack<U extends Record<string, originalPl.Series> = any>(
    columns: this,
  ): this;
  hstack<U extends originalPl.Series[]>(columns: U): this;
  hstack(
    columns: Array<originalPl.Series> | this,
  ): this;
  hstack(
    columns: Array<originalPl.Series> | this,
    inPlace?: boolean,
  ): void;

  interpolate(): this;

  join(
    other: this,
    options:
      & { on: ValueOrArray<string> }
      & Omit<originalPl.JoinOptions, "leftOn" | "rightOn">,
  ): this;
  join(
    other: this,
    options:
      & { leftOn: ValueOrArray<string>; rightOn: ValueOrArray<string> }
      & Omit<originalPl.JoinOptions, "on">,
  ): this;
  join(
    other: this,
    options: { how: "cross"; suffix?: string },
  ): this;

  joinAsof(
    other: this,
    options: {
      leftOn?: string;
      rightOn?: string;
      on?: string;
      byLeft?: string | string[];
      byRight?: string | string[];
      by?: string | string[];
      strategy?: "backward" | "forward" | "nearest";
      suffix?: string;
      tolerance?: number | string;
      allowParallel?: boolean;
      forceParallel?: boolean;
    },
  ): this;

  // lazy(): LazyDataFrame;

  limit(length?: number): this;

  max(): this;
  max(axis: 0): this;
  max(axis: 1): originalPl.Series;

  mean(): this;
  mean(axis: 0): this;
  mean(axis: 1): originalPl.Series;
  mean(axis: 1, nullStrategy?: "ignore" | "propagate"): originalPl.Series;

  median(): this;

  unpivot(idVars: ColumnSelection, valueVars: ColumnSelection): this;

  min(): this;
  min(axis: 0): this;
  min(axis: 1): originalPl.Series;

  nullCount(): this;

  partitionBy(
    cols: string | string[],
    stable?: boolean,
    includeKey?: boolean,
  ): this[];
  partitionBy<T>(
    cols: string | string[],
    stable: boolean,
    includeKey: boolean,
    mapFn: (df: this) => T,
  ): T[];

  pivot(
    values: string | string[],
    options: {
      index: string | string[];
      on: string | string[];
      aggregateFunc?:
        | "sum"
        | "max"
        | "min"
        | "mean"
        | "median"
        | "first"
        | "last"
        | "count"
        | originalPl.Expr;
      maintainOrder?: boolean;
      sortColumns?: boolean;
      separator?: string;
    },
  ): this;
  pivot(options: {
    values: string | string[];
    index: string | string[];
    on: string | string[];
    aggregateFunc?:
      | "sum"
      | "max"
      | "min"
      | "mean"
      | "median"
      | "first"
      | "last"
      | "count"
      | originalPl.Expr;
    maintainOrder?: boolean;
    sortColumns?: boolean;
    separator?: string;
  }): this;

  quantile(quantile: number): this;

  rechunk(): this;

  rename<const U extends Partial<Record<string, string>>>(
    mapping: U,
  ): this;
  rename(mapping: Record<string, string>): this;

  select<U extends string>(...columns: U[]): this;
  select(...columns: ExprOrString[]): this;

  shift(periods: number): this;
  shift({ periods }: { periods: number }): this;

  shiftAndFill(n: number, fillValue: number): this;
  shiftAndFill({
    n,
    fillValue,
  }: { n: number; fillValue: number }): this;

  shrinkToFit(): this;
  shrinkToFit(inPlace: true): void;
  shrinkToFit({ inPlace }: { inPlace: true }): void;

  slice({ offset, length }: { offset: number; length: number }): this;
  slice(offset: number, length: number): this;

  sort(
    by: ColumnsOrExpr,
    descending?: boolean,
    nullsLast?: boolean,
    maintainOrder?: boolean,
  ): this;
  sort({
    by,
    reverse, // deprecated
    maintainOrder,
  }: {
    by: ColumnsOrExpr;
    /** @deprecated *since 0.16.0* @use descending */
    reverse?: boolean; // deprecated
    nullsLast?: boolean;
    maintainOrder?: boolean;
  }): this;
  sort({
    by,
    descending,
    maintainOrder,
  }: {
    by: ColumnsOrExpr;
    descending?: boolean;
    nullsLast?: boolean;
    maintainOrder?: boolean;
  }): this;

  std(): this;

  sum(): this;
  sum(axis: 0): this;
  sum(axis: 1): originalPl.Series;
  sum(axis: 1, nullStrategy?: "ignore" | "propagate"): originalPl.Series;

  tail(length?: number): this;

  transpose(options?: {
    includeHeader?: boolean;
    headerName?: string;
    columnNames?: Iterable<string>;
  }): this;

  unique(
    maintainOrder?: boolean,
    subset?: ColumnSelection,
    keep?: "first" | "last",
  ): this;
  unique(opts: {
    maintainOrder?: boolean;
    subset?: ColumnSelection;
    keep?: "first" | "last";
  }): this;

  unnest(names: string | string[]): this;

  var(): this;

  vstack(df: this): this;

  withColumn<
    SeriesTypeT extends originalPl.DataType,
    SeriesNameT extends string,
  >(
    column: originalPl.Series<SeriesTypeT, SeriesNameT>,
  ): this;
  withColumn(column: originalPl.Series | originalPl.Expr): this;

  withColumns(...columns: (originalPl.Expr | originalPl.Series)[]): this;

  withColumnRenamed<Existing extends string, New extends string>(
    existingName: Existing,
    replacement: New,
  ): this;
  withColumnRenamed(existing: string, replacement: string): this;
  withColumnRenamed<Existing extends string, New extends string>(opts: {
    existingName: Existing;
    replacement: New;
  }): this;
  withColumnRenamed(opts: { existing: string; replacement: string }): this;

  withRowCount(name?: string): this;

  where(predicate: any): this;

  upsample(
    timeColumn: string,
    every: string,
    by?: string | string[],
    maintainOrder?: boolean,
  ): this;
  upsample(opts: {
    timeColumn: string;
    every: string;
    by?: string | string[];
    maintainOrder?: boolean;
  }): this;
}
