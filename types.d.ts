import type ExcelJS from "@tinkie101/exceljs-wrapper";
import type originalPl from "polars";
import type { JsToDtype } from "polars/datatypes.ts";
import type {
  ColumnSelection,
  ColumnsOrExpr,
  ExprOrString,
  Simplify,
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

export interface ExtendedDataFrame<
  T extends Record<string, originalPl.Series> = any,
> extends originalPl.DataFrame<T> {
  writeExcel: (filePath: string, options?: WriteExcelOptions) => Promise<void>;

  clone(): ExtendedDataFrame<T>;
  describe(): this;
  drop<U extends string>(name: U): ExtendedDataFrame<Simplify<Omit<T, U>>>;
  drop<const U extends string[]>(
    names: U,
  ): ExtendedDataFrame<Simplify<Omit<T, U[number]>>>;
  drop<U extends string, const V extends string[]>(
    name: U,
    ...names: V
  ): ExtendedDataFrame<Simplify<Omit<T, U | V[number]>>>;
  dropNulls(column: keyof T): ExtendedDataFrame<T>;
  dropNulls(columns: (keyof T)[]): ExtendedDataFrame<T>;
  dropNulls(...columns: (keyof T)[]): ExtendedDataFrame<T>;
  explode(column: ExprOrString): this;
  explode(columns: ExprOrString[]): this;
  explode(column: ExprOrString, ...columns: ExprOrString[]): this;
  extend(other: ExtendedDataFrame<T>): ExtendedDataFrame<T>;
  fillNull(strategy: originalPl.FillNullStrategy): ExtendedDataFrame<T>;
  filter(predicate: any): ExtendedDataFrame<T>;
  frameEqual(other: ExtendedDataFrame): boolean;
  frameEqual(other: ExtendedDataFrame, nullEqual: boolean): boolean;
  head(length?: number): ExtendedDataFrame<T>;
  hstack<U extends Record<string, originalPl.Series> = any>(
    columns: ExtendedDataFrame<U>,
  ): ExtendedDataFrame<Simplify<T & U>>;
  hstack<U extends originalPl.Series[]>(columns: U): ExtendedDataFrame<
    Simplify<
      & T
      & {
        [K in U[number] as K["name"]]: K;
      }
    >
  >;
  hstack(
    columns: Array<originalPl.Series> | ExtendedDataFrame,
  ): this;
  hstack(
    columns: Array<originalPl.Series> | ExtendedDataFrame,
    inPlace?: boolean,
  ): void;
  interpolate(): ExtendedDataFrame<T>;
  join(
    other: ExtendedDataFrame,
    options: {
      on: ValueOrArray<string>;
    } & Omit<originalPl.JoinOptions, "leftOn" | "rightOn">,
  ): this;
  join(
    other: ExtendedDataFrame,
    options: {
      leftOn: ValueOrArray<string>;
      rightOn: ValueOrArray<string>;
    } & Omit<originalPl.JoinOptions, "on">,
  ): this;
  join(other: ExtendedDataFrame, options: {
    how: "cross";
    suffix?: string;
  }): this;
  joinAsof(other: ExtendedDataFrame, options: {
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
  }): this;
  // lazy(): LazyDataFrame;
  limit(length?: number): ExtendedDataFrame<T>;
  max(): ExtendedDataFrame<T>;
  max(axis: 0): ExtendedDataFrame<T>;
  mean(): ExtendedDataFrame<T>;
  mean(axis: 0): ExtendedDataFrame<T>;
  median(): ExtendedDataFrame<T>;
  unpivot(
    idVars: ColumnSelection,
    valueVars: ColumnSelection,
  ): this;
  min(): ExtendedDataFrame<T>;
  min(axis: 0): ExtendedDataFrame<T>;
  nullCount(): ExtendedDataFrame<
    {
      [K in keyof T]: originalPl.Series<JsToDtype<number>, K & string>;
    }
  >;
  partitionBy(
    cols: string | string[],
    stable?: boolean,
    includeKey?: boolean,
  ): ExtendedDataFrame<T>[];
  partitionBy<T>(
    cols: string | string[],
    stable: boolean,
    includeKey: boolean,
    mapFn: (df: ExtendedDataFrame) => T,
  ): T[];
  pivot(values: string | string[], options: {
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
  quantile(quantile: number): ExtendedDataFrame<T>;
  rechunk(): ExtendedDataFrame<T>;
  rename<const U extends Partial<Record<keyof T, string>>>(
    mapping: U,
  ): ExtendedDataFrame<
    {
      [K in keyof T as U[K] extends string ? U[K] : K]: T[K];
    }
  >;
  rename(mapping: Record<string, string>): this;
  select<U extends keyof T>(...columns: U[]): ExtendedDataFrame<
    {
      [P in U]: T[P];
    }
  >;
  select(...columns: ExprOrString[]): ExtendedDataFrame<T>;
  shift(periods: number): ExtendedDataFrame<T>;
  shift({ periods }: {
    periods: number;
  }): ExtendedDataFrame<T>;
  shiftAndFill(n: number, fillValue: number): ExtendedDataFrame<T>;
  shiftAndFill({ n, fillValue }: {
    n: number;
    fillValue: number;
  }): ExtendedDataFrame<T>;
  shrinkToFit(): ExtendedDataFrame<T>;
  slice({ offset, length }: {
    offset: number;
    length: number;
  }): ExtendedDataFrame<T>;
  slice(offset: number, length: number): ExtendedDataFrame<T>;
  sort(
    by: ColumnsOrExpr,
    descending?: boolean,
    nullsLast?: boolean,
    maintainOrder?: boolean,
  ): ExtendedDataFrame<T>;
  sort({ by, reverse, maintainOrder }: {
    by: ColumnsOrExpr;
    reverse?: boolean;
    nullsLast?: boolean;
    maintainOrder?: boolean;
  }): ExtendedDataFrame<T>;
  sort({ by, descending, maintainOrder }: {
    by: ColumnsOrExpr;
    descending?: boolean;
    nullsLast?: boolean;
    maintainOrder?: boolean;
  }): ExtendedDataFrame<T>;
  std(): ExtendedDataFrame<T>;
  sum(): ExtendedDataFrame<T>;
  sum(axis: 0): ExtendedDataFrame<T>;
  tail(length?: number): ExtendedDataFrame<T>;
  transpose(options?: {
    includeHeader?: boolean;
    headerName?: string;
    columnNames?: Iterable<string>;
  }): this;
  unique(
    maintainOrder?: boolean,
    subset?: ColumnSelection,
    keep?: "first" | "last",
  ): ExtendedDataFrame<T>;
  unique(opts: {
    maintainOrder?: boolean;
    subset?: ColumnSelection;
    keep?: "first" | "last";
  }): ExtendedDataFrame<T>;
  unnest(names: string | string[]): this;
  var(): ExtendedDataFrame<T>;
  vstack(df: ExtendedDataFrame<T>): ExtendedDataFrame<T>;
  withColumn<
    SeriesTypeT extends originalPl.DataType,
    SeriesNameT extends string,
  >(
    column: originalPl.Series<SeriesTypeT, SeriesNameT>,
  ): ExtendedDataFrame<
    Simplify<
      & T
      & {
        [K in SeriesNameT]: originalPl.Series<SeriesTypeT, SeriesNameT>;
      }
    >
  >;
  withColumn(column: originalPl.Series | originalPl.Expr): this;
  withColumns(
    ...columns: (originalPl.Expr | originalPl.Series)[]
  ): this;
  withColumnRenamed<Existing extends keyof T, New extends string>(
    existingName: Existing,
    replacement: New,
  ): ExtendedDataFrame<
    {
      [K in keyof T as K extends Existing ? New : K]: T[K];
    }
  >;
  withColumnRenamed(existing: string, replacement: string): this;
  withColumnRenamed<Existing extends keyof T, New extends string>(opts: {
    existingName: Existing;
    replacement: New;
  }): ExtendedDataFrame<
    {
      [K in keyof T as K extends Existing ? New : K]: T[K];
    }
  >;
  withColumnRenamed(opts: {
    existing: string;
    replacement: string;
  }): this;
  withRowCount(name?: string): this;
  where(predicate: any): ExtendedDataFrame<T>;
  upsample(
    timeColumn: string,
    every: string,
    by?: string | string[],
    maintainOrder?: boolean,
  ): ExtendedDataFrame<T>;
  upsample(opts: {
    timeColumn: string;
    every: string;
    by?: string | string[];
    maintainOrder?: boolean;
  }): ExtendedDataFrame<T>;
}
