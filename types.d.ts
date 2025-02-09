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
  string | number | bigint | boolean | null | undefined
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

/**
 * ExtendedDataFrame interface extends the original nodejs-polars DataFrame with `writeExcel` method
 * and overrides the methods that take/return DataFrame to take/return the ExtendedDataFrame instead.
 */
export interface ExtendedDataFrame extends originalPl.DataFrame {
  /**
   * Writes the DataFrame to an Excel file.
   * @param filePath - The path to the Excel file.
   * @param options - Optional settings for writing to Excel.
   * @returns A promise that resolves when the write operation is complete.
   */
  writeExcel(filePath: string, options?: WriteExcelOptions): Promise<void>;

  /**
   * Creates a deep copy of the DataFrame.
   * @returns A new instance of the DataFrame.
   */
  clone(): any;

  /**
   * Generates descriptive statistics of the DataFrame.
   * @returns The DataFrame with descriptive statistics.
   */
  describe(): any;

  /**
   * Drops the specified column(s) from the DataFrame.
   * @param name - The name of the column to drop.
   * @param names - Additional column names to drop.
   * @returns The DataFrame without the specified column(s).
   */
  drop<U extends string>(name: U): any;
  drop<const U extends string[]>(names: U): any;
  drop<U extends string, const V extends string[]>(name: U, ...names: V): any;

  /**
   * Drops rows with null values in the specified column(s).
   * @param column - The name of the column to check for null values.
   * @param columns - Additional column names to check for null values.
   * @returns The DataFrame without rows containing null values in the specified column(s).
   */
  dropNulls(column: string): any;
  dropNulls(columns: string[]): any;
  dropNulls(...columns: string[]): any;

  /**
   * Explodes the specified column(s) into multiple rows.
   * @param column - The column to explode.
   * @param columns - Additional columns to explode.
   * @returns The DataFrame with exploded columns.
   */
  explode(column: ExprOrString): any;
  explode(columns: ExprOrString[]): any;
  explode(column: ExprOrString, ...columns: ExprOrString[]): any;

  /**
   * Extends the DataFrame with another DataFrame.
   * @param other - The DataFrame to extend with.
   * @returns The extended DataFrame.
   */
  extend(other: any): any;

  /**
   * Fills null values in the DataFrame using the specified strategy.
   * @param strategy - The strategy to use for filling null values.
   * @returns The DataFrame with null values filled.
   */
  fillNull(strategy: originalPl.FillNullStrategy): any;

  /**
   * Filters the DataFrame based on the specified predicate.
   * @param predicate - The predicate to use for filtering.
   * @returns The filtered DataFrame.
   */
  filter(predicate: any): any;

  /**
   * Compares the DataFrame with another DataFrame for equality.
   * @param other - The DataFrame to compare with.
   * @param nullEqual - Whether to consider null values as equal.
   * @returns True if the DataFrames are equal, otherwise false.
   */
  frameEqual(other: any): boolean;
  frameEqual(other: any, nullEqual: boolean): boolean;

  /**
   * Returns the first `length` rows of the DataFrame.
   * @param length - The number of rows to return.
   * @returns The first `length` rows of the DataFrame.
   */
  head(length?: number): any;

  /**
   * Horizontally stacks the specified columns to the DataFrame.
   * @param columns - The columns to stack.
   * @param inPlace - Whether to perform the operation in place.
   * @returns The DataFrame with stacked columns.
   */
  hstack<U extends Record<string, originalPl.Series> = any>(
    columns: any,
  ): any;
  hstack<U extends originalPl.Series[]>(columns: U): any;
  hstack(columns: Array<originalPl.Series> | any): any;
  hstack(columns: Array<originalPl.Series> | any, inPlace?: boolean): void;

  /**
   * Interpolates missing values in the DataFrame.
   * @returns The DataFrame with interpolated values.
   */
  interpolate(): any;

  /**
   * Joins the DataFrame with another DataFrame based on the specified options.
   * @param other - The DataFrame to join with.
   * @param options - The options for the join operation.
   * @returns The joined DataFrame.
   */
  join(
    other: any,
    options:
      & { on: ValueOrArray<string> }
      & Omit<originalPl.JoinOptions, "leftOn" | "rightOn">,
  ): any;
  join(
    other: any,
    options:
      & { leftOn: ValueOrArray<string>; rightOn: ValueOrArray<string> }
      & Omit<originalPl.JoinOptions, "on">,
  ): any;
  join(other: any, options: { how: "cross"; suffix?: string }): any;

  /**
   * Performs an asof join with another DataFrame based on the specified options.
   * @param other - The DataFrame to join with.
   * @param options - The options for the asof join operation.
   * @returns The joined DataFrame.
   */
  joinAsof(
    other: any,
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
  ): any;

  /**
   * Limits the number of rows in the DataFrame.
   * @param length - The number of rows to limit to.
   * @returns The DataFrame with limited rows.
   */
  limit(length?: number): any;

  /**
   * Calculates the maximum value in the DataFrame.
   * @param axis - The axis to calculate the maximum value along.
   * @returns The DataFrame with the maximum value.
   */
  max(): any;
  max(axis: 0): any;
  max(axis: 1): originalPl.Series;

  /**
   * Calculates the mean value in the DataFrame.
   * @param axis - The axis to calculate the mean value along.
   * @param nullStrategy - The strategy to use for null values.
   * @returns The DataFrame with the mean value.
   */
  mean(): any;
  mean(axis: 0): any;
  mean(axis: 1): originalPl.Series;
  mean(axis: 1, nullStrategy?: "ignore" | "propagate"): originalPl.Series;

  /**
   * Calculates the median value in the DataFrame.
   * @returns The DataFrame with the median value.
   */
  median(): any;

  /**
   * Unpivots the DataFrame from wide to long format.
   * @param idVars - The columns to use as identifier variables.
   * @param valueVars - The columns to use as value variables.
   * @returns The unpivoted DataFrame.
   */
  unpivot(idVars: ColumnSelection, valueVars: ColumnSelection): any;

  /**
   * Calculates the minimum value in the DataFrame.
   * @param axis - The axis to calculate the minimum value along.
   * @returns The DataFrame with the minimum value.
   */
  min(): any;
  min(axis: 0): any;
  min(axis: 1): originalPl.Series;

  /**
   * Counts the number of null values in the DataFrame.
   * @returns The DataFrame with the count of null values.
   */
  nullCount(): any;

  /**
   * Partitions the DataFrame by the specified columns.
   * @param cols - The columns to partition by.
   * @param stable - Whether to maintain the order of rows.
   * @param includeKey - Whether to include the key in the partitioned DataFrames.
   * @param mapFn - A function to apply to each partitioned DataFrame.
   * @returns An array of partitioned DataFrames or the result of applying the map function.
   */
  partitionBy(
    cols: string | string[],
    stable?: boolean,
    includeKey?: boolean,
  ): any[];
  partitionBy<T>(
    cols: string | string[],
    stable: boolean,
    includeKey: boolean,
    mapFn: (df: any) => T,
  ): T[];

  /**
   * Pivots the DataFrame from long to wide format.
   * @param values - The columns to use as values.
   * @param options - The options for the pivot operation.
   * @returns The pivoted DataFrame.
   */
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
  ): any;
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
  }): any;

  /**
   * Calculates the quantile value in the DataFrame.
   * @param quantile - The quantile to calculate.
   * @returns The DataFrame with the quantile value.
   */
  quantile(quantile: number): any;

  /**
   * Rechunks the DataFrame to improve performance.
   * @returns The rechunked DataFrame.
   */
  rechunk(): any;

  /**
   * Renames columns in the DataFrame based on the specified mapping.
   * @param mapping - The mapping of old column names to new column names.
   * @returns The DataFrame with renamed columns.
   */
  rename<const U extends Partial<Record<string, string>>>(mapping: U): any;
  rename(mapping: Record<string, string>): any;

  /**
   * Selects the specified columns from the DataFrame.
   * @param columns - The columns to select.
   * @returns The DataFrame with selected columns.
   */
  select<U extends string>(...columns: U[]): any;
  select(...columns: ExprOrString[]): any;

  /**
   * Shifts the values in the DataFrame by the specified number of periods.
   * @param periods - The number of periods to shift.
   * @returns The DataFrame with shifted values.
   */
  shift(periods: number): any;
  shift({ periods }: { periods: number }): any;

  /**
   * Shifts the values in the DataFrame by the specified number of periods and fills the empty spaces with the specified value.
   * @param n - The number of periods to shift.
   * @param fillValue - The value to fill the empty spaces with.
   * @returns The DataFrame with shifted and filled values.
   */
  shiftAndFill(n: number, fillValue: number): any;
  shiftAndFill({ n, fillValue }: { n: number; fillValue: number }): any;

  /**
   * Shrinks the DataFrame to fit its contents.
   * @param inPlace - Whether to perform the operation in place.
   * @returns The DataFrame with reduced memory usage.
   */
  shrinkToFit(): any;
  shrinkToFit(inPlace: true): void;
  shrinkToFit({ inPlace }: { inPlace: true }): void;

  /**
   * Slices the DataFrame to include only the specified range of rows.
   * @param offset - The starting index of the slice.
   * @param length - The number of rows to include in the slice.
   * @returns The sliced DataFrame.
   */
  slice({ offset, length }: { offset: number; length: number }): any;
  slice(offset: number, length: number): any;

  /**
   * Sorts the DataFrame by the specified columns or expressions.
   * @param by - The columns or expressions to sort by.
   * @param descending - Whether to sort in descending order.
   * @param nullsLast - Whether to place null values at the end.
   * @param maintainOrder - Whether to maintain the order of rows.
   * @returns The sorted DataFrame.
   */
  sort(
    by: ColumnsOrExpr,
    descending?: boolean,
    nullsLast?: boolean,
    maintainOrder?: boolean,
  ): any;
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
  }): any;
  sort({
    by,
    descending,
    maintainOrder,
  }: {
    by: ColumnsOrExpr;
    descending?: boolean;
    nullsLast?: boolean;
    maintainOrder?: boolean;
  }): any;

  /**
   * Calculates the standard deviation in the DataFrame.
   * @returns The DataFrame with the standard deviation.
   */
  std(): any;

  /**
   * Calculates the sum of values in the DataFrame.
   * @param axis - The axis to calculate the sum along.
   * @param nullStrategy - The strategy to use for null values.
   * @returns The DataFrame with the sum of values.
   */
  sum(): any;
  sum(axis: 0): any;
  sum(axis: 1): originalPl.Series;
  sum(axis: 1, nullStrategy?: "ignore" | "propagate"): originalPl.Series;

  /**
   * Returns the last `length` rows of the DataFrame.
   * @param length - The number of rows to return.
   * @returns The last `length` rows of the DataFrame.
   */
  tail(length?: number): any;

  /**
   * Transposes the DataFrame, swapping rows and columns.
   * @param options - The options for the transpose operation.
   * @returns The transposed DataFrame.
   */
  transpose(options?: {
    includeHeader?: boolean;
    headerName?: string;
    columnNames?: Iterable<string>;
  }): any;

  /**
   * Returns a DataFrame with unique rows.
   * @param maintainOrder - Whether to maintain the order of rows.
   * @param subset - The columns to consider for uniqueness.
   * @param keep - Whether to keep the first or last occurrence of duplicate rows.
   * @returns The DataFrame with unique rows.
   */
  unique(
    maintainOrder?: boolean,
    subset?: ColumnSelection,
    keep?: "first" | "last",
  ): any;
  unique(opts: {
    maintainOrder?: boolean;
    subset?: ColumnSelection;
    keep?: "first" | "last";
  }): any;

  /**
   * Unnests the specified columns in the DataFrame.
   * @param names - The names of the columns to unnest.
   * @returns The DataFrame with unnested columns.
   */
  unnest(names: string | string[]): any;

  /**
   * Calculates the variance in the DataFrame.
   * @returns The DataFrame with the variance.
   */
  var(): any;

  /**
   * Vertically stacks another DataFrame to the current DataFrame.
   * @param df - The DataFrame to stack.
   * @returns The DataFrame with stacked rows.
   */
  vstack(df: any): any;

  /**
   * Adds a new column to the DataFrame.
   * @param column - The column to add.
   * @returns The DataFrame with the new column.
   */
  withColumn<
    SeriesTypeT extends originalPl.DataType,
    SeriesNameT extends string,
  >(
    column: originalPl.Series<SeriesTypeT, SeriesNameT>,
  ): any;
  withColumn(column: originalPl.Series | originalPl.Expr): any;

  /**
   * Adds multiple columns to the DataFrame.
   * @param columns - The columns to add.
   * @returns The DataFrame with the new columns.
   */
  withColumns(...columns: (originalPl.Expr | originalPl.Series)[]): any;

  /**
   * Renames a column in the DataFrame.
   * @param existingName - The current name of the column.
   * @param replacement - The new name of the column.
   * @returns The DataFrame with the renamed column.
   */
  withColumnRenamed<Existing extends string, New extends string>(
    existingName: Existing,
    replacement: New,
  ): any;
  withColumnRenamed(existing: string, replacement: string): any;
  withColumnRenamed<Existing extends string, New extends string>(opts: {
    existingName: Existing;
    replacement: New;
  }): any;
  withColumnRenamed(opts: { existing: string; replacement: string }): any;

  /**
   * Adds a row count column to the DataFrame.
   * @param name - The name of the row count column.
   * @returns The DataFrame with the row count column.
   */
  withRowCount(name?: string): any;

  /**
   * Filters the DataFrame based on the specified predicate.
   * @param predicate - The predicate to use for filtering.
   * @returns The filtered DataFrame.
   */
  where(predicate: any): any;

  /**
   * Upsamples the DataFrame based on the specified time column and interval.
   * @param timeColumn - The name of the time column.
   * @param every - The interval for upsampling.
   * @param by - Additional columns to group by.
   * @param maintainOrder - Whether to maintain the order of rows.
   * @returns The upsampled DataFrame.
   */
  upsample(
    timeColumn: string,
    every: string,
    by?: string | string[],
    maintainOrder?: boolean,
  ): any;
  upsample(opts: {
    timeColumn: string;
    every: string;
    by?: string | string[];
    maintainOrder?: boolean;
  }): any;
}
