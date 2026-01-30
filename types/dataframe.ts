/**
 * Represents a single row of data as a record of column name to value.
 */
export type RowData = Record<
  string,
  string | number | bigint | boolean | null | undefined
>;
