import type originalPl from "polars";
import type { ExtendedDataFrame } from "./extended_dataframe.ts";
import type { ReadExcelOptions, WriteExcelOptions } from "./excel.ts";

/**
 * Type representing the extended Polars object, which includes:
 * - Wrapped DataFrame with `writeExcel`
 * - Additional functions like `readExcel` and `writeExcel`
 */
export type ExtendedPolars = Omit<typeof originalPl, "DataFrame"> & {
  DataFrame: <
    T extends Record<string, originalPl.Series<any, string>>,
  >(
    ...args: Parameters<typeof originalPl.DataFrame>
  ) => ExtendedDataFrame<T>;
  readExcel: (
    filePath: string,
    options?: ReadExcelOptions,
  ) => Promise<ExtendedDataFrame<any>>;
  writeExcel: (
    df:
      | ExtendedDataFrame<any>
      | ExtendedDataFrame<any>[]
      | import("polars").DataFrame<any>
      | import("polars").DataFrame<any>[],
    filePath: string,
    options?: WriteExcelOptions,
  ) => Promise<void>;
};
