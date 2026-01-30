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

  readExcel: <T extends Record<string, originalPl.Series<any, string>>>(
    filePath: string,
    options?: ReadExcelOptions,
  ) => Promise<ExtendedDataFrame<T>>;

  writeExcel: <T extends Record<string, originalPl.Series<any, string>>>(
    df:
      | ExtendedDataFrame<T>
      | ExtendedDataFrame<T>[]
      | originalPl.DataFrame<any>
      | originalPl.DataFrame<any>[],
    filePath: string,
    options?: WriteExcelOptions,
  ) => Promise<void>;
};
