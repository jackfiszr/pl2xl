import originalPl from "polars";
import { readExcel } from "./read_excel.ts";
import { writeExcel } from "./write_excel.ts";
import type {
  ExtendedDataFrame,
  ExtendedPolars,
  ReadExcelOptions,
  WriteExcelOptions,
} from "./types.d.ts";

export type * from "./types.d.ts";

/**
 * A wrapper function for the original DataFrame constructor from the `nodejs-polars` library.
 * This function ensures that the `writeExcel` method is available on the DataFrame instance.
 *
 * @param {...Parameters<typeof originalPl.DataFrame>} args - The arguments to be passed to the original DataFrame constructor.
 * @returns {ExtendedDataFrame} - The extended DataFrame instance with `writeExcel` method.
 *
 * @remarks
 * This function dynamically adds the `writeExcel` method to the DataFrame instance if it doesn't already exist.
 * It also extends various DataFrame methods to ensure that any new DataFrame returned by these methods
 * also includes the `writeExcel` method.
 *
 * The methods extended include:
 * - clone
 * - describe
 * - drop
 * - dropNulls
 * - explode
 * - extend
 * - fillNull
 * - filter
 * - frameEqual
 * - head
 * - hstack
 * - interpolate
 * - join
 * - joinAsof
 * - limit
 * - max
 * - mean
 * - median
 * - unpivot
 * - min
 * - nullCount
 * - partitionBy
 * - pivot
 * - quantile
 * - rechunk
 * - rename
 * - select
 * - shift
 * - shiftAndFill
 * - shrinkToFit
 * - slice
 * - sort
 * - std
 * - sum
 * - tail
 * - transpose
 * - unique
 * - unnest
 * - var
 * - vstack
 * - withColumn
 * - withColumns
 * - withColumnRenamed
 * - withRowCount
 * - where
 * - upsample
 *
 * @example
 * ```typescript
 * const df = WrappedDataFrame(data);
 * await df.writeExcel('output.xlsx');
 * ```
 */
const WrappedDataFrame = function <
  T extends Record<string, originalPl.Series<any, string>>,
>(
  ...args: Parameters<typeof originalPl.DataFrame>
): ExtendedDataFrame<T> {
  const instance = originalPl.DataFrame(
    ...args,
  ) as unknown as ExtendedDataFrame<T>;

  // Add the `writeExcel` method if it doesn't exist
  if (!instance.writeExcel) {
    instance.writeExcel = async function (
      filePath: string,
      options?: WriteExcelOptions,
    ): Promise<void> {
      await writeExcel(this, filePath, options);
    };
  }

  // Extend the methods that return a new DataFrame
  ([
    "clone",
    "describe",
    "drop",
    "dropNulls",
    "explode",
    "extend",
    "fillNull",
    "filter",
    "frameEqual",
    "head",
    "hstack",
    "interpolate",
    "join",
    "joinAsof",
    // "lazy",
    "limit",
    "max",
    "mean",
    "median",
    "unpivot",
    "min",
    "nullCount",
    "partitionBy",
    "pivot",
    "quantile",
    "rechunk",
    "rename",
    "select",
    "shift",
    "shiftAndFill",
    "shrinkToFit",
    "slice",
    "sort",
    "std",
    "sum",
    "tail",
    "transpose",
    "unique",
    "unnest",
    "var",
    "vstack",
    "withColumn",
    "withColumns",
    "withColumnRenamed",
    "withRowCount",
    "where",
    "upsample",
  ] as Array<keyof ExtendedDataFrame<any>>).forEach((method) => {
    const originalMethod = instance[method].bind(instance);
    Object.defineProperty(instance, method, {
      value: function (
        ...args: Parameters<typeof originalMethod>
      ): ExtendedDataFrame<any> {
        const newDf = originalMethod(...args);
        return WrappedDataFrame(
          newDf as unknown as Parameters<typeof originalPl.DataFrame>[0],
        );
      },
      writable: true,
      configurable: true,
    });
  });

  return instance;
};

// Replace the DataFrame factory with the wrapped one
const extendedPl = {
  ...originalPl,
  DataFrame: WrappedDataFrame as <
    T extends Record<string, originalPl.Series<any, string>>,
  >(
    ...args: Parameters<typeof originalPl.DataFrame>
  ) => ExtendedDataFrame<T>,
  readExcel: async function (
    filePath: string,
    options?: ReadExcelOptions,
  ): Promise<ExtendedDataFrame<any>> {
    const df = await readExcel(filePath, options);

    // Wrap the returned DataFrame to add the writeExcel method
    return WrappedDataFrame(
      df as unknown as Parameters<typeof originalPl.DataFrame>[0],
    );
  },
  writeExcel,
  ExtendedDataFrame: null as unknown as ExtendedDataFrame<any>,
} as ExtendedPolars;

// Override the top-level `concat` function to return an ExtendedDataFrame
extendedPl.concat = function (
  items: any[],
  options?: { rechunk?: boolean; how?: "vertical" | "diagonal" | "horizontal" },
): any {
  const result = originalPl.concat(items, options);

  // Check if result is a DataFrame and wrap it, else return raw
  if (
    originalPl.DataFrame.prototype.constructor &&
    result instanceof originalPl.DataFrame
  ) {
    return WrappedDataFrame(result);
  }

  // Series or LazyDataFrame — return as-is
  return result;
};

/**
 * Extended version of the `polars` library that provides:
 * - the original `polars` object with all its features,
 * - a wrapped `DataFrame` class (called `WrappedDataFrame`) with an added `writeExcel` method,
 * - `readExcel` and `writeExcel` functions for reading and writing Excel files,
 * - and the `ExtendedDataFrame` type for the extended DataFrame.
 *
 * Main enhancements compared to the original library:
 * - the `writeExcel` method is available directly on DataFrame instances,
 * - all DataFrame methods returning new DataFrames return the wrapped ExtendedDataFrame with `writeExcel`,
 * - the `concat` function returns an ExtendedDataFrame,
 * - the `readExcel` and `writeExcel` functions are exposed directly on the export.
 *
 * @remarks
 * This export allows seamless use of Excel-related functionality without manually wrapping DataFrame objects.
 *
 * @example
 * ```ts
 * import pl from "@jackfiszr/pl2xl";
 *
 * // Create a DataFrame
 * const df = pl.DataFrame({ a: [1, 2, 3] });
 *
 * // Write to Excel
 * await df.writeExcel("file.xlsx");
 *
 * // Read from Excel
 * const df2 = await pl.readExcel("file.xlsx");
 * ```
 */
export default extendedPl as ExtendedPolars;

export { type ExtendedDataFrame, readExcel, writeExcel };
