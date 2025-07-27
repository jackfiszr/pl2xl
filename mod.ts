import originalPl from "polars";
import { readExcel } from "./read_excel.ts";
import { writeExcel } from "./write_excel.ts";
import type {
  ExtendedDataFrame,
  ReadExcelOptions,
  WriteExcelOptions,
} from "./types.d.ts";

export type * from "polars";
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
const WrappedDataFrame = function (
  ...args: Parameters<typeof originalPl.DataFrame>
): ExtendedDataFrame<any> {
  const instance = originalPl.DataFrame(...args) as ExtendedDataFrame<any>;

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
  ] as Array<
    keyof ExtendedDataFrame<any>
  >)
    .forEach(
      (method) => {
        const originalMethod = instance[method].bind(instance);

        Object.defineProperty(instance, method, {
          value: function (
            ...args: Parameters<typeof originalMethod>
          ): ExtendedDataFrame<any> {
            // Call the original method
            const newDf = originalMethod(...args);

            // Wrap the returned DataFrame to add the writeExcel method
            return WrappedDataFrame(newDf);
          },
          writable: true,
          configurable: true,
        });
      },
    );

  return instance;
};

// Replace the DataFrame factory with the wrapped one
const extendedPl = {
  ...originalPl,
  DataFrame: WrappedDataFrame,
  readExcel: async function (
    filePath: string,
    options?: ReadExcelOptions,
  ): Promise<ExtendedDataFrame<any>> {
    const df = await readExcel(filePath, options);

    // Wrap the returned DataFrame to add the writeExcel method
    return WrappedDataFrame(df);
  },
  writeExcel,
  ExtendedDataFrame: null as unknown as ExtendedDataFrame<any>,
} as unknown as typeof originalPl & { DataFrame: typeof WrappedDataFrame };

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

export { type ExtendedDataFrame, readExcel, writeExcel };

// Export the extended Polars object
export default extendedPl;
