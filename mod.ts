import originalPl from "polars";
import { readExcel } from "./read_excel.ts";
import { writeExcel } from "./write_excel.ts";
import type {
  ExtendedDataFrame,
  ReadExcelOptions,
  WriteExcelOptions,
} from "./types.d.ts";

// Wrap the original DataFrame factory to add the `writeExcel` method
const WrappedDataFrame = function (
  ...args: Parameters<typeof originalPl.DataFrame>
): ExtendedDataFrame {
  const instance = originalPl.DataFrame(
    ...args,
  ) as unknown as ExtendedDataFrame;

  // Add the `writeExcel` method
  if (!instance.writeExcel) {
    instance.writeExcel = async function (
      filePath: string,
      options?: WriteExcelOptions,
    ): Promise<void> {
      await writeExcel(this, filePath, options);
    };
  }

  // Wrap methods that return a new DataFrame
  ([
    "clone",
    "describe",
    "drop",
    "dropNulls",
    "explode",
    "extend",
    "fillNull",
    "filter",
    "head",
    "hstack",
    "interpolate",
    "join",
    "joinAsof",
    "limit",
    "max",
    "mean",
    "median",
    "min",
    "pivot",
    "rechunk",
    "rename",
    "select",
    "shift",
    "slice",
    "sort",
    "sum",
    "tail",
    "transpose",
    "unique",
    "vstack",
    "withColumn",
    "withColumns",
    "withColumnRenamed",
    "withRowCount",
    "where",
  ] as Array<keyof Omit<ExtendedDataFrame, "writeExcel">>)
    .forEach(
      (method) => {
        const originalMethod = (instance[method] as Function).bind(instance);

        Object.defineProperty(instance, method, {
          value: function (...args: any[]): ExtendedDataFrame {
            const newDf = originalMethod(...args);
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
  ): Promise<ExtendedDataFrame> {
    const df = await readExcel(filePath, options);

    // Wrap the returned DataFrame to add the writeExcel method
    return WrappedDataFrame(df);
  },
  writeExcel,
};

export { readExcel, writeExcel };

// Export the extended Polars object
export default extendedPl;
