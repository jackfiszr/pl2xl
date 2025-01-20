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
  const instance = originalPl.DataFrame(...args) as ExtendedDataFrame;

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
  (["withColumn", "withColumns"] as Array<keyof ExtendedDataFrame>).forEach(
    (method) => {
      const originalMethod = instance[method].bind(instance);

      Object.defineProperty(instance, method, {
        value: function (
          ...args: Parameters<typeof originalMethod>
        ): ExtendedDataFrame {
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
