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

  // Extend the `withColumns` method
  const originalWithColumns = instance.withColumns.bind(instance);

  instance.withColumns = function (
    columns: originalPl.Series | originalPl.Expr,
  ): ExtendedDataFrame {
    // Call the original withColumns method
    const newDf = originalWithColumns(columns);

    // Wrap the returned DataFrame to add the writeExcel method
    return WrappedDataFrame(newDf);
  };

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
};

export { readExcel, writeExcel };

// Export the extended Polars object
export default extendedPl;
