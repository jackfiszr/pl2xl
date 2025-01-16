import pl from "polars";
import { readExcel } from "./read_excel.ts";
import { writeExcel } from "./write_excel.ts";
import type { ReadExcelOptions, WriteExcelOptions } from "./types.d.ts";

// Extend the Polars DataFrame type to include the writeExcel method
declare module "polars" {
  interface DataFrame {
    writeExcel: (
      filePath: string,
      options?: WriteExcelOptions,
    ) => Promise<void>;
  }
}

// Save a reference to the original DataFrame factory function
const OriginalDataFrame = pl.DataFrame;

// Create a wrapper for the DataFrame function
function WrappedDataFrame(
  ...args: Parameters<typeof OriginalDataFrame>
): pl.DataFrame {
  // Create the original DataFrame instance
  const instance = OriginalDataFrame(...args);

  // Dynamically add the `writeExcel` method to the DataFrame instance
  instance.writeExcel = async function (
    filePath: string,
    options?: WriteExcelOptions,
  ): Promise<void> {
    await writeExcel(this, filePath, options);
  };

  return instance;
}

// Replace the original DataFrame function with the wrapped one
Object.defineProperty(pl, "DataFrame", {
  value: WrappedDataFrame,
  writable: false, // Ensure that the replacement is not modifiable
});

// Extend the Polars type to include the readExcel method
type ExtendedPolars = typeof pl & {
  readExcel: (
    filePath: string,
    options?: ReadExcelOptions,
  ) => Promise<pl.DataFrame>;
};

// Attach the readExcel function to the Polars module
const extendedPl: ExtendedPolars = {
  ...pl,
  readExcel: async function (
    filePath: string,
    options?: ReadExcelOptions,
  ): Promise<pl.DataFrame> {
    return await readExcel(filePath, options);
  },
};

export { readExcel, writeExcel };

// Export the extended Polars object
export default extendedPl;
