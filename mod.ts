import pl from "polars";
import { readExcel } from "./read_excel.ts";
import { writeExcel } from "./write_excel.ts";
import type { ReadExcelOptions } from "./types.d.ts";

type ExtendedPolars = typeof pl & {
  readExcel: (filePath: string, options?: ReadExcelOptions) => Promise<pl.DataFrame>;
};

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

export default extendedPl;
