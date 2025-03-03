import { assertEquals, assertRejects } from "@std/assert";
import { readExcel } from "./read_excel.ts";
import { createTestExcelFile, removeTestFile } from "./test_utils.ts";
import ExcelJS from "@tinkie101/exceljs-wrapper";

Deno.test({
  name: "readExcel - Reads data from a valid Excel file",
  async fn() {
    const filePath = "./test-read-valid.xlsx";
    const testData = {
      headers: ["Name", "Age", "Country"],
      rows: [
        ["Alice", 30, "USA"],
        ["Bob", 25, "Canada"],
      ],
    };
    await createTestExcelFile(filePath, testData);

    const df = await readExcel(filePath);

    const expected = [
      { Name: "Alice", Age: 30, Country: "USA" },
      { Name: "Bob", Age: 25, Country: "Canada" },
    ];
    assertEquals(df.toRecords(), expected);

    await removeTestFile(filePath);
  },
});

Deno.test({
  name: "readExcel - Throws error for missing sheet",
  async fn() {
    const filePath = "./test-read-missing-sheet.xlsx";
    const testData = {
      headers: ["Name", "Age"],
      rows: [["Alice", 30]],
    };
    await createTestExcelFile(filePath, testData);

    await assertRejects(
      async () => {
        await readExcel(filePath, { sheetName: ["NonExistentSheet"] });
      },
      Error,
      "Worksheet NonExistentSheet not found in the Excel file.",
    );
    await removeTestFile(filePath);
  },
});

Deno.test({
  name: "readExcel - Handles empty Excel file",
  async fn() {
    const filePath = "./test-empty.xlsx";
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("Sheet1");
    await workbook.xlsx.writeFile(filePath);

    const df = await readExcel(filePath);

    assertEquals(df.shape, { height: 0, width: 0 });

    await removeTestFile(filePath);
  },
});
