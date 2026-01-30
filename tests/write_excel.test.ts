import { assertEquals, assertRejects } from "@std/assert";
import { writeExcel } from "../src/write_excel.ts";
import { getRows, removeTestFile } from "./test_utils.ts";
import ExcelJS from "@tinkie101/exceljs-wrapper";
import pl from "../mod.ts";

Deno.test({
  name: "writeExcel - Writes a DataFrame to a valid Excel file",
  async fn() {
    const filePath = "./test-write-valid.xlsx";
    const df = pl.DataFrame({
      Name: ["Alice", "Bob"],
      Age: [30, 25],
      Country: ["USA", "Canada"],
    });

    // Write the DataFrame to Excel
    await writeExcel(df, filePath, { sheetName: "Sheet1" });

    // Read back and validate
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet("Sheet1");
    if (!worksheet) {
      throw new Error("Worksheet 'Sheet1' does not exist.");
    }
    const rows = getRows(worksheet);
    const expected = [
      ["Name", "Age", "Country"],
      ["Alice", 30, "USA"],
      ["Bob", 25, "Canada"],
    ];

    assertEquals(rows, expected);

    await removeTestFile(filePath);
  },
});

Deno.test({
  name: "writeExcel - Handles empty DataFrame",
  async fn() {
    const filePath = "./test-write-empty.xlsx";
    const df = pl.DataFrame({});

    await assertRejects(
      async () => {
        await writeExcel(df, filePath);
      },
      Error,
      "The DataFrame is empty. Nothing to write.",
    );
  },
});

Deno.test({
  name: "writeExcel - Writes to a specific sheet",
  async fn() {
    const filePath = "./test-write-specific-sheet.xlsx";
    const df = pl.DataFrame({
      Product: ["Widget", "Gadget"],
      Price: [19.99, 29.99],
    });

    // Write the DataFrame to a specific sheet
    await writeExcel(df, filePath, { sheetName: "Products" });

    // Read back and validate the correct sheet
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet("Products");
    if (!worksheet) {
      throw new Error("Worksheet 'Products' does not exist.");
    }
    const rows = getRows(worksheet);
    const expected = [
      ["Product", "Price"],
      ["Widget", 19.99],
      ["Gadget", 29.99],
    ];

    assertEquals(rows, expected);

    await removeTestFile(filePath);
  },
});

Deno.test({
  name: "writeExcel - Applies table styling",
  async fn() {
    const filePath = "./test-write-styled.xlsx";
    const df = pl.DataFrame({
      Month: ["January", "February"],
      Sales: [15000, 20000],
    });

    // Write the DataFrame with styling
    await writeExcel(df, filePath, {
      sheetName: "Report",
      tableStyle: "TableStyleMedium4",
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet("Report");
    if (!worksheet) {
      throw new Error("Worksheet 'Report' does not exist.");
    }
    const tables = worksheet.getTables();
    if (!tables.length) {
      throw new Error("Table not found.");
    }
    const { table } = JSON.parse(JSON.stringify(tables[0]));

    assertEquals(table.name, "Table_Report");
    assertEquals(table.style.theme, "TableStyleMedium4");

    await removeTestFile(filePath);
  },
});
