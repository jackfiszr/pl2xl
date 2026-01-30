import pl from "../mod.ts";

// Create a sample DataFrame
const inputDf = pl.DataFrame({
  Name: ["Alice", "Bob", "Charlie"],
  Age: [25, 30, 35],
  Empty: [null, null, null],
  City: ["New York", "Los Angeles", "Chicago"],
});

// Write the DataFrame to an Excel file
await inputDf.writeExcel("input.xlsx");

// Read the DataFrame back from the Excel file
const df = await pl.readExcel("input.xlsx");
console.log("Read DataFrame:", df.toString());

// Modify the DataFrame by increasing the "Age" column by 1
const modifiedDf = df.withColumn(pl.col("Age").add(1).alias("Age"));

console.log("Modified DataFrame:", modifiedDf.toString());

// Write the modified DataFrame to a new Excel file
await modifiedDf.writeExcel("output.xlsx");
console.log("Modified DataFrame written to output.xlsx");

// Create multiple DataFrames, one of which is empty
const emptyDf = pl.DataFrame([]);
await pl.writeExcel([inputDf, modifiedDf, emptyDf], "multiple_sheets.xlsx", {
  sheetName: ["Input", "Modified", "Empty"],
});
console.log("Multiple DataFrames written to multiple_sheets.xlsx");

// Clean up the Excel files
["input.xlsx", "output.xlsx", "multiple_sheets.xlsx"].forEach(async (file) => {
  await Deno.remove(file);
});
