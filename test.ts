import { readExcel, writeExcel } from "./mod.ts";
import pl from "polars";

const inputDf = pl.DataFrame({
  "Name": ["Alice", "Bob", "Charlie"],
  "Age": [25, 30, 35],
  "City": ["New York", "Los Angeles", "Chicago"],
});

writeExcel(inputDf, "input.xlsx");

const df = readExcel("input.xlsx");
console.log("Read DataFrame:", df);

const modifiedDf = df.withColumn(
  pl.col("Age").add(1).alias("Age"),
);

console.log("Modified DataFrame:", modifiedDf);

writeExcel(modifiedDf, "output.xlsx");
console.log("Modified DataFrame written to output.xlsx");
