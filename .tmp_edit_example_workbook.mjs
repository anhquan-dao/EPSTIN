import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "examples/can_signal_input_example.xlsx";
const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

if (process.argv[2] === "preview") {
  const sheetName = workbook.worksheets.getItemOrNullObject("NET_RX").isNullObject ? "Signals" : "NET_RX";
  const preview = await workbook.render({ sheetName, range: "A1:K6", scale: 1.5, format: "png" });
  await fs.writeFile(".tmp_signals_preview.png", new Uint8Array(await preview.arrayBuffer()));
  process.exit(0);
}

const variantSheet = workbook.worksheets.getItemOrNullObject("VSO_VAR");
if (variantSheet.isNullObject) workbook.worksheets.getItem("Variant").name = "VSO_VAR";
const signals = workbook.worksheets.getItemOrNullObject("NET_RX");
const activeSignals = signals.isNullObject ? workbook.worksheets.getItem("Signals") : signals;
if (signals.isNullObject) activeSignals.name = "NET_RX";
const firmwareSheet = workbook.worksheets.getItemOrNullObject("DDL");
if (firmwareSheet.isNullObject) workbook.worksheets.getItem("Firmware_DTC").name = "DDL";

activeSignals.getRange("K3:K5").values = [
  ["Test type"],
  ["Signal"],
  ["Signal"],
];
activeSignals.getRange("K3").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
activeSignals.getRange("K3").format.columnWidth = 18;
workbook.recalculate();

const check = await workbook.inspect({ kind: "table", range: "NET_RX!A3:K5", include: "values", tableMaxRows: 3, tableMaxCols: 11 });
console.log(check.ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(workbookPath);
