/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  console.log("[ddguo][taskpane.ts]: office ready");
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("runweb").onclick = runweb;
  }
});

export async function run() {
  console.log("[ddguo][taskpane.ts]: run is executed.");

  try {
    await Excel.run({ mergeUndoGroup: true }, async (context) => {
      context.workbook.worksheets.load("items,items/name");
      await context.sync();

      const currentWorkbook = context.workbook;
      const activeCell = currentWorkbook.getActiveCell();
      const worksheetCount = currentWorkbook.worksheets.items.length;

      // Set the standardWidth property of the first worksheet to clear the Undo stack.
      currentWorkbook.worksheets.getActiveWorksheet().position = worksheetCount - 1;

      // Insert dummy data to current cell
      activeCell.values = [["TEST"]];

      await context.sync();
    });
  } catch (error) {
    console.error(error);
    console.error("Error message:", error.message);
    console.error("Error stack:", error.stack);
    console.error("Error code:", error.code);
    console.error("Error traceMessages:", error.traceMessages);
    console.error("Error debugInfo:", error.debugInfo);
  }
}

export async function runweb() {
  console.log("[ddguo][taskpane.ts]: run is executed.");

  try {
    await Excel.run({ mergeUndoGroup: true }, async (context) => {
      const currentWorkbook = context.workbook;
      const activeCell = currentWorkbook.getActiveCell();

      // Set the standardWidth property of the first worksheet to clear the Undo stack.
      currentWorkbook.worksheets.getActiveWorksheet().standardWidth = 8.43;

      // Insert dummy data to current cell
      activeCell.values = [["TEST"]];

      await context.sync();
    });
  } catch (error) {
    console.error(error);
    console.error("Error message:", error.message);
    console.error("Error stack:", error.stack);
    console.error("Error code:", error.code);
    console.error("Error traceMessages:", error.traceMessages);
    console.error("Error debugInfo:", error.debugInfo);
  }
}

function runDemo(event: Office.AddinCommands.Event | null) {
  console.log("[ddguo][taskpane.ts]: command received");
  if (event) {
    event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("runDemo", runDemo);
