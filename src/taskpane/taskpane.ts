/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var coordtransform=require('coordtransform');
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem("GPSData");
      // Load the table data
      table.load("columns/items");
      await context.sync();
      const longitudeColumn = table.columns.getItemAt(0).getDataBodyRange();
      const latitudeColumn = table.columns.getItemAt(1).getDataBodyRange();
      const newLongitudeColumn = table.columns.getItemAt(2).getDataBodyRange();
      const newLatitudeColumn = table.columns.getItemAt(3).getDataBodyRange();
      longitudeColumn.load("values");
      latitudeColumn.load("values");
      await context.sync();
      const longitudes = longitudeColumn.values;
      const latitudes = latitudeColumn.values;
      const newLongitudes = [];
      const newLatitudes = [];
      for (let i = 0; i < longitudes.length; i++) {
        const [newLongitude, newLatitude] = coordtransform.wgs84togcj02(longitudes[i][0], latitudes[i][0]);
        newLongitudes.push([newLongitude]);
        newLatitudes.push([newLatitude]);
      }
      newLongitudeColumn.values = newLongitudes;
      newLatitudeColumn.values = newLatitudes;
      await context.sync();
      console.log("Coordinates converted and updated successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}
