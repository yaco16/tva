/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
}

// Assign event handlers and other initialization logic.
document.getElementById("create-table").onclick = createTable;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("filter-table").onclick = filterTable;
  }
});

function createTable() {
  Excel.run(function (context) {

    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

expensesTable.getHeaderRowRange().values =
    [["Date", "Merchant", "Category", "Amount"]];

expensesTable.rows.add(null /*add at the end*/, [
    ["1/1/2017", "The Phone Company", "Communications", "120"],
    ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
    ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
    ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
    ["1/11/2017", "Bellows College", "Education", "350.1"],
    ["1/15/2017", "Trey Research", "Other", "135"],
    ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
]);

expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
expensesTable.getRange().format.autofitColumns();
expensesTable.getRange().format.autofitRows();

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function filterTable() {
  Excel.run(function (context) {

    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
