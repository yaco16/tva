/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
// import "../../assets/icon-16.png";
// import "../../assets/icon-32.png";
// import "../../assets/icon-80.png";

import { createTable, filterTable, sortTable, addSheet, addData, selectAll, selectOne } from '../functions';

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("create-table").onclick = createTable;
    // document.getElementById("filter-table").onclick = filterTable;
    // document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("add-sheet").onclick = addSheet;
    // document.getElementById("add-data").onclick = addData;
    document.getElementById("select-all").onclick = selectAll;
    document.getElementById("select-one").onclick = selectOne;
  }
});
