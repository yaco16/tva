/* global console, document, Excel, Office */
import { addSheets, copyN } from '../functions/copyData';
import { process, calcul } from '../functions/calcul';

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("add-sheet").onclick = addSheets;
    document.getElementById("copyN").onclick = copyN;
    document.getElementById("copyColumns").onclick = process;
    document.getElementById("calcul").onclick = calcul;
  }
});
