/* eslint-disable no-undef */
// import {transactions} from '../../data/transactions';
import {data} from '../../data/csvjson';

// import { startValues } from '../../data';

// export const createTable = () => {
//   Excel.run(function(context) {
//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
//     expensesTable.name = "ExpensesTable";

//     expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

//     expensesTable.rows.add(null /*add at the end*/, startValues);

//     expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
//     expensesTable.getRange().format.autofitColumns();
//     expensesTable.getRange().format.autofitRows();

//     return context.sync();
//   }).catch(function(error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// export const sortTable = () => {
//   Excel.run(function(context) {
//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
//     const sortFields = [
//       {
//         key: 1, // Merchant column
//         ascending: false
//       }
//     ];

//     expensesTable.sort.apply(sortFields);
//     return context.sync();
//   }).catch(function(error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// export const filterTable = () => {
//   Excel.run(function(context) {
//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
//     const categoryFilter = expensesTable.columns.getItem("Category").filter;
//     categoryFilter.applyValuesFilter(["Education", "Groceries"]);

//     return context.sync();
//   }).catch(function(error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

//Ajouter une feuille
export const addSheet = () => {
  Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.add("T9_bal_agee");
    sheets.add("Quadra_GL2021");
    sheets.add("Quadra_GL2020");
    sheets.add("Quadra_GL_groupe");
    sheets.add("Calcul");

    //supprimer la feuille de départ Sheet1
    context.workbook.worksheets.getItem("Sheet1").delete();

    return context.sync()
        // .then(function () {
        //     console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        // });
  }).catch(errorHandlerFunction);
}

// export const addData = () => {

//   Excel.run(function (context) {
//     var sheet = context.workbook.worksheets.getItem("Sample1");

//     var expensesTable = sheet.tables.add("A1:Q2", true /*hasHeaders*/);
//     expensesTable.name = "ExpensesTable";
//     expensesTable.getHeaderRowRange().values = [["A", "B", "C", "D", "E","F","G","H","I", "J","K","L","M","N","O","P", "Q"]];

//     var newData = data.map(item =>
//         [item.A, item.B, item.C, item.D, item.E, item.F, item.G, item.H, item.I, item.J, item.K, item.L, item.M, item.N, item.O, item.P, item.Q ]);

//     expensesTable.rows.add(null, newData);

//     if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
//         sheet.getUsedRange().format.autofitColumns();
//         sheet.getUsedRange().format.autofitRows();
//     }

//     sheet.activate();

//     return context.sync();
// }).catch(errorHandlerFunction);
// }

export const selectAll = () => {
  Excel.run(function (context) {
    console.log('je suis dans select all');
    //sélectionner toutes les cellules used du journal 2020-2021
    var quadraN = context.workbook.worksheets.getItem("Quadra_GL2021");
    var range = quadraN.getUsedRange();

    //copier les cellules dans le ourrnal quadraGroupe
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    quadraGroupe.getRange("A1").copyFrom(range);

    //sélectionner toutes les cellules used du journal quadraGroupe
    var usedRangeQuadraGroupe = quadraGroupe.getUsedRange();

    //sélectionner la dernière cellule
    var lastCell = usedRangeQuadraGroupe.getLastCell();
    lastCell.load('address');

    // console.log('lastCell:', lastCell);

    return context.sync()
    .then(function () {
      console.log('address', lastCell.address);
      const parse = justCell(lastCell.address);
        const parse2 = decomposeCell(parse);
        const newCell = createCell(parse2, "A", 2);
        activeCell(newCell);
        });
}).catch(errorHandlerFunction);
}

export const activeCell = (newCell) => {
  Excel.run(function (context) {
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    var rangeY = quadraGroupe.getRange(`${newCell}`);
    rangeY.select();
    return context.sync()
  })
}

export const justCell = (data) => {
  return data.split("!").pop();
}

export const decomposeCell = (data) => {
  const cellrowIsString = data.substring(1);
  const cellrowIsParsed = parseInt(cellrowIsString, 10);
  const cellColumn = data.substring(0, 1);
  return { column: cellColumn,
    row: cellrowIsParsed
   }
}

export const createCell = (data, newColumn, rowShift) => {
  // console.log('data:', data);
  const newRow = data.row + rowShift;
  // console.log('newRow:', newRow);
  const newRowToString = newRow.toString();
  return (newColumn+newRowToString);
}