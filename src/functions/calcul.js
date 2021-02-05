/* eslint-disable no-undef */
// export const copyColumns = () => {
  // Excel.run(function (context) {
  // return context.sync()
  // })
// }

import { getLastCell} from './copyData';

export const copyColumns = () => { //recopier les colonnes qui m'intéressent
  Excel.run(function (context) {
    //sélectionner toutes les cellules used du journal 2019-2020
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    var calcul = context.workbook.worksheets.getItem("calcul");
    calcul.activate();

    var column1 = quadraGroupe.getRange("F:F");
    var column2 = quadraGroupe.getRange("E:E");
    var column3 = quadraGroupe.getRange("H:H");
    var column4 = quadraGroupe.getRange("J:J");
    var column5 = quadraGroupe.getRange("I:I");

    //copier les colonnes dans la feuille calcul
    calcul.getRange("A1").copyFrom(column1);
    calcul.getRange("B1").copyFrom(column2);
    calcul.getRange("C1").copyFrom(column3);
    calcul.getRange("D1").copyFrom(column4);
    calcul.getRange("F1").copyFrom(column5);

    //adapter la largeur de la colonne A
    calcul.getRange("A:A").format.autofitColumns();

    return context.sync()
  })
}

export const calcul = () => {
  Excel.run( (context) => {
    var calcul = context.workbook.worksheets.getItem("calcul");
    const usedRange = calcul.getUsedRange().getLastCell();
    usedRange.load('address')

    return context.sync()
    .then(() => {
      console.log('usedRange:', usedRange.address);
      const lastCell = getLastCell(usedRange.address, "E");
      // console.log('lastCell:', lastCell);
      var cell = calcul.getRange("E3");
      cell.values = [["=C3*-1"]];
      cell.autoFill(`E3:${lastCell}`, Excel.AutoFillType.fillCopy);


    })
  })

  // Excel.run( async (context) => {
  //   var calcul = context.workbook.worksheets.getItem("calcul");
  //   var cell = calcul.getRange("E3");
  //   cell.values = [["=C3*-1"]];

  //   await context.sync()
  //   .then(() => {
  //     const lastCell = getLastCell(usedRange, "E");
  //     console.log('lastCell:', lastCell);
  //     cell.autoFill(`E3:${lastCell}`, Excel.AutoFillType.fillCopy);
  //   })
  // })
}