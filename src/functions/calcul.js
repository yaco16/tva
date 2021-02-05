/* eslint-disable no-undef */
// export const copyColumns = () => {
//   Excel.run(
//     (context = () => {
//       return context.sync().then(() => {});
//     })
//   );
// };

import { getLastCell } from "./copyData";

const copyColumns = () => {
  //recopier les colonnes qui m'intéressent
  Excel.run(function(context) {
    //sélectionner toutes les cellules used du journal 2019-2020
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    var calcul = context.workbook.worksheets.getItem("Calcul");
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

    return context.sync();
  });
};

const invertDebitSign = () => {
  console.log('je suis dans invertDebitSign')
  Excel.run(async context => {
    var calcul = context.workbook.worksheets.getItem("Calcul");
    //récupération de la plage de toutes les cellules utilisées
    const usedRange = calcul.getUsedRange().getLastCell();
    usedRange.load("address");

    await context.sync();
    //création de lastcell de la colonne E
    const lastCell = getLastCell(usedRange.address, "E");
    //création de la formule et copie jusqu'à la dernière cellule utilisée
    var cell = calcul.getRange("E3");
    cell.values = [["=C3*-1"]];
    cell.autoFill(`E3:${lastCell}`, Excel.AutoFillType.fillCopy);
  });
};

// const changeDebitContent = () => {
//   Excel.run((context) = () => {
//     var calcul = context.workbook.worksheets.getItem("Calcul");
//     var range = calcul.getRange("E8");
//     range.load("values");
//     return context.sync().then(() => {
//       var range1 = JSON.stringify(range.values, null, 4);
//       console.log('range1:', range1);
//       // calcul.getRange("H1").copyFrom(range.values);

//     })
//   });
// };

async function moveInvertDebitSign() {
  await Excel.run(async (context) => {
    console.log('ici')
      const sheet = context.workbook.worksheets.getItem("calcul");

      // Copy the resulting value of a formula.
      sheet.getRange("c:c").copyFrom("e:e", Excel.RangeCopyType.values);
      await context.sync();
  });
}

export const process = () => {
  copyColumns();
  invertDebitSign();
};

export const calcul = () => {
  moveInvertDebitSign();
};