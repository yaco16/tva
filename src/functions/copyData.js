/* eslint-disable no-undef */

export const addSheets = () => {//ajouter les feuilles de départ
  Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.add("Quadra_GL2021");
    sheets.add("Quadra_GL2020");
    sheets.add("Quadra_GL_groupe");
    sheets.add("Calcul");

    //supprimer la feuille de départ Sheet1
    context.workbook.worksheets.getItem("Sheet1").delete();

    return context.sync()

  }).catch(errorHandlerFunction);
}

export const copyN = () => {
  Excel.run(function (context) {
    //sélectionner toutes les cellules used du journal 2020-2021
    var quadraN = context.workbook.worksheets.getItem("Quadra_GL2021");
    var range = quadraN.getUsedRange();

    //copier les cellules dans le journal quadraGroupe
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    quadraGroupe.getRange("A1").copyFrom(range);

    //on veut sélectionner les données déjà copiées pour ensuite se placer sur la dernière cellule utilisée
    //sélectionner toutes les cellules used du journal quadraGroupe
    var usedRangeQuadraGroupe = quadraGroupe.getUsedRange();

    //sélectionner la dernière cellule
    var lastCell = usedRangeQuadraGroupe.getLastCell();
    lastCell.load('address');

    return context.sync()
    .then(function () {
      console.log('address', lastCell.address);
      const parse = justCell(lastCell.address);
      const parse2 = decomposeCell(parse);
      const newCell = createCell(parse2, "A", 2);
      copyN_1(newCell);
        });
}).catch(errorHandlerFunction);
}

export const copyN_1 = (newCell) => {
  Excel.run(function (context) {
    //sélectionner toutes les cellules used du journal 2019-2020
    var quadraN_1 = context.workbook.worksheets.getItem("Quadra_GL2020");
    var range = quadraN_1.getUsedRange();

    //copier les cellules dans le journal quadraGroupe
    var quadraGroupe = context.workbook.worksheets.getItem("Quadra_GL_groupe");
    quadraGroupe.getRange(`${newCell}`).copyFrom(range);
    return context.sync()
  })
}

//supprimer le nom de la feuille
export const justCell = (data) => {
  console.log('data justCell:', data);
  return data.split("!").pop();
}

//séparer le nom de la colonne et le numéro de la ligne
export const decomposeCell = (data) => {
  const cellrowIsString = data.substring(1);
  const cellrowIsParsed = parseInt(cellrowIsString, 10);
  const cellColumn = data.substring(0, 1);
  return { column: cellColumn,
    row: cellrowIsParsed
   }
}

//créer la nouvelle cellule où on se positionnera
export const createCell = (data, newColumn, rowShift) => {
  rowShift ? rowShift : rowShift = 0;
  // console.log('data:', data);
  const newRow = data.row + rowShift;
  // console.log('newRow:', newRow);
  const newRowToString = newRow.toString();
  return (newColumn+newRowToString);
}

export const getLastCell = (data, letter) => {
  const justCell = data.split("!").pop();
  const cellrowIsString = justCell.substring(1);
  const cellrowIsParsed = parseInt(cellrowIsString, 10);
  const cellColumn = justCell.substring(0, 1);
  const newCell = { column: cellColumn,
    row: cellrowIsParsed
   };
  const newRow = newCell.row;
  const newRowToString = newRow.toString();
  return (letter+newRowToString);


  // const decomposedCell = decomposeCell(justCell);
  // console.log('decomposedCell:', decomposedCell);
  // const createdCell = createCell(decomposedCell, letter);
  // console.log('createdCell:', createdCell);
  // return createdCell;
}