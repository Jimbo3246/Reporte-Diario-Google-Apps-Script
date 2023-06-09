function BorrarRango() {
  let ssRoaming     = SpreadsheetApp.getActiveSpreadsheet();
  let sheetRoaming  = ssRoaming.getSheetByName("Roaming");
  let roamingRange  = sheetRoaming.selectDataRange();
  let roamingValues = sourceRange.getValues(); 
  let rowCount      = roamingValues.length;//La variable se lleva la cantidad de la filas
  let columCount    = roamingValues[0].length;//La variable se lleva la cantidad de la columnas
  let rangodata     = sheet.getRange("C:C").getValues();

  for(let i = rowCount-1;i>= 0;i--)
     { 
       sheet.deleteRow[i+1]  
     }

}

