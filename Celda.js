function copiarCelda() 
{
  //No se puede pegar en celdas que no existan en la hoja
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  //const hojaCopy =libro.getSheetByName('Base')
  const rangofijo =hojaActiva.getRange('A1')
  rangofijo.copyTo(hojaActiva.getRange('AO1')) 
}

function moverCelda() 
{
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  //const hojaCopy =libro.getSheetByName('Base')
  const rangofijo =hojaActiva.getRange('A1')
  rangofijo.moveTo(hojaActiva.getRange('AO1'))
  
}
function pegarEspecial()
{
    //No se puede pegar en celdas que no existan en la hoja
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  //const hojaCopy =libro.getSheetByName('Base')
  const rangofijo =hojaActiva.getRange('A1')
  rangofijo.copyTo(hojaActiva.getRange('AO1'),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false)//Se puede pegar valores/formula/formato.etc
}
function pegarValores2()
{
    //No se puede pegar en celdas que no existan en la hoja
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  //const hojaCopy =libro.getSheetByName('Base')
  const rangofijo =hojaActiva.getRange('A1')
  rangofijo.copyTo(hojaActiva.getRange('AO1'),{contentsOnly:true})
}
function getYsetValues()
{
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  const valorOrigen = hojaActiva.getRange('A1').getValue();
  const rangoDestino = hojaActiva.getRange('AO1');
  rangoDestino.setValue(valorOrigen)

}
function getYsetFormulas()
{
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva=libro.getActiveSheet();
  const formulaOrigen = hojaActiva.getRange('A1').getFormula();
  const rangoDestino = hojaActiva.getRange('AO1');
  rangoDestino.setFormula(formulaOrigen)
}
function copiarCeldaOtraHoja() 
{
  //No se puede pegar en celdas que no existan en la hoja
  const libro=SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva= libro.getActiveSheet();
  const hojaDestino= libro.getSheetByName("PTP 2")
  //const hojaCopy =libro.getSheetByName('Base')
  const rangoOrigen =hojaActiva.getRange('A1')
  const rangoDestino =hojaDestino.getRange('AO1')
  rangoOrigen.copyTo(rangoDestino)
  
}
