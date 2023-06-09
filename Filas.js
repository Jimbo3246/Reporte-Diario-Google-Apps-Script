// @ts-nocheck
function copiarFilas() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getActiveSheet()
  const rangoOrigen = hoja.getRange(1,1,1,40)//('A1:AN1')
  const rangoDestino = hoja.getRange(1421,1)//('A1421:AN1421')
  rangoOrigen.copyTo(rangoDestino)
}
function copiarFilasSetYGet() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getActiveSheet()
  const valoresOrigen = hoja.getRange(1,1,1,40).getValues();//('A1:AN1')
  Logger.log(valoresOrigen)
  const rangoDestino = hoja.getRange(1421,1,1,40)//('A1421:AN1421')
  rangoDestino.setValues(valoresOrigen)
}
function copiarFilasOtraHoja() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getActiveSheet();
  const hojaDestino = libro.getSheetByName("PTP 2");
  const rangoOrigen = hojaOrigen.getRange(1,1,1,40);//('A1:AN1')
  const rangoDestino = hojaDestino.getRange(1513,1,1,40)//('A1513:AN1513')
  rangoOrigen.copyTo(rangoDestino)
}
function copiarFilasOtraHojaSyG() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getActiveSheet();
  const hojaDestino = libro.getSheetByName("PTP 2");
  const rangoOrigen = hojaOrigen.getRange(1,1,1,40).getValues();//('A1:AN1')
  const rangoDestino = hojaDestino.getRange(1513,1,1,40).setValues(rangoOrigen)//('A1513:AN1513')  
}

function copiarPTPaPTP2() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getSheetByName("PTP");
  const hojaDestino = libro.getSheetByName("PTP 2");
  const rangoOrigen = hojaOrigen.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).getValues()
  const rangoDestino = hojaDestino.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).setValues(rangoOrigen)
}
function copiarSHaSH2() 
{ const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = libro.getActiveSheet();
  const hojaDestino = libro.getSheetByName("Soporte Hogar 2");
  const rangoOrigen = hojaOrigen.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).getValues()
  const rangoDestino = hojaDestino.getRange(1,1,hojaOrigen.getLastRow(),hojaOrigen.getLastColumn()).setValues(rangoOrigen)
}

//Funcion para realizar la copia de una hoja A hacia una hoja B con un filtro
function copy_BasetoRoaming()
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Base");
  var originalData = hoja.getRange(2,1,hoja.getLastRow()-1,39).getValues();
  var hojaDestinoRoaming = libro.getSheetByName("Roaming");
  var data =originalData.filter(function(item){return item[22]==="BACK ROAMING"});// con una sola linea
  Logger.log(data)
  hojaDestinoRoaming.getRange(2,1,hojaDestinoRoaming.getLastRow(),hojaDestinoRoaming.getLastColumn()).clearContent();
  hojaDestinoRoaming.getRange(2,1,data.length,data[0].length).setValues(data);
}

