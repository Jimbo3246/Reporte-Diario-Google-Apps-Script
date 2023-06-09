/*Usar este codigo para pegar y copiar mas de una hoja*/
//El problema es que copia pero no elimina la data
function macro_1() 
{
   const libro = SpreadsheetApp.getActiveSpreadsheet();
//Copiarypegar PTP a PTP2
  const hojaOrigenPtp = libro.getSheetByName("PTP");
  const hojaDestinoPtp = libro.getSheetByName("PTP 2");
  const hojaBase= libro.getSheetByName("Base");
  const rangoOrigenPtp = hojaOrigenPtp.getRange(1,1,hojaOrigenPtp.getLastRow(),hojaOrigenPtp.getLastColumn()).getValues()
  hojaDestinoPtp.clearContents()
  const rangoDestinoPtp = hojaDestinoPtp.getRange(1,1,hojaOrigenPtp.getLastRow(),hojaOrigenPtp.getLastColumn()).setValues(rangoOrigenPtp)
//Copiarypegar Soporte Hogar a Soporte Hogar 2
  const hojaOrigenSh = libro.getSheetByName("Soporte Hogar");
  const hojaDestinoSh = libro.getSheetByName("Soporte Hogar 2");
  const rangoOrigenSh = hojaOrigenSh.getRange(1,1,hojaOrigenSh.getLastRow(),hojaOrigenSh.getLastColumn()).getValues()
  hojaDestinoSh.clearContents()
  const rangoDestinoSh = hojaDestinoSh.getRange(1,1,hojaOrigenSh.getLastRow(),hojaOrigenSh.getLastColumn()).setValues(rangoOrigenSh)
//Eliminando la data de la  hoja base
hojaBase.getRange(2,1,hojaBase.getLastRow(),hojaBase.getLastColumn()).clearContent();
}

function macro_2() {
//Sirve para crear y colocar el mes corto
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigenBase = libro.getSheetByName("Base");
//Elimina una fila en especifico, en este caso la fila 2
  hojaOrigenBase.deleteRow(2);
  var arregloDatos = hojaOrigenBase.getDataRange().getValues();
  var arregloMesCorto=[]
  for(var fila=1;fila <= arregloDatos.length-1;fila++)
  {
    var mes= new Date(arregloDatos[fila][9]).getMonth(); 
    //getMonth() te vota un valor numerico que va desde 0 a 11(Enero a Diciembre)
    switch(mes)
    {
      //Meses se cuenta desde 0 = Enero
      case 0:
       var month="ENE";
       arregloMesCorto.push([month])
       //se coloca el [] porque se va agregar a una columna de sheet es decir un arreglo de arreglos
      break;
      case 1:
       month="FEB";
       arregloMesCorto.push([month])  
      break;
      case 2:
       month="MAR";
       arregloMesCorto.push([month]) 
      break;
      case 3:
       month="ABR";
       arregloMesCorto.push([month])
      break;
      case 4:
       month="MAY";
       arregloMesCorto.push([month])
      break;
      case 5:
       month="JUN";
       arregloMesCorto.push([month])
      break;
      case 6:
       month="JUL";
       arregloMesCorto.push([month])
      break;      
      case 7:
       month="AGO";
       arregloMesCorto.push([month])
      break;
      case 8:
       month="SET";
       arregloMesCorto.push([month])
      break;
      case 9:
       month="OCT";
       arregloMesCorto.push([month])
      break;
      case 10:
       month="NOV";
       arregloMesCorto.push([month])
      break;
      case 11:
       month="DIC";
       arregloMesCorto.push([month])
      break;
    }
    //Logger.log(mes)   
  }

hojaOrigenBase.getRange(2,40,arregloDatos.length-1).setValues(arregloMesCorto)

//Copia y pega datos filtrados en cada pestaña
  var originalData = hojaOrigenBase.getRange(2,1,hojaOrigenBase.getLastRow()-1,40).getValues();
//Copiando de BD a Roaming 
  var hojaDestinoRoaming = libro.getSheetByName("Roaming");
  var dataR =originalData.filter(function(item){return item[22]==="BACK ROAMING"});// con una sola linea
  //Esta secuencia de seleccion se realiza para validar si el arreglo dataR[0] es 0 o nulo
 if(dataR[0]==null)
 {

 }
 else
 {
  hojaDestinoRoaming.getRange(2,1,hojaDestinoRoaming.getLastRow(),hojaDestinoRoaming.getLastColumn()).clearContent();
  hojaDestinoRoaming.getRange(2,1,dataR.length,dataR[0].length).setValues(dataR);
 }
//Copiando de BD a BackInper
  var hojaDestinoBackInper = libro.getSheetByName("Back in Per");
  var dataBackInper =originalData.filter(function(item){return item[22]==="Back Office Inc Per"});// con una sola linea

  hojaDestinoBackInper.getRange(2,1,hojaDestinoBackInper.getLastRow(),hojaDestinoBackInper.getLastColumn()).clearContent();
  hojaDestinoBackInper.getRange(2,1,dataBackInper.length,dataBackInper[0].length).setValues(dataBackInper);
//Copiando de BD a Niccs
  var hojaDestinoNiccs = libro.getSheetByName("Niccs");
  var dataNiccs =originalData.filter(function(item){return item[22]==="NICCs Externo"});// con una sola linea

  hojaDestinoNiccs.getRange(2,1,hojaDestinoNiccs.getLastRow(),hojaDestinoNiccs.getLastColumn()).clearContent();
  hojaDestinoNiccs.getRange(2,1,dataNiccs.length,dataNiccs[0].length).setValues(dataNiccs);
//Copiando de BD a PTP
  var hojaDestinoPTP_1 = libro.getSheetByName("PTP");
  var dataPTP_1 =originalData.filter(function(item){return item[22]==="Plataforma Técnica"});// con una sola linea

  hojaDestinoPTP_1.getRange(2,1,hojaDestinoPTP_1.getLastRow(),hojaDestinoPTP_1.getLastColumn()).clearContent();
  hojaDestinoPTP_1.getRange(2,1,dataPTP_1.length,dataPTP_1[0].length).setValues(dataPTP_1);
//Copiando de BD a Soporte Hogar
  var hojaDestinoSH_1 = libro.getSheetByName("Soporte Hogar");
  var dataSH_1 =originalData.filter(function(item){return item[22]==="SOPORTE HOGAR"});// con una sola linea

  hojaDestinoSH_1.getRange(2,1,hojaDestinoSH_1.getLastRow(),hojaDestinoSH_1.getLastColumn()).clearContent();
  hojaDestinoSH_1.getRange(2,1,dataSH_1.length,dataSH_1[0].length).setValues(dataSH_1);

//Creando el Backup
 var archivo=DriveApp.getFileById("1LXb5qGLhU6d37Req2xIxqJrl8-6X25kvOOc4GqJSG3Y")
 var destino=DriveApp.getFolderById("1RcSg8jfZmBik2kHU8kn3Fh1Rkbri7qa0")

//Restando 5 horas a la hora amaricana por default
 var five_Hours = 1000 * 60 * 60 * 5;
 var now = new Date();
 var day_Before_Five_Hours = new Date(now.getTime() - five_Hours);
 var formattedDate=Utilities.formatDate(day_Before_Five_Hours,'ETC/GMT',"yyyy-MM-dd' 'HH:mm:ss")
 var name=SpreadsheetApp.getActiveSpreadsheet().getName()+" Copy "+formattedDate;
 archivo.makeCopy(name,destino)
}
function creaBackUp()
{
//Creando el Backup
 var archivo=DriveApp.getFileById("1LXb5qGLhU6d37Req2xIxqJrl8-6X25kvOOc4GqJSG3Y")
 var destino=DriveApp.getFolderById("1RcSg8jfZmBik2kHU8kn3Fh1Rkbri7qa0")

//Restando 5 horas a la hora amaricana por default
 var five_Hours = 1000 * 60 * 60 * 5;
 var now = new Date();
 var day_Before_Five_Hours = new Date(now.getTime() - five_Hours);
 var formattedDate=Utilities.formatDate(day_Before_Five_Hours,'ETC/GMT',"yyyy-MM-dd' 'HH:mm:ss")
 var name=SpreadsheetApp.getActiveSpreadsheet().getName()+" Copy "+formattedDate;
 archivo.makeCopy(name,destino)
}

function fallas()
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
//SEPARANDO DATA PTP
  var hojaOrigenBasePtp = libro.getSheetByName("PTP");
  var arregloDatosPtp = hojaOrigenBasePtp.getRange(2,2,hojaOrigenBasePtp.getLastRow()-1,1).getValues();
  //Separando data de SOPORTE HOGAR
  var hojaOrigenBase = libro.getSheetByName("Soporte Hogar");
  var arregloNSSPtp=[]
  var hojaPruebaPtp ="PTP_Y_SH_FALLAS"
  nordenPTP=0
//Se utiliza la funcion map para crear un arreglo que contenga solo valores[1, Nº de SS]
//Con la condicion de q cada q el numero de orden es multiplo de  465  y  este sea el ultimo valor no añada el OR
  var arregloNSSPtp=arregloDatosPtp.map(x=>
  {
    nordenPTP+=1
    var nSSPtp=x[0]
   if(nordenPTP % 465 !=0 && nordenPTP<arregloDatosPtp.length)
    {
      nSSPtp = nSSPtp+" OR" 
    }
   return [[nordenPTP],[nSSPtp]]
  })
//Se almacena la data desde fila 2 columna 2 hasta la ultima fila, solo la columna "Nº de SS"
  var arregloDatos = hojaOrigenBase.getRange(2,2,hojaOrigenBase.getLastRow()-1,1).getValues();
  var arregloNSS=[]
  nordenSH=0
//Se utiliza la funcion map para crear un arreglo que contenga solo valores[1, Nº de SS]
//Con la condicion de q cada q el numero de orden es multiplo de  465  y  este sea el ultimo valor no añada el OR
  var arregloNSS=arregloDatos.map(x=>
  {
    nordenSH+=1
    var nSS=x[0]
   if(nordenSH % 465 !=0 && nordenSH<arregloDatos.length)
    {
      nSS = nSS+" OR"
    }
   return [[nordenSH],[nSS]]
  })
  var targetSheetPtp = libro.insertSheet(hojaPruebaPtp);
  targetSheetPtp.getRange(1,1,arregloNSSPtp.length,arregloNSSPtp[0].length).setValues(arregloNSSPtp);
  targetSheetPtp.getRange(1,4,arregloNSS.length,arregloNSS[0].length).setValues(arregloNSS);
}
