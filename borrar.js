function myFunction() {
     var spreadsheet = SpreadsheetApp.getActive();
     const sheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Roaming'), true);
     const rangodata = sheet.getRange("C:C").getValues();
  for(let i = rangodata.length-1;i>= 0;i--)
  {
      sheet.deleteRow[i+1]
  }  
}
function indexOf()
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = libro.getSheetByName("Base");
  var data = activeSheet.getRange(1,1,activeSheet.getLastRow(),activeSheet.getLastColumn()).getValues();
  var newdata = data.map(function(r){ 
    if(typeof r[0]==="string")
    {
      return r[0].toLowerCase();
    }
    });
var searchText ="expired"
  Logger.log(newdata);//En arreglos la el primer elemento comienzo en la coordenada(fila=0,columna=0)
  Logger.log(newdata.indexOf(searchText,5));
}
function cadena()
{
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = libro.getSheetByName("Base");
  var data = activeSheet.getRange(2,1,activeSheet.getLastRow(),activeSheet.getLastColumn()).getValues();
  var arreglito =[]
  var newdata = data.map(function(searchRow){if(searchRow!=''){return [searchRow[10]]}})
  var datacelda =newdata.map(function(buscafila)
  {
    if(buscafila != '')
    {
      for(i=0;i<buscafila[0].length;i++)
      {
       return [buscafila[i]]    
      } 
    }
  }
    )

Logger.log(datacelda)
/*Sub Macro_Ordenes4()
ufila = Range("A" & Rows.Count).End(xlUp).Row
cuentacel = 0
Range("B2:B" & ufila).ClearContents
For i = 2 To ufila
largo = Len(Cells(i, 1))
For j = 1 To largo
tres = Mid(Cells(i, 1), j, 1)
If tres = 5 Then
inicio = j
num = tres
contador = 1
For k = j + 1 To largo
signum = Mid(Cells(i, 1), k, 1)
If Not IsNumeric(signum) Then
j = k
Exit For
Else
contador = contador + 1
num = num & signum
If contador = 9 Then
signum = Mid(Cells(i, 1), k + 1, 1)
If Not IsNumeric(signum) Then
If cuentacel = 0 Then
Cells(i, 2) = num
Else
Cells(i, 2) = Cells(i, 2) & " y " & num
End If
cuentacel = cuentacel + 1
j = k
Exit For
End If
End If
End If
Next k
End If
Next j
If cuentacel = 0 Then
Cells(i, 2) = "No hay orden"
End If
cuentacel = 0
Next i
End Sub*/
}
