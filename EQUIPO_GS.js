//CONTSNATES GLOBALES
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LISTAS DINAMICAS');
  var hojaGastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
  var hojaColab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRESTAMOS COLABORADORES');




function doGet(){

   const libro = SpreadsheetApp.getActiveSpreadsheet();
   const pythonDta = libro.getSheetByName('PYTHON');

   const totalExigibles = pythonDta.getRange('b9').getDisplayValue();
   const totalVencidos = pythonDta.getRange('b5').getDisplayValue();
   const dif = pythonDta.getRange('b12').getDisplayValue();

  let plantilla = HtmlService.createTemplateFromFile('EQUIPO');

  plantilla.getData = getData();
  plantilla.indicadores = indicadores();
  plantilla.deHoy = deHoy();
  plantilla.totalExigibles = totalExigibles;
  plantilla.getColocacion = getColocacion()
  plantilla.getPagos = getPagos()
  plantilla. getClickeos = getClickeos()
  plantilla.getColab = getColab()
  plantilla.getGastos = getGastos()
  plantilla.dif = dif
  plantilla.getVencidos = getVencidos()
  plantilla.totalVencidos = totalVencidos;

  let web = plantilla.evaluate();
  web.addMetaTag('viewport','width=device-width, initial-scale=1')
  

  return web;
}



function include(filename){

  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



function team(){
    let plantilla = HtmlService.createTemplateFromFile('EQUIPO');
    //var  url = crearPdf();
    plantilla.getData = getData();
    let web = plantilla.evaluate().setWidth(900).setHeight(400);
    let window = SpreadsheetApp.getUi();

    window.showModalDialog(web, "A T T I");


}

function getData(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName('LISTAS DINAMICAS');
 

  var data = hoja.getRange('ar2:ax4').getValues();
  //data.shift();
  //console.log(data);

  return data;
};

function indicadores(){
   const libro = SpreadsheetApp.getActiveSpreadsheet();
   const pythonDta = libro.getSheetByName('PYTHON');

   let keys = pythonDta.getRange('A2:B17').getDisplayValues();

   return keys;
};



function deHoy(){

  var hojaData = sica.getSheetByName('LISTAS DINAMICAS');

   
  var filaInicio = hojaData.getRange('T1').getDisplayValue();
  var nFilas= hojaData.getRange('X1').getDisplayValue();
  var filas = hojaData.getRange('af1').getValue();
  var clientes =hojaData.getRange(1, 20 ,nFilas, 4).getDisplayValues();
  let suma = clientes.reduce((suma,monto)=> suma + monto[1],0);  
  //console.log(clientes);

  return clientes

};



//obtener datos de la cocloccion
function getColocacion(){
   const libro = SpreadsheetApp.getActiveSpreadsheet();
   const colocacion = libro.getSheetByName('LISTAS DINAMICAS');

   const colocaciones =  colocacion.getRange('a35:d38').getValues();

  return colocaciones

};


//FUNCION OARA OBTENER LAS OPERACIONES DEL DIA

function getPagos(){
 
   //ALGORITMO PARA PAGOS HOY
  var hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mesActual);
  var fI = hojaOrigen.getRange('c199').getDisplayValue();
  var nF = hojaOrigen.getRange('b199').getDisplayValue();
  var  pagosHoy =hojaOrigen.getRange(200,2,nF,2).getValues(); 
  

  return pagosHoy

};



function getClickeos(){

    var filas = hojaData.getRange('af1').getValue();
    //SECCION PÁRA VISUALIZAR EDICIONES DE CELDAS CLICKEO
    var clickeos = hojaData.getRange(2,31,filas,1).getValues();

    return clickeos


};

function getColab(){
   //SECCION PRA CAPTURAR OPERACION DEL DIA DE COLABORADORES
  var nRowscolab = hojaColab.getRange('aa1').getValue();
  var colabHoy = hojaColab.getRange(2,26,nRowscolab,4).getDisplayValues();

  return colabHoy
};


function getGastos(){
    //SECCION PARA CAPTURAR LOS GASTOS DE EL DIA
  var nRows = hojaGastos.getRange('ak1').getValue();
  var gastosHoy = hojaGastos.getRange(2,36,nRows,4).getDisplayValues();

    return gastosHoy
};


function getVencidos(){
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const sheetData = libro.getSheetByName('LISTAS DINAMICAS')
    const filas = sheetData.getRange('an1').getValue();

    const bigData = sheetData.getRange(2,35,filas,5).getDisplayValues();

    return bigData


};




  


