//VARIABLES GLOBALES

const sica = SpreadsheetApp.getActiveSpreadsheet();
var hojaActiva = sica.getActiveSheet();
var mesActual = "CONTROL PAGOS JUNIO";
const fecha = new Date();
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



function onOpen() {
 
  var Dajer = SpreadsheetApp.getActiveSpreadsheet();
  var login =Dajer.setActiveSheet(Dajer.getSheetByName('LOGIN'),true); 
  var lastRow =login.getRange('n2').getValue();
  var vencidos = login.getRange(4,13,lastRow,2).getDisplayValues();
 
  //Browser.msgBox(" Recordatorio de caso especial RODRIGO ALVAREZ NAVARRO	$2985, para que eplique el pago diferente se debe capturar hoy ");
  //Browser.msgBox("Casos vencidos recintes para seguimiento "+ " " + vencidos);
  
   
};

//FUNCION PARA EMITIR NOTIDICACIONES TOAST 
function notify(mensaje) {
   
  SpreadsheetApp.getActive ().toast(mensaje,"A T T I SYSTEM");
  
  
};



//FUNCION PARA OBTENER EL USUARIO ACTRIVO
function getUser(){
  let usuario = Session.getEffectiveUser().getEmail();
  let usuario1 = Session.getEffectiveUser().getUserLoginId();

  console.log(usuario);
   console.log(usuario1);


};

function getHora(){
  let fecha = new Date();
  let hora = fecha.getHours();
  let minutos = fecha.getMinutes();


  var momento = hora + ":" + minutos;

  return momento;



};

function onEdit(e){

  alertas();
  activeCelda();
  capturarCeldas();
  


};


//FUNCION PARA EMITIR NOTIDICACIONES TOAST 
function notify(mensaje) {
   
  SpreadsheetApp.getActive ().toast(mensaje,"A T T I SYSTEM");
  
};


function crearMenu(){
    var menu = SpreadsheetApp.getUi().createMenu('Menu Dajer');
    var menu1 = SpreadsheetApp.getUi().createMenu('Reportes Contables');
    var menu2 = SpreadsheetApp.getUi().createMenu('Funciones Admin');
    var menu3 = SpreadsheetApp.getUi().createMenu('Cartera Credito');

    menu1.addItem('Detalle_Recuperacion', 'contables')
    .addItem('Detalle_Gastos', 'reporteGastos')
    .addItem('Detalle_Ahorro', 'clientesAhorro')
    .addItem('Kpis', 'kipisModal')
    .addItem('Detalle_flujo_caja', 'cashflow');

    menu2.addItem('Crear Cedula','crearCedula')
    .addItem('Conectar Cedula','conectarCedula')
    .addItem('Abrir Cedula','abrirCedula')
    .addItem('Nuestro Equipo','team')
    .addItem('Prestamos Colaboradores','colabModal')
    .addItem('Aplicar Pagos en Cedula Cliente','buscarCliente')
    .addItem(mesActual,'pagosModal')
    .addItem("Limpiar Lista",'borrarLista')
    .addItem('Cotizador','cotizadorForm');


    menu3.addItem('Exigibles_Hoy ', 'exigibles')
    .addItem('C_Vigente ', 'vigentes')
    .addItem('C_Vencida ', 'vencidos')
    .addItem("Casos Recientes","casosRecientes")
    .addItem("Tabla Vencimientos","tablaVencimientos")
    .addItem('Rentabilidad x Cliente','rentModal');



    menu.addItem('Kips.', 'indicadores')
    menu.addSeparator();
    menu.addSubMenu(menu1);
    menu.addSeparator();
    menu.addSubMenu(menu3);
    menu.addSeparator();
    menu.addSubMenu(menu2);
    menu.addToUi();



};


function menuDario(){
    var menu = SpreadsheetApp.getUi().createMenu('Menu Dario');
    var menu2 = SpreadsheetApp.getUi().createMenu('Reportes Contables');
    var menu1 = SpreadsheetApp.getUi().createMenu('Funciones Admin');
    var menu3 = SpreadsheetApp.getUi().createMenu('Cartera Credito');

    menu2.addItem('Detalle_Recuperacion', 'contables')
    .addItem('Detalle_Gastos', 'reporteGastos')
    .addItem('Detalle_Ahorro', 'clientesAhorro')
    .addItem('Kpis', 'kipisModal')
    .addItem('Detalle_flujo_caja', 'cashflow');

    menu1.addItem('Nuestro Equipo','team')
    .addItem('Prestamos Colaboradores','colabModal')
    .addItem('Cotizador','cotizadorForm');
    /*.addItem('Pagos en Cedula Cliente','buscarCliente')
    .addItem('Registro de Pagos Control de Pagos','pagosModal');*/


    menu3.addItem('Exigibles_Hoy ', 'exigibles')
    .addItem('C_Vigente ', 'vigentes')
    .addItem('C_Vencida ', 'vencidos')
    .addItem("Casos Recientes","casosRecientes")
    .addItem("Tabla Vencimientos","tablaVencimientos")
    .addItem('Rentabilidad x Cliente','rentModal');



    menu.addItem('Kips.', 'indicadores')
    menu.addSeparator();
    menu.addSubMenu(menu3);
    menu.addSeparator();
    menu.addSubMenu(menu2);
    menu.addSeparator();
    menu.addSubMenu(menu1);
    menu.addToUi();

    

};


//FUNCION PARA ACTIVAR UNA CASILLA DE VERIICACION
function click(){
 let casilla = hojaActiva.getRange('C19');
 casilla.setValue(true);

};


//FUNCION PAR CAPTURAR CELDAS EDITADAS

function capturarCeldas(){
    const usuario = Session.getActiveUser().getEmail();
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaLista = libro.getSheetByName('LISTAS DINAMICAS');
    const fila = hojaLista.getRange('af1').getValue();


    let nombreHoja = hojaActiva.getName();
    let celdaActiva = hojaActiva.getActiveCell();
    let valor = hojaActiva.getActiveCell().getValue();
    let ubicacion =  hojaActiva.getActiveCell().getRow();
    let fecha = new Date().getDate() + " / " + (new Date().getMonth() +1);
    //var user = e.user;


    if(celdaActiva.getRow() > 1 & celdaActiva.getColumn() > 1 & valor == true & hojaActiva.getIndex() >20){

    hojaLista.getRange(fila+1,31,1,1).setValue("Se clickeo un pago  de : " + nombreHoja + " en la fila  : " +  ubicacion + " con fecha : " + fecha + " a las " + getHora()+ " " + " Usuario : " + usuario );

    //notify("Se edito la hoja con nombre  : " + nombreHoja);

    }

};






//Funcion para emitir alertas al realizar cierta funcion o cumplir una condicion.

function alertas(){
  
 
  var pagos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mesActual);
  var celdaActiva = pagos.getActiveCell(); 
  var filaActiva = celdaActiva.getRow();
  var columnActiva= celdaActiva.getColumn();
  var ultimaFila = pagos.getRange('c124').getValue()+1;
  //Logger.log(ultimaFila);
  var Monto = pagos.getRange(ultimaFila, columnActiva).getValue();
  var Cliente =  pagos.getRange(1, columnActiva).getValue();
  //var fecha = pagos.getRange(filaActiva, 2).getValue();
  var fecha = new Date();
  //Logger.log(fecha);
  
  
  if(filaActiva==ultimaFila && columnActiva >3 && Monto>0 && pagos.getName()==mesActual){
     
    var mensaje ="El pago por la cantidad de : $"+ Monto + "  " + "ha sido registrado con exito al cliente "+ " "+ Cliente ;
    notify(mensaje);
  }
  
   
};


//FUNCION PARA ACTIVAR UNA CELDA DE ACUERDO A UN VALOR

function activeCelda(){

    let book = SpreadsheetApp.getActiveSpreadsheet();
    let hojaDatos = book.getSheetByName(mesActual);
    let hojaActiva = book.getActiveSheet();
    var celdaActiva = hojaDatos.getActiveCell();
    var filaActiva = celdaActiva.getRow();
    var columActiva = celdaActiva.getColumn();
    let valorBuscaddo = hojaDatos.getRange('B1').getValue();
    let ultimaFila = hojaDatos.getRange('c122').getValue() + 1;
    let listaBusqueda = hojaDatos.getRange(1,4,1,hojaDatos.getLastColumn()).getValues();

    //console.log(listaBusqueda[0]);
    //console.log(valorBuscaddo);


    listaBusqueda.forEach(cliente => {

    if(hojaActiva.getName() == hojaDatos.getName() && filaActiva == 1 && columActiva == 2 && valorBuscaddo != ""){
    var indice = cliente.indexOf(valorBuscaddo) +4;
    //console.log(indice);
    celdaActiva = hojaDatos.getRange(ultimaFila, indice).activate();


    }


    }
    );


};







//Proteger KYC 

function protectKYC(){
  var hojaControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KYC');
  var celdaActiva = hojaControl.getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();  
  var valor = celdaActiva.getValue();
  var ultimafila = hojaControl.getRange('AL2').getValue();
  var ultimaColumna = hojaControl.getLastColumn();
  /*Logger.log(ultimaFila);*/
  
  
  var rango = hojaControl.getRange(1, 1, ultimafila, ultimaColumna);
  var proteccion = rango.protect().setDescription('no tocar').removeEditors(['contacto.prodemex@gmail.com']);
  hojaControl.hideRows(2,280);
  hojaControl.getRange('D1').activate();
};





function copyPaste(){
  
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getActiveSheet();
  var hojaEspecifica = libro.getSheetByName('LOGIN');  
  var celdaActiva = hojaEspecifica.getActiveCell();
  var valor = celdaActiva.getValue();
  var filaActiva = celdaActiva.getRow();
  var columnActiva = celdaActiva.getColumn();  
    
  if(filaActiva >99 && columnActiva == 7 && valor == "OK" && hojaEspecifica.getName()== "LOGIN"){
  var rangoOrigen =hojaEspecifica.getRange(filaActiva,2, 1, 4).getValues();
  //use en rangodestino filaactiva porqe esta en la misma hoja para prueba prodebemos byscar laultima fila + 1  
  var rangoDestino = hoja.getRange(filaActiva, 2, 1, 4);
    //Browser.msgBox(rangoOrigen);
   rangoDestino.setValues(rangoOrigen).setFontFamily('Comfortaa');
  //para mover la fila se usa moveTo
   // rangoOrigen.moveTo(rangoDestino);
  //en estos casos se debe eliminar la fila que quedaria vacia con 
  //hoja.deleteRow(filaActiva);  
  }

};


function protectRangoFecha(){
  
  var hojaControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mesActual);
  var celdaActiva = hojaControl.getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();  
  var valor = celdaActiva.getValue();
  var fecha = hojaControl.getRange('C122').getValue();
  var ultimaColumna = hojaControl.getLastColumn();
  /*Logger.log(ultimaFila);*/
  
  
  var rango = hojaControl.getRange(1, 1, fecha, ultimaColumna);
  var proteccion = rango.protect().setDescription('no tocar').removeEditors(['contacto.prodemex@gmail.com']);
};






 


function NOTCOLOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K13:L14').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('K9:L9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E9:F9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E13:F13').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('E18:F18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('H9:I9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#9900ff');
  spreadsheet.getRange('L26:N26').activate();
  spreadsheet.getActiveRangeList().setFontColor('red');
  spreadsheet.getRange('K17:L18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  


};

function WHITCOLOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K13:L14').activate();
  spreadsheet.getActiveRangeList().setFontColor('#0000ff');
  spreadsheet.getRange('K9:L9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E9:F9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E13:F13').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('E18:F18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('H9:I9').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('l26:n26').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('K17:L18').activate();
  spreadsheet.getActiveRangeList().setFontColor('#0000ff');
};





function salirDashboard(){
  libro = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = libro.getSheetByName('DASHBOARD DIGITAL');
  dashboard.getRange('K1').clearContent();
  dashboard.hideSheet();
  libro.setActiveSheet(libro.getSheetByName('LOGIN'),true);
  
};



function salir(){
 
  var gla =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TC');
  gla.getRange('M4').clear();
  gla.hideSheet();
  /*gla.setActiveSheet(gla.getSheetByName('LOGIN'),true);*/
  
};


function KEY(){  
  NOTCOLOR();
  salir();
  
};

function TDV() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:B4').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('VENCIMIENTOS'), true);
};



function BLOQUEO() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var protection = spreadsheet.getActiveSheet().protect();
  protection.removeEditors(['contacto.prodemex@gmail.com']);
  spreadsheet.getRange('A1').activate();
  
};

function DESBLOQUEO() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var protection = spreadsheet.getActiveSheet().protect();
  protection.addEditors(['contacto.prodemex@gmail.com']);
  spreadsheet.getRange('P2').activate();
};




function saliredoglobal() {
  /*Browser.msgBox("Favor de verficar la edicion de celdas,mencionar en las notas los movimientos realizados,Recuerda registrar los pagos, dando click y capturando el importe en el control de pagos del mes vigente!");*/
  var spreadsheet = SpreadsheetApp.getActive(); 
  spreadsheet.getActiveSheet().hideSheet();
  
  
  
};




function duplicar_hoja() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.duplicateActiveSheet();
};


function ELIMINAR_HOJA() {
 var spreadsheet = SpreadsheetApp.getActive();
     spreadsheet.getRange('A1').activate();
     spreadsheet.deleteactiveSheet();
  
};

function MENSAJE() {
  Browser.msgBox("DAJER sapi INFORMA " + " " + " " +"$201,570 VIGENTE" + " " + "$130,606.63 vencido  al dia de hoy"+ " " + "31 CLIENTES VENCIDOS"+""+""+"33 VIGENTES ");
  
  /*Browser.msgBox(" VIVA MEXICO!",Browser.Buttons.OK_CANCEL );*/
 
  
  
};






function KYC() {
   Browser.msgBox("Bienvenido a Base de Clientes ! te recordamos capturar correctamente cada dato , para evitar inconsistencias ");
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('KYC'), true);
    spreadsheet.getRange('C1').activate();
};



function OCULTAR_TABLA() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A43:al170').activate();
  spreadsheet.getActiveRangeList().setFontColor('#ffffff');
  spreadsheet.getRange('a43').activate();
};



function MOSTRAR_TABLA() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A43:al170').activate();
  spreadsheet.getActiveRangeList().setFontColor('#0000ff');
  spreadsheet.getRange('a43').activate();
};




 
function SIJDJR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SICDJR'), true);
  spreadsheet.getRange('A1').activate();
};

function BALANCE() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BALANCE GENERAL PRODEMEX'), true);
  spreadsheet.getRange('e21').activate();
};


function lof() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LOF'), true);
  spreadsheet.getRange('E21').activate();
};



function VENCIMIENTOS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TABLADINAMICAVENCIMIENTOS'), true);
  spreadsheet.getRange('A1').activate();
};

function LISTAS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTAS DINAMICAS'), true);
  spreadsheet.getRange('A1').activate();
};



function formato(){
  var spreadsheet =SpreadsheetApp.getActiveSheet();
  spreadsheet.getActiveRangeList().setFontFamily('Comfortaa').setFontSize(14);
    
   
}

function pesomexicano() {
  var spreadsheet =SpreadsheetApp.getActiveSheet();
  spreadsheet.getActiveRangeList().setNumberFormat('$ 0.00');
  
}








function calculadora(){
  var hojadecalculo =SpreadsheetApp.getActive();
  var hojaActiva =hojadecalculo.getSheetByName('prueba');
  var num1 = parseFloat(Browser.inputBox('captura el monto'));
  var num2 = parseFloat(Browser.inputBox('captura el monto'));
  var suma = num1 + num2 ;
  var resta =num1 - num2;
  var division = num1 / num2;
  var multiplicacion =num1 * num2;
  var porcentajes = num1 * num2  /100 ;
  hojaActiva.getRange(4, 11).setValue(suma);
  hojaActiva.getRange(5, 11).setValue(resta);
  hojaActiva.getRange(6, 11).setValue(multiplicacion);
  hojaActiva.getRange(7, 11).setValue(division);
  hojaActiva.getRange(8, 11).setValue(porcentajes);
  hojaActiva.getRange(2,11).setValue(num1);
  hojaActiva.getRange(2,12).setValue(num2);
  /*Browser.msgBox("el resultado de tu operacion es" + " " + multiplicacion);*/
  
      
  
  
}

function limpiar() {
  var spreadsheet = SpreadsheetApp.getActive();  
  spreadsheet.getRange('L12:L14').activate();  
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('L18:L20').activate();  
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B6').activate();
 
};


function calcular(){
  var sgcc = SpreadsheetApp.getActive();
  var hoja =sgcc.getSheetByName('LOGIN');
  var monto =parseFloat(Browser.inputBox('CAPTURA EL MONTO'));
  var plazo =parseFloat(Browser.inputBox('CAPTURA EL PLAZO (en meses)'));
  var tasa =parseFloat(Browser.inputBox('CAPTURA LA TASA'));
  var frecuencia =Browser.inputBox('CAPTURA LA FRECUENCIA DE PAGO'+ " " + "s=semanal,q=quincenal");
   	if (frecuencia == "s"){
	frecuencia = 4 * plazo;
	
	} else {
	
	frecuencia = 2 * plazo;
	}
	
  
  var capital = monto /  frecuencia ;
  var intereses = monto * tasa/100 * plazo;
  var interes = intereses / frecuencia;
  var iva =(interes * 16)/100; 
    
  var pago = capital + interes + iva ;
  var total =pago * frecuencia;
  
  hoja.getRange(12,12).setValue(monto);
  hoja.getRange(13,12).setValue(plazo);
  hoja.getRange(14,12).setValue(tasa/100);
  hoja.getRange(15,13).setValue(frecuencia);
  hoja.getRange(18,12).setValue(pago);
   hoja.getRange(20,12).setValue(total);
   hoja.getRange('k20').activate();
  Browser.msgBox("TU PAGO DE ACUERDO A LOS DATOS CAPTURADOS ES DE : " + " " + "$"+ pago + " " + "(tasa promedio prodemex 8%)" );
  
  
  
  
};


function indicadores(){
  var sgcc = SpreadsheetApp.getActive();
  var hoja =sgcc.getSheetByName('TABLERO CONSULTA');
 
  var BG=sgcc.getSheetByName('LOGIN');
  
  var cp= hoja.getRange(13,10).getValue()*100;
  var caja= hoja.getRange(13,11).getValue();
  var cv=hoja.getRange(9,5).getValue();
  var vhoy=hoja.getRange(13,5).getValue();
   var clvi=hoja.getRange(17, 14).getValue();
  /*var clve=RP.getRange(2, 8).getValue();*/
  var dif=BG.getRange(1,15).getValue();
  var nclientes=BG.getRange(1,13).getValue();
  
  Browser.msgBox("capital vigente=" +"$" + cv );
  Browser.msgBox("C V =" + " %" + cp);
  Browser.msgBox("exigible del dia" + " "+ "$"+ vhoy);
   Browser.msgBox("clientes d hoy ="+ nclientes);
  Browser.msgBox("creditos activos ="+ clvi);
  /*Browser.msgBox("clientes vencidos ="+ clve);*/
  Browser.msgBox("diferencia balance ="+ dif);
  Browser.msgBox("saldo caja ="+"$"+ caja + " " + "MXN");
 
  /*Browser.msgBox("PARA QUE EL SISTEMA NO PRESENTE FALLAS, ES NECESARIO CUBRIR ALA BREVEDAD, EL SALDO PENDIENTE DE ESTE MES! ATTE CERATTI TEC");*/
     
  
};


//FUNCION PARA LLENAR EL FORMULARIO GGOLE FORMS AUTOMATICAMENTE
function llenarFormulario(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojabusqueda = libro.getSheetByName('Formulario');
  var hojadatos = libro.getSheetByName('KYC');
  //var valorbusqueda = Browser.inputBox('Captura el nombre');
  var valorbusqueda = hojabusqueda.getRange("d4").getValue();
  
  //var valorbusqueda = hojabusqueda.getRange('a6').getValue();
  //Logger.log(valorbusqueda);
  var listaBusqueda = hojadatos.getRange('d2:h').getValues();
  //Logger.log(listaBusqueda);  
  
  //var listaNombres = listaBusqueda.map(function (nombre){return nombre[0]});
    var listaNombres = listaBusqueda.map(nombre=> nombre[0]);
  //Logger.log(listaNombres);
  var indice =listaNombres.indexOf(valorbusqueda);
  
  //Logger.log(indice);
  
   var rfc = listaBusqueda[indice][1];
   hojabusqueda.getRange('d6').setValue(rfc);
   var curp = listaBusqueda[indice][2];
   hojabusqueda.getRange('d8').setValue(curp);
   var telCasa = listaBusqueda[indice][4];
   hojabusqueda.getRange('d10').setValue(telCasa);
   var celular = listaBusqueda[indice][3];
  //Browser.msgBox(celular);
   hojabusqueda.getRange('d12').setValue(celular);
  
  
};



//FUNCION PARA BUSCAR UN VALOR COMO BUSCAR V PRO CON MACROS
function buscarV(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojabusqueda = libro.getSheetByName('EDICION DATOS CLIENTES');
  var hojadatos = libro.getSheetByName('LISTAS DINAMICAS');
  //var valorbusqueda = Browser.inputBox('Captura el nombre');
  var valorbusqueda = hojabusqueda.getRange("E5").getValue();
  
  //var valorbusqueda = hojabusqueda.getRange('a6').getValue();
  //Logger.log(valorbusqueda);
  var listaBusqueda = hojadatos.getRange('AI2:AL').getValues();
  //Logger.log(listaBusqueda);  
  
  //var listaNombres = listaBusqueda.map(function (nombre){return nombre[0]});
    var listaNombres = listaBusqueda.map(nombre=> nombre[0]);
  //Logger.log(listaNombres);
  var indice =listaNombres.indexOf(valorbusqueda);
  
  //Logger.log(indice);
  var celular = listaBusqueda[indice][1];
  //Browser.msgBox(celular);
  hojabusqueda.getRange('E13').setValue(celular);
  var nombre = listaBusqueda[indice][0];
  hojabusqueda.getRange('B9').setValue(nombre);
   var direccion = listaBusqueda[indice][2];
   hojabusqueda.getRange('H9').setValue(direccion);
   var index = listaBusqueda[indice][3];
   hojabusqueda.getRange('a9').setValue(index);
  
  
};


//funcion para ediatr los generales de dajer clientes

function editar_clientes(){
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDICION DATOS CLIENTES");
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var listas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LISTAS DINAMICAS");
    var activo = hoja.getRange("E5");
    var dato_modificado = hoja.getRange("j13");
    var rangoClear = hoja.getRange("a9:k13");
    var rangoOrigen = hoja.getRange("c19:e19").getValues();
    //Logger.log(rangoOrigen);
    var fila = hoja.getRange("a9").getValue();
    var rangoDestino = listas.getRange(fila,35,1,3);


    if(listas.getName() == "LISTAS DINAMICAS"){
    rangoDestino.setValues(rangoOrigen);

    Browser.msgBox('Datos modificados con exito!');
    //rangoClear.clearContent();
    activo.activate();
    buscarV();



    }else{
      Browser.msgBox('ERROR EN LA EDICION');
    }


};






function abrirCedula(){

    const libro = SpreadsheetApp.getActiveSpreadsheet();

    const hojaCliente = libro.getSheetByName('LOGIN');
    //var ultimaFila = hojaCliente.getRange('al2').getValue();
    var cliente = hojaCliente.getRange('A3' ).getValue();
    var  hojaBuscada = cliente;
    //console.log(cliente);

    let hojas = libro.getSheets();
    hojas.forEach((hoja,index) => {
      if(hoja.getName() == hojaBuscada){
        hoja.showSheet();
        //console.log('hoja encontrada' + hoja.getName() + "N HOJA :"+ index);
        var hojaactiva =libro.setActiveSheet(libro.getSheetByName(hojaBuscada),true);
      }
    }); 

};



//<<<<<<<<<<<<<<<<<<<----------------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//FUNCIONES PARA HABILITAT CEDULA PARA LOS CASOS
//FUNCION PARA LLAMAR A LA CEDULA DE EL CASO NUEVO PARA CREAR CEDULA NUEVA O RENOVACION


function crearCedula(){
 const libro = SpreadsheetApp.getActiveSpreadsheet();

 const hojaCliente = libro.getSheetByName('KYC');
 var ultimaFila = hojaCliente.getRange('al2').getValue();

 //var cliente = hojaCliente.getRange("d"+ (ultimaFila -1)).getValue();
 var cliente = hojaCliente.getRange("d"+ ultimaFila).getValue();
 var caso = hojaCliente.getRange("am"+ ultimaFila).getValue();
 var credito = hojaCliente.getRange("AB"+ ultimaFila).getValue();
 var referenciaCelda = hojaCliente.getRange("C"+ ultimaFila).getA1Notation();
 var nombre = cliente +"_"+ credito;
 var  hojaBuscada = cliente;

  if(caso >1){
      let hojas = libro.getSheets();
      
      hojas.forEach((hoja,index) => {
      if(hoja.getName() == hojaBuscada){
        hoja.showSheet();
        var hojaCedula =libro.setActiveSheet(libro.getSheetByName(hojaBuscada),true);
        notify("Cedula encontrada a nombre de :" + cliente);
        copiarCedula();
        
      }


     });


      }else{
        Browser.msgBox("CLIENTE NUEVO, CREAR CEDULA ON REFERENCIA  ");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var cedulaNueva = ss.setActiveSheet(ss.getSheetByName('CEDULA NUEVA'),true);
        cedulaNueva.getRange('c10').setValue("=" + hojaCliente.getName()  + "!" + referenciaCelda);
        
      }
  
 

};


function copiarCedula(){
    //hojaactiva
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaActiva = libro.getActiveSheet();
    var lastRow = hojaActiva.getLastRow();
    var cliente = hojaActiva.getName();
    var filas = hojaActiva.getRange("O" + lastRow).getValue()-1;
    var filaIinicio = lastRow- filas;

    console.log(cliente + "-" + lastRow+"/"+ filaIinicio); 

    var filasAcopiar = hojaActiva.getRange("o" + lastRow).getValue();
    
    


    var rangoCopiar = hojaActiva.getRange(filaIinicio,1,filasAcopiar,15);
    var rangoDestino = hojaActiva.getRange(lastRow+5,1,filasAcopiar,15);

    rangoCopiar.copyTo(rangoDestino);

    /*console.log(hojaActiva.getName());
    console.log(cliente);
    console.log(filasAcopiar);
    console.log(lastRow); 
    console.log(inicio);*/

};




//funcion para insertar los valores pro considerando la anotacion osea un rango de celdas de otra hoja

function conectarCedula(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();

  //KYC HOJA DATOS
  const hojaCliente = libro.getSheetByName('KYC');
  var ultimaFila = hojaCliente.getRange('al2').getValue();
  var clienteAnterior =  hojaCliente.getRange("d"+ (ultimaFila-1)).getValue();
  var cliente = hojaCliente.getRange("d"+ ultimaFila).getValue(); 

  //HOJA ACTIVA , QUE DEBE SER LA CEDULA DE EL CLIENTE ABIERTA
  var hojaActiva = libro.getActiveSheet();
  var nombreHoja = hojaActiva.getName();


  if(cliente == nombreHoja | clienteAnterior == nombreHoja){

      const lastRow = hojaActiva.getLastRow();
   
      var nCredito = hojaActiva.getRange('k'+ (lastRow -1)).getValue();  
      
      var nombre = nombreHoja +"_"+nCredito;
    

      var lista1 = hojaActiva.getRange(lastRow-1,1,1,1 ).getA1Notation();
      var lista2 = hojaActiva.getRange(lastRow-1,2,1,1 ).getA1Notation();
      var lista3 = hojaActiva.getRange(lastRow-1,3,1,1 ).getA1Notation();
      var lista4 = hojaActiva.getRange(lastRow-1,4,1,1 ).getA1Notation();
      var lista5 = hojaActiva.getRange(lastRow-1,5,1,1 ).getA1Notation();
      var lista6 = hojaActiva.getRange(lastRow-1,6,1,1 ).getA1Notation();
      var lista7 = hojaActiva.getRange(lastRow-1,7,1,1 ).getA1Notation();
      
      var lista8 = hojaActiva.getRange(lastRow-1,8,1,1 ).getA1Notation();
      var lista9 = hojaActiva.getRange(lastRow-1,9,1,1 ).getA1Notation();
      var lista10= hojaActiva.getRange(lastRow-1,10,1,1 ).getA1Notation();
      var lista11= hojaActiva.getRange(lastRow-1,11,1,1 ).getA1Notation();

      var lista12 = hojaActiva.getRange(lastRow,1,1,1 ).getA1Notation();
      var lista13 = hojaActiva.getRange(lastRow,2,1,1 ).getA1Notation();
      var lista14 = hojaActiva.getRange(lastRow,3,1,1 ).getA1Notation();
      var lista15 =  hojaActiva.getRange(lastRow,4,1,1 ).getA1Notation();
      var lista16 = hojaActiva.getRange(lastRow,5,1,1 ).getA1Notation();

      var dato1=  hojaActiva.getRange(lastRow-4,1,1,1 ).getA1Notation();
      var dato2=  hojaActiva.getRange(lastRow-3,1,1,1 ).getA1Notation();
      var dato3=  hojaActiva.getRange(lastRow-2,1,1,1 ).getA1Notation();  
      

      //SICDJR HOJA DE DATOS CARTERA SE REGISTRAN LOS PAGOS DE CREDITOS
      const sicdjr = libro.getSheetByName('SICDJR');
      //const ultimaFila = sicdjr.getRange('A1').getValue()-1;se usa para casos que no sena la ultima fila
      const ultimaFila = sicdjr.getRange('A1').getValue();
      
      sicdjr.getRange(ultimaFila,10,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista1);
      sicdjr.getRange(ultimaFila,11,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista2);
      sicdjr.getRange(ultimaFila,12,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista3);
      sicdjr.getRange(ultimaFila,13,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista4);
      sicdjr.getRange(ultimaFila,14,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista5);
      sicdjr.getRange(ultimaFila,15,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista6);
      sicdjr.getRange(ultimaFila,16,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista7);

      sicdjr.getRange(ultimaFila,19,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista8);
      sicdjr.getRange(ultimaFila,20,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista9);
      sicdjr.getRange(ultimaFila,21,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista10);
      sicdjr.getRange(ultimaFila,22,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista11);

      sicdjr.getRange(ultimaFila,25,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista12);
      sicdjr.getRange(ultimaFila,26,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista13);
      sicdjr.getRange(ultimaFila,27,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista14);
      sicdjr.getRange(ultimaFila,28,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista15);
      sicdjr.getRange(ultimaFila,29,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista16);
    
      

      
      
      //HOJA DE CONTROL DE PAGOS DEL MES EN CURSO , SE CONTABILIZA A BALANCE
      var control = libro.getSheetByName(mesActual);
      var lastColumn = control.getRange('a1').getValue() + 3;
      var multiplo = control.getRange(39,lastColumn+1).getA1Notation();

      
      control.getRange(1,lastColumn + 1,1,1).setValue(nombre);
      control.getRange(40,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato1 +"*"+multiplo);
      control.getRange(41,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato2 +"*"+multiplo);
      control.getRange(42,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato3 +"*"+multiplo);

      
      //HOJA DE VENCIMIENTOS PARA INSERTAR NUEVO CALENDARIO
      const calendario = libro.getSheetByName('VENCIMIENTOS');
      const ultimafilaTabla = calendario.getRange('A25').getValue();
      const filas = calendario.getRange('d23').getValue();
      const tabla = calendario.getRange(27,4,filas,5).getValues();


      //console.log(nombre + "/"+ lista1 + "/" + lista2 + "/" + lista3);
      //console.log(nombre + "/"+ dato1 + "/" + dato2 + "/" + dato3);
      //console.log("sicdjr ultima fila " + ultimaFila + "control ultima columna " + lastColumn);
    
      notify("¡La conexion del credito a nombre de : " + nombreHoja + " con n de Credito " + nCredito + " fue Exitosa!");
   

     }else{

       notify("¡La cedula activa , no corresponde a el caso a dar de alta, verifique el proceso!");

     }

    
   

};


//FUNCION PARA LIMPIAR FILAS DE ACUERDO A UN VALOR PARA HOJA VENCIMIENTOS CADA INICIO DE MES 
function limpiarCalendario(){
      const libro = SpreadsheetApp.getActiveSpreadsheet();
     //HOJA DE VENCIMIENTOS PARA INSERTAR NUEVO CALENDARIO
      const calendario = libro.getSheetByName('VENCIMIENTOS');
      const ultimafilaTabla = calendario.getRange('A25').getValue();
      const filas = calendario.getRange('d23').getValue();
      const tabla = calendario.getRange(27,4,filas,5);
      const condicion = calendario.getRange(27,10,filas,1).getValues();
      const fechas = calendario.getRange(27,14,filas,1).getValues();

      //console.log(tabla);

      condicion.forEach((vigente,index) =>{
        
        if(vigente == 0){
          //console.log(index + "fila");        
          
          var dato =calendario.getRange(index+27,4,1,5).clearContent();
          //var rango = calendario.getRange("A" + (index+27));
          //calendario.hideRow(rango);
          
          
          //console.log(dato);
         
        }

        
      })
      notify('Los registros han sido limpiados.. ');
};



  //FUNCION PARA COPIAR FECHAS DE INICIO DE MES

  function copyinicioMes(){
      const libro = SpreadsheetApp.getActiveSpreadsheet();
     //HOJA DE VENCIMIENTOS PARA INSERTAR NUEVO CALENDARIO
      const calendario = libro.getSheetByName('VENCIMIENTOS');
      const hojaActiva = libro.getActiveSheet();
      const ultimafilaTabla = calendario.getRange('A25').getValue();
      const filas = calendario.getRange('d23').getValue();
      const tabla = calendario.getRange(27,4,filas,5);
     
      const fechas = calendario.getRange(27,14,filas,1).getValues();
      const fechasInicio = calendario.getRange(27,4,filas,1).getValues();

      var rangoOrigen = calendario.getRange(27,14,filas,1).getValues();

      rangoOrigen.forEach((inicio,index) =>{
        if(inicio > 0){
           var rangoDestino = calendario.getRange(27,4,filas,1).setValues(rangoOrigen);
          // console.log(inicio);
      
        }

      })
      
     
      //notify('Se ha actualizado la lista de inicio de mes ...');
      //console.log(rangoOrigen)


        
     
      
  };












