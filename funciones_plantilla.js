
//FUNCION PARA DESPLEGAR UN HTML CON LLENADO EXIGIBLES, PAGOS CEDULAS ALTAS GASTOS DEL DIA

function exigibles(){

  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LISTAS DINAMICAS');
  var hojaGastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
  var hojaColab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRESTAMOS COLABORADORES');

   
  var filaInicio = hojaData.getRange('T1').getDisplayValue();
  var nFilas= hojaData.getRange('X1').getDisplayValue();
  var filas = hojaData.getRange('af1').getValue();
  var rangoData =hojaData.getRange(1, 20 ,nFilas, 4).getDisplayValues();
  let suma = rangoData.reduce((suma,monto)=> suma + monto[1],0);  
  //console.log(suma);

  
  //ALGORITMO PARA PAGOS HOY
  var hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mesActual);
  var fI = hojaOrigen.getRange('c199').getDisplayValue();
  var nF = hojaOrigen.getRange('b199').getDisplayValue();
  var  pagosHoy =hojaOrigen.getRange(200,2,nF,2).getValues();
  //Logger.log(pagosHoy);
  let = nPagos = pagosHoy.reduce((contar,elemento)=>  contar + 1,0);
  //console.log(rangoData);

  //SECCION PÁRA VISUALIZAR EDICIONES DE CELDAS CLICKEO

  var clickeos = hojaData.getRange(2,31,filas,1).getValues();
  //console.log(clickeos);

  //SECCION PARA CAPTURAR LOS GASTOS DE EL DIA
  var nRows = hojaGastos.getRange('ak1').getValue();
 var gastosHoy = hojaGastos.getRange(2,36,nRows,4).getDisplayValues();

 //SECCION PRA CAPTURAR OPERACION DEL DIA DE COLABORADORES
  var nRowscolab = hojaColab.getRange('aa1').getValue();
 var colabHoy = hojaColab.getRange(2,26,nRowscolab,4).getDisplayValues();


  var totales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TC');
  var totalExigibles = totales.getRange('E13').getDisplayValue();
  var numero = totales.getRange('H22').getDisplayValue();
   //var rangoData = [nombre,pago];
  //Logger.log(rangoData);
 var plantilla = HtmlService.createTemplateFromFile("Exigibles");
    //plantilla.nombres = nombres;
    //plantilla.ids = ids;
     plantilla.rangoData = rangoData;
     plantilla.totalExigibles = totalExigibles;

     
     plantilla.numero = numero;
     plantilla.pagosHoy = pagosHoy;
     plantilla.clickeos = clickeos;
     plantilla.gastosHoy = gastosHoy;
     plantilla.colabHoy = colabHoy;
  
  const pagina = plantilla.evaluate();
  pagina.setWidth(550).setHeight(400);
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, "A T T I")

   /* var alerta = SpreadsheetApp.getUi();
      var respuesta =  alerta.alert("Deseas descargar en PDF ?",alerta.ButtonSet.YES_NO);

       if(respuesta == 'YES'){       
         
        descargarPdf(pagina); 
     
      
      } else{        

        var plantilla = HtmlService.createTemplateFromFile("Exigibles");
   
            plantilla.rangoData = rangoData;
            plantilla.totalExigibles = totalExigibles;
             plantilla.pagosHoy = pagosHoy;
            plantilla.numero = numero;
          
        const pagina = plantilla.evaluate();
        pagina.setWidth(500).setHeight(400);
        
        const ui = SpreadsheetApp.getUi();
        ui.showModalDialog(pagina, "A T T I")
                  
      
      }     */           
  

};  
  

//FUNCION PARA ACTUALIZAR POR COLUMNA EL VIGENTE  
  
function actualizar(columna){
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTERA VIGENTE');
  columna = columna;
  var numeroColumna = hojaData.getRange('e1');

  numeroColumna.setValue(columna);


};



//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CARTERA VIGENTE

function vigentes(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTERA VIGENTE');
  //var saldo = Math.trunc(hojaData.getRange('k4').getDisplayValue());
  var saldo = hojaData.getRange('k4').getDisplayValue();
  var lastRow = hojaData.getRange('C1').getDisplayValue();
  var rangoData = hojaData.getRange(6,2, lastRow,13).getValues();
  
  let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
  Logger.log(nCasos);
  
  var titulo = "FUNCIONA"
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("VIGENTES");
  plantilla.rangoData = rangoData;
  plantilla.nCasos = nCasos;
  plantilla.saldo = saldo;
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
 
  
};



//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CASOS RECIEN OTORGADOS

function casosRecientes(){
  

  let fechaMax = new Date ("November 30, 2021");
  let dateMin = new Date("November 01, 2021");
  let porFecha = rango => rango[16] > dateMin && rango[16] < fechaMax;//filter
  let porYear = rango => rango[16] > dateMin;//filter

    var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KYC');
    
    var lastRow = hojaData.getRange('AL2').getDisplayValue();
  
    var rangoData = hojaData.getRange(lastRow -29,4, 30,24).getDisplayValues();

    var filtroxNombre = rangoData.filter(filtrado => filtrado[0] == "SALVADOR LOPEZ SANCHEZ");
    let rangoxFecha = rangoData.filter(mes => mes[16] == "18/10/2021");
    
    //let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
    //Logger.log(rangoData[0][16] );
    //console.log(dateMin);    
    
  

    var plantilla = HtmlService.createTemplateFromFile("RECIENTES");
    plantilla.rangoData = rangoData;
    plantilla.ultimaFila = lastRow;
    plantilla.filtroxNombre = filtroxNombre;
    plantilla.rangoxFecha = rangoxFecha;

    const pagina = plantilla.evaluate();
    pagina.setWidth(950).setHeight(500);
    
    
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
    
 
  
};


//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML TABLA DE VENCIMIENTOS

function tablaVencimientos(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VENCIMIENTOS');
  var total = hojaData.getRange('C13').getDisplayValue();
  var lastRow = hojaData.getRange('AK1').getDisplayValue();
 
  var rangoData = hojaData.getRange(2,38, lastRow,9).getValues();
  var calendario = hojaData.getRange(3,1,13,6).getDisplayValues();
  var mes = hojaData.getRange('A17').getValue();
  
  //let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
  Logger.log(calendario);
  
  var titulo = "FUNCIONA"
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("TABLE");
  plantilla.rangoData = rangoData;
  plantilla.calendario = calendario;
  plantilla.ultimaFila = lastRow;
  plantilla.total = total;
  plantilla.mes = mes;
  
  
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
 
  
};





function vencidos(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTERA VIGENTE');
  var lastRow = hojaData.getRange('Q1').getDisplayValue();
  var rangoData = hojaData.getRange(6,16, lastRow,29).getValues();
  //Logger.log(rangoData);
  

  let fechaMax = new Date ("Jun 30, 2021");
  let dateMin = new Date("January 01, 2021");


  //index 10 monto vencido
  //index 9 vigente
  let resultado = rangoData.reduce((suma,monto)=> suma + monto[10],0);//total vencido
  let filtradoxvigente = rangoData.filter(filtro => filtro[9] > 0);//total vencido filtrando vigentes
  let filtradoxfecha = rangoData.filter(filtro => filtro[2] > dateMin);//filtrado por rango de fechas
  let fechas = filtradoxfecha.map(fecha => fecha[0]);//impresion de los nombres de el filtrado 
  let total = filtradoxfecha.reduce((suma,monto)=> suma + monto[10],0);//suma de vencido poe el filtro de fecha  
  let totalClientes = filtradoxfecha.reduce((suma,monto)=> suma + 1,0);//conteo de clientes vencidos por el filtro
  
  //for (var i = 0; i <rangoData.length; i++)
   Logger.log(totalClientes);

  var plantilla = HtmlService.createTemplateFromFile("VENCIDOS");
  plantilla.rangoData = rangoData;
  plantilla.total = total;
  plantilla.resultado = resultado;
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
    
};



function clientesAhorro(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LISTAS DINAMICAS');
   var saldo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TC');
  var filaInicio = hojaData.getRange('AB1').getDisplayValue();
  var nFilas = hojaData.getRange('AD1').getDisplayValue();
  var rangoData =hojaData.getRange(1, 28 ,nFilas, 2).getValues();
  var ahorro = saldo.getRange('k9').getDisplayValue();
  
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("ahorroClientes");
  plantilla.rangoData = rangoData;
  plantilla.ahorro = ahorro;
  const pagina = plantilla.evaluate();
  pagina.setWidth(500).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
    
};



//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CARTERA VIGENTE
//DE LOS GASTOS
function reporteGastos(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
  var lastRow = hojaData.getRange('Q1').getDisplayValue();
  var rangoData = hojaData.getRange(2,27, lastRow-33,9).getValues();
  rangoData.sort(function(a, b){return b-a}); 
  //Logger.log(rangoData);
  
  //variables para filter
  let category = "GASTOS_ADMON";//GASTOS_DE_VENTA,SALIDA_AHORRO GASTOS_ADMON
  let descripcion = "SISTEMA MAN";//APOYO SOCIOS, SISTEMA MAN EJEC COBRANZA
  let fechaMax = new Date ("January 31, 2021");
  let dateMin = new Date("January 01, 2021");



  //Predicados
  let paraMap = monto => monto[0];//map
  let paraReduce = (acumulado, monto) => acumulado + monto[1];//reduce
  let paraCategory = categoria => categoria[2] == category;//filter
  let paraConcepto = concepto => concepto[3] == descripcion;//filter
  let porFecha = rango => rango[0] > dateMin & rango[0] < fechaMax;//filter
  let porYear = rango => rango[0] > dateMin;//filter
  
  let datosFila = rangoData.map(paraMap);

  let categorias = rangoData.filter(paraCategory);  
  let conceptos = rangoData.filter(paraConcepto);
  let xconceptoyFecha = conceptos.filter(porFecha);
  let xcategoriayFecha = categorias.filter(porFecha);
  let rangosFecha = rangoData.filter(porFecha);

  let tXcategoria = categorias.reduce(paraReduce,0);
  let tXconcepto = conceptos.reduce(paraReduce,0);
  let txRango = rangosFecha.reduce(paraReduce,0);
  let txrangoyconcepto = xconceptoyFecha.reduce(paraReduce,0);
  let txrangoycategoria = xcategoriayFecha.reduce(paraReduce,0);
  //let tGastos = rangoData.reduce((suma,monto )=> suma + monto[1],0);
  let tGastos = rangoData.reduce(paraReduce,0);
  
  
  let socios = 50000;
  let sistema = 24000;

  let maximo = [socios,sistema]

  
    console.log(tGastos);
  //console.log(Math.max.apply(null,maximo));



  var plantilla = HtmlService.createTemplateFromFile("reporte_gastos");
  plantilla.rangoData = rangoData;
  plantilla.tGastos = tGastos;
  const pagina = plantilla.evaluate();
  pagina.setWidth(900).setHeight(500);
 
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I_SYSTEM ')
    
};




//FUNCION PARA DESPLEGAR LA PALNTILLA DE INDICADORES
//DASHBOARD
function indicadores(){

  var sgcc = SpreadsheetApp.getActive();
  var hoja =sgcc.getSheetByName('TC');
 
  var BG=sgcc.getSheetByName('LOGIN');
  
  var porcentajevencido= hoja.getRange(13,10).getValue()*100;
  var montovencido = hoja.getRange('h13').getDisplayValue();
  var caja= hoja.getRange(13,11).getDisplayValue();
   var dif=BG.getRange(1,11).getDisplayValue();
  var cashflow = BG.getRange('K4').getDisplayValue();
  var cv=hoja.getRange(9,5).getDisplayValue();
  var vhoy=hoja.getRange(13,5).getDisplayValue();
   var clvi=hoja.getRange(17, 14).getDisplayValue();
  /*var clve=RP.getRange(2, 8).getDisplayValue();*/
  var ahorro = hoja.getRange('K9').getDisplayValue();
  var bancos = hoja.getRange('k17').getDisplayValue();
  
 
  var nclientes=BG.getRange(1,13).getDisplayValue();
  
    //Logger.log(rangoData);
    let  page =  "Indicadores";
      
      
    var html =HtmlService.createTemplateFromFile(page);
    html.caja = caja;
    html.porcentajevencido = porcentajevencido.toFixed(0);
    html.cv = cv;
    html.clvi = clvi;
    html.dif = dif;
    html.cashflow = cashflow;
    html.montovencido = montovencido;
    html.ahorro = ahorro;
    html.bancos = bancos;
   // html.prueba = prueba;
    const pagina = html.evaluate();
      pagina.setHeight(600).setWidth(700);
  
  
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(pagina, 'A T T I  SYSTEM ')  
    

};


//FUNCION PARA DESPLEGAR INIDICADORES CONTABLES POR AÑO
//INGRESOS Y GASTOS
function contables(){
  var hojatablero =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TC');
  
  var datos = hojatablero.getRange('C'+2+':I'+5).getDisplayValues();
  var esperado =hojatablero.getRange('m10:p12').getDisplayValues();
  
  var ingresos = hojatablero.getRange('M8').getDisplayValue();
  var gastos = hojatablero.getRange('N8').getDisplayValue();
  var utilidad = hojatablero.getRange('O8').getValue();
  var hoy = hojatablero.getRange('e22').getDisplayValue();
  var mes = hojatablero.getRange('e20').getValue();
  // html.bancos = new Intl.NumberFormat().format(bancos);
  //Logger.log(datos);
  var year = datos.map(function(años){ return(años[6])});


  var html =HtmlService.createTemplateFromFile('balance');
      html.datos = datos;
      html.esperado = esperado;
      html.ingresos = ingresos;
      html.gastos = gastos;
      html.utilidad = utilidad;
      html.hoy = hoy;
      html.mes = mes;
    
    //   html.utilidad = new Intl.NumberFormat().format(utilidad);
  
  const pagina =html.evaluate();
        pagina.setHeight(500).setWidth(550);
  
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'A T T I');

};




//FUNCION PARA CONVERTIR A PDF UN HTML 

function pdfHtml(){
  var hojatablero =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TC');
  
  var datos = hojatablero.getRange('C'+2+':I'+4).getValues();
  var esperado =hojatablero.getRange('m10:p12').getValues();
  var ingresos = hojatablero.getRange('M8').getDisplayValue();
  var gastos = hojatablero.getRange('N8').getDisplayValue();
  var utilidad = hojatablero.getRange('O8').getDisplayValue();
  var hoy = hojatablero.getRange('e22').getDisplayValue();
  // html.bancos = new Intl.NumberFormat().format(bancos);
  //Logger.log(datos);
  var year = datos.map(function(años){ return(años[6])});

    //archivo a convertir
    let html = HtmlService.createTemplateFromFile('balancepdf');
    html.datos = datos;
      html.esperado = esperado;
      html.ingresos = ingresos;
      html.gastos = gastos;
      html.utilidad = new Intl.NumberFormat().format(utilidad);
      html.hoy = hoy;
      
    
    
    let web = html.evaluate();    
    let file = web.getAs('application/pdf');


    var idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";
    //llamamos nuetra carpeta y ahi guardamos el pdf 
    var carpetaMaestra = DriveApp.getFolderById(idCarpeta);
    var dia = new Date().getDate();
    var mes = new Date().getMonth()+1;
    let pdf =carpetaMaestra.createFile(file).setName("Reporte" + dia+"/"+mes );
    var link = pdf.getUrl();

    var formula = 'HYPERLINK("'+ link+'";"PDF")'
    //sheet.getRange('g2').setFormula(formula);

    /*
    let enlace = HtmlService.createHtmlOutput('<a href="'+ link +'">DOWMLOAD PDF</a>').setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(150).setHeight(150);


    var modal = SpreadsheetApp.getUi();

    modal.showModalDialog(enlace,'A T T I ');*/

};



//FUNCION PARA OBTENER EL URL DE EL ULTIMO ARCHIVO CREADO EN LA CARPETA
function getUltimaDescarga(){
    var idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";
    //llamamos nuetra carpeta y ahi guardamos el pdf 
    var carpetaMaestra = DriveApp.getFolderById(idCarpeta);
    var files =carpetaMaestra.getFiles().next();

    var nombre = files.getName();
    var id = files.getId();
    var link = files.getUrl();
    var dia = new Date().getDate();
    var mes = new Date().getMonth()+1;
    console.log(nombre + "/ url : "+ link);
    //console.log(dia+"/"+ mes);
    //console.log(files.getId());
    var formula = 'HYPERLINK("'+ link+'";"'+ nombre +'")'
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = libro.getSheetByName('LOGIN');
    sheet.getRange('B21').setFormula(formula);


};




function cashflow(){
  
  var hojaData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOF');
  var fluujo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGIN');
  
  //var saldo_inicial = Math.trunc(hojaData.getRange('as5').getDisplayValue());  
  var saldo_inicial = hojaData.getRange('as5').getDisplayValue();
  var ingresos = hojaData.getRange('AV6').getDisplayValue();
  var salidas = hojaData.getRange('AV10').getDisplayValue();
  var flujo = hojaData.getRange('AS20').getDisplayValue();
  var diferencia = fluujo.getRange('K4').getDisplayValue();
  var rangoData = hojaData.getRange('AR7:AS19').getValues();
  
  
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("Cashflow");
  plantilla.rangoData = rangoData;
  plantilla.saldo_inicial = saldo_inicial;
  plantilla.ingresos = ingresos;
  plantilla.salidas = salidas;
  plantilla.flujo = flujo;
  plantilla.diferencia =diferencia;

  const pagina = plantilla.evaluate();
  pagina.setWidth(600).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
    
};




  



  
//FUNCION PARA SEPARAR EL NEW DATE
function fechas(){
  var login = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGIN');
  var valor = login.getRange('I4').getDisplayValue();
  
  Logger.log(valor);
 
  
  var time = new Date();
   var dia = time.getDate();
   var mes = time.getMonth();
       mes++;
   var year = time.getFullYear();
  
  var fecha = [dia,mes,year];
  //Logger.log(fecha);



};


//COTIZADOR EN LINEA 

function cotizadorForm(){
  var html = HtmlService.createTemplateFromFile('cotizador');
  const pagina =html.evaluate();
          pagina.setHeight(500).setWidth(550);
    
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'COTIZADOR DAJER');



};



function cotizarr(){
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



//FUNCION PARA AP0LICAR PAGOS SIN ABRIR CEDULAS GENIAL

function pagarCedula(cliente,importe){
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaLista = libro.getSheetByName('LISTAS DINAMICAS');
    const fila = hojaLista.getRange('af1').getValue();

    var cliente = cliente;
    var importe = importe;
    var  hojaBuscada = cliente;
    let hojas = libro.getSheets();

    hojas.forEach((hoja,index) => {
      
    if(hoja.getName() == hojaBuscada){
    var  lastRow = hoja.getLastRow();
    var nFilas = hoja.getRange('l'+ (lastRow-1)).getValue();
    var filas = hoja.getRange('n'+lastRow).getValue()-1;
    var calendario = hoja.getRange(lastRow-filas,4,nFilas,4).getDisplayValues();
    var vigente = hoja.getRange('B'+ (lastRow-1)).getDisplayValue();
    var vencido = hoja.getRange('C'+ (lastRow-1)).getDisplayValue();
    var pagosVencidos = hoja.getRange('A'+ (lastRow-1)).getDisplayValue();
    var nCredito = hoja.getRange('k'+ (lastRow -1)).getValue();
    var montoInicial = hoja.getRange(lastRow-(filas-1),3,1,1).getDisplayValue();
    var fechaInicial = hoja.getRange(lastRow-(filas-2),3,1,1).getDisplayValue();
    var pagados = hoja.getRange('m'+ (lastRow -1)).getValue();
    var proximoPago = hoja.getRange(lastRow-(filas-pagados),4,1,1).setValue(true);
    var moratorios = hoja.getRange(lastRow-(filas-pagados),11,1,1).setValue(importe); 
    let ubicacion =  proximoPago.getRow();
   
  
    
   }  
  
  hojaLista.getRange(fila+1,31,1,1).setValue("Se edito la celda : " + " Caso " + cliente);

 }); 


};

    //ELIMINA LA LISTA DE CLEDAS EDITADAS HAY Q PONERLE UN TRIGGER PARA LO HAGA DIARIO 
function borrarLista(){
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaLista = libro.getSheetByName('LISTAS DINAMICAS');
    const fila = hojaLista.getRange('af1').getValue();
    const rango = hojaLista.getRange('ae2:ae20').clearContent();

};


//MODAL PARA BISQUEDA DE CEDULA
function buscarCliente(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();

  const hojaLista = libro.getSheetByName('LISTAS DINAMICAS');
  const ultimaFila = hojaLista.getRange('L54').getValue();

  //const listado = hojaLista.getRange('B' + 6 + ':B'+ ultimaFila).getDisplayValues();
  const listado = hojaLista.getRange( 54,13,ultimaFila,1).getDisplayValues();

  //console.log(listado); 

  var html =HtmlService.createTemplateFromFile('buscador');
        html.listado = listado; 
        const pagina =html.evaluate();
          pagina.setHeight(200).setWidth(370);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');



};








//ESTADOS DE CUENTA
function edodeCuenta(cliente){

    const libro = SpreadsheetApp.getActiveSpreadsheet();

  var cliente = cliente;
  var  hojaBuscada = cliente;
  let hojas = libro.getSheets();

  hojas.forEach((hoja,index) => {
  
    if(hoja.getName() == hojaBuscada){
    var  lastRow = hoja.getLastRow();
    var nFilas = hoja.getRange('l'+ (lastRow-1)).getValue();
    var filas = hoja.getRange('n'+lastRow).getValue()-1;
    var calendario = hoja.getRange(lastRow-filas,4,nFilas,4).getDisplayValues();
    var vigente = hoja.getRange('B'+ (lastRow-1)).getDisplayValue();
    var vencido = hoja.getRange('C'+ (lastRow-1)).getDisplayValue();
    var pagosVencidos = hoja.getRange('A'+ (lastRow-1)).getDisplayValue();
    var nCredito = hoja.getRange('k'+ (lastRow -1)).getValue();
    var montoInicial = hoja.getRange(lastRow-(filas-1),3,1,1).getDisplayValue();
    var fechaInicial = hoja.getRange(lastRow-(filas-2),3,1,1).getDisplayValue();
    var pago = hoja.getRange(lastRow-(filas-7),3,1,1).getDisplayValue();
    

    console.log(montoInicial + "" + fechaInicial);

    var html =HtmlService.createTemplateFromFile('CEDULA');
       
        html.calendario = calendario;  
        html.cliente = cliente;
        html.vigente = vigente;
        html.vencido = vencido;
        html.pagosVencidos = pagosVencidos;
        html.nCredito = nCredito;
        html.montoInicial = montoInicial;
        html.fechaInicial = fechaInicial;
        html.pago = pago;
  const pagina =html.evaluate();
          pagina.setHeight(500).setWidth(550);
    
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'A T T I');
   
   
   /*  */ 
    
   }  
  
 }); 



};


//MODAL PARA RENTABILIDAD X CLIENTE
function rentModal(){

    const libro = SpreadsheetApp.getActiveSpreadsheet();

    const hojaLista = libro.getSheetByName('KYC');
    const ultimaFila = hojaLista.getRange('AQ1').getValue();

    const listado = hojaLista.getRange('AN' + 2 + ':AN'+ ultimaFila).getDisplayValues();

    //console.log(listado); 

    var html =HtmlService.createTemplateFromFile('buscador2');
      html.listado = listado; 
      const pagina =html.evaluate();
          pagina.setHeight(200).setWidth(370);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

};



//HTML RENTABILIDAD
function rentabilidad(cliente){

  const libro = SpreadsheetApp.getActiveSpreadsheet();


  const hojaCliente = libro.getSheetByName('CARTERA VIGENTE');
  var nombreCliente = hojaCliente.getRange('B41'); 
  //var cliente = hojaCliente.getRange('A3' ).getValue();
  var cliente = cliente;

  nombreCliente.setValue(cliente);

  var lastRow = hojaCliente.getRange('C41').getValue();
  var  rangoData = hojaCliente.getRange(45,2,lastRow,11).getDisplayValues();
  const evaluacion = hojaCliente.getRange('F39').getDisplayValue();
  const totales = hojaCliente.getRange('b43:L43').getDisplayValues();
  const mensaje = hojaCliente.getRange('h40').getDisplayValue();
  const totalRecup = hojaCliente.getRange('h42').getDisplayValue();


  var html =HtmlService.createTemplateFromFile('Rentabilidad');
        
      html.rangoData = rangoData; 
      html.cliente = cliente;
      html.evaluacion = evaluacion;
      html.totales = totales;
      html.mensaje = mensaje;
      html.totalRecup = totalRecup;
      
    const pagina =html.evaluate();
            pagina.setHeight(500).setWidth(850);
         
           
            var modal = SpreadsheetApp.getUi();
                modal.showModalDialog(pagina, 'A T T I');      
                 
  
};



//MODAL PARA APLICAR PAGOS EN CNTROL DE PAGOS DEL MES
function pagosModal(){
    const libro = SpreadsheetApp.getActiveSpreadsheet();

    const hojaLista = libro.getSheetByName(mesActual);

    const ultimaFila = hojaLista.getRange('B138').getValue();
    const filaIinicio = hojaLista.getRange('A5').getValue();
    const listado = hojaLista.getRange(139,3,ultimaFila,1).getValues();

    //console.log(listado); 

    var html =HtmlService.createTemplateFromFile('pagosMes');
          html.listado = listado; 
          const pagina =html.evaluate();
              pagina.setHeight(300).setWidth(380);
              var modal = SpreadsheetApp.getUi();
                  modal.showModalDialog(pagina, 'A T T I');

};





//FUNCION PARA APLICAR PAGOS EN EL CONTROL DE MES MEDIANTE FORMULARIO 

function pagos(nombre,importe,nota){
    let book = SpreadsheetApp.getActiveSpreadsheet();
    let hojaDatos = book.getSheetByName(mesActual);
    let hojaActiva = book.getActiveSheet();
    var celdaActiva = hojaDatos.getActiveCell();
    var filaActiva = celdaActiva.getRow();
    var columActiva = celdaActiva.getColumn();
    //let valorBuscaddo = hojaDatos.getRange('b125').getValue();
    let valorBuscaddo = nombre;
    //let importe = parseFloat(Browser.inputBox('Captura el monto a pagar:')) ;
    let ultimaFila = hojaDatos.getRange('c122').getValue() + 1;
    let listaBusqueda = hojaDatos.getRange(1,4,1,hojaDatos.getLastColumn()).getValues();

    //console.log(listaBusqueda[0]);
    //console.log(valorBuscaddo);


    listaBusqueda.forEach(cliente => {

    if(mesActual == hojaDatos.getName() && valorBuscaddo != ""){
    var indice = cliente.indexOf(valorBuscaddo) +4;
    //console.log(indice);
    //celdaActiva = hojaDatos.getRange(ultimaFila, indice).activate();
    celdaActiva = hojaDatos.getRange(ultimaFila, indice).setValue(importe);
    var anotacion = hojaDatos.getRange(ultimaFila, indice).setNote(nota);
    var mensaje = "El pago se ha procesado correctamente a nombre de : " + valorBuscaddo;
    notify(mensaje);


    }


    }
    );


};




//MODAL PARA INDICADORES
function kipisModal(){

  const libro = SpreadsheetApp.getActiveSpreadsheet();

  const hojaLista = libro.getSheetByName('PYTHON');
  const ultimaFila = hojaLista.getRange('E1').getValue();

  const listado = hojaLista.getRange(2,1,ultimaFila,2).getDisplayValues();

  //console.log(listado); 

  var html =HtmlService.createTemplateFromFile('KPIS');
        html.listado = listado; 
        const pagina =html.evaluate();
            pagina.setHeight(500).setWidth(610);
            var modal = SpreadsheetApp.getUi();
                modal.showModalDialog(pagina, 'A T T I');

      


};


//MODAL PARA PRESTAMO COLABORADORES
function colabModal(){

        const libro = SpreadsheetApp.getActiveSpreadsheet();
        const hojaLista = libro.getSheetByName('PRESTAMOS COLABORADORES');
        const sheet = libro.getSheetByName('LOGIN');
        const ultimaFila = hojaLista.getRange('A1').getValue();

        const listado = hojaLista.getRange(ultimaFila-9,1,10,7).getDisplayValues();
        const cabeceras = hojaLista.getRange(1,1,1,7).getDisplayValues();
        var saldoAaron = hojaLista.getRange('e1').getDisplayValue();
        var saldoRob = hojaLista.getRange('f1').getDisplayValue();
        var saldoJesus = hojaLista.getRange('g1').getDisplayValue();

        //console.log(saldoAaron); 

        var html =HtmlService.createTemplateFromFile('colaboradores');

              html.listado = listado; 
              html.cabeceras = cabeceras;
              html.saldoAaron = saldoAaron;
              html.saldoRob = saldoRob;
              html.saldoJesus = saldoJesus;

              var pagina =html.evaluate();
                  pagina.setHeight(500).setWidth(900);
                  var modal = SpreadsheetApp.getUi();
                      modal.showModalDialog(pagina, 'A T T I');
   

     

};




function descargarPdf(page){
   //var page = Browser.inputBox("Captura la plantilla ");
  
   const libro = SpreadsheetApp.getActiveSpreadsheet();  
   const sheet = libro.getSheetByName('LOGIN');
 
   var pagina = page;
   var file = pagina.getAs('application/pdf');


      var idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";
      //llamamos nuetra carpeta y ahi guardamos el pdf 
      var carpetaMaestra = DriveApp.getFolderById(idCarpeta);
      var dia = new Date().getDate();
      var mes = new Date().getMonth()+1;
      let pdf =carpetaMaestra.createFile(file).setName("Reporte" + dia+"/"+mes );
      var link = pdf.getUrl();

      var formula = 'HYPERLINK("'+ link+'";"REPORTE_PDF")'
      sheet.getRange('C4').setFormula(formula);
      sheet.getRange('C4').activate();

};



