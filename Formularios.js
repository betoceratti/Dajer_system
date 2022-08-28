

//FUNCION PARA EL REGSITRO DE DATOS MEDIANTE FORMULARIO EN HOJA DE GASTOS

function gastos(){
  
  var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
  var lastRow = gastos.getRange('q2').getValue();
  var ro = gastos.getRange('Y1:aC1').getValues();
  var fecha = gastos.getRange('c4').getValue();
  var monto = gastos.getRange('c9').getValue();
  var categoria = gastos.getRange('c13').getValue();
  var concepto = gastos.getRange('c17').getValue();
  //Logger.log(categoria);
  //Logger.log(concepto);
  var rd = gastos.getRange(lastRow + 1,18, 1,5);
  
  if(fecha != "" && monto >0 && categoria != "" && concepto != ""){
  
    rd.setValues(ro);
    Browser.msgBox('Gasto registrado con exito'); 
    
    var alerta = SpreadsheetApp.getUi();
    var respuesta = alerta.alert('Deseas continuar agregando gastos?', alerta.ButtonSet.YES_NO);
    if(respuesta == "YES"){
      var limpiar = gastos.getRange('c4:c21').clearContent();
      gastos.getRange('c3').activate();
      
    }else{
       var limpiar = gastos.getRange('c4:c21').clearContent();
       gastos.hideSheet();
      var Dajer = SpreadsheetApp.getActiveSpreadsheet();
     var hojaactiva =Dajer.setActiveSheet(Dajer.getSheetByName('LOGIN'),true);
     }
  
  }else{
    Browser.msgBox('Datos incompletos, revisa tu captura');
    gastos.getRange('c3').activate();
 }

};


//FUNCION PARA EDITAR EL CAMPO AHORRO

//MODAL PARA AHORRO
function ahorroForm(){
  var html =HtmlService.createHtmlOutputFromFile('ahorro');
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(html, 'A T T I'); 


};

function actualizar(importe,nota){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTERA VIGENTE');
  var kyc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KYC');
  var sicdjr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SICDJR');
  var importe = importe;
  var nota = nota;
  var cliente = hoja.getRange('f62').getValue();
  var credito = hoja.getRange('d62').getValue();
  var filakyc = hoja.getRange('E62').getValue();
  var filasicdjr = hoja.getRange('E63').getValue();
  var columnakyc = hoja.getRange('H62').getValue();
  var columnasicdjr =hoja.getRange('H63').getValue();
  //Logger.log(sicdjr);
  
  //RANGOS DESTINO
  var kycdestino = kyc.getRange(filakyc, columnakyc);
  var sicdjrdestino = sicdjr.getRange(filasicdjr,columnasicdjr);
  //insertamos los valores
    kycdestino.setValue(importe);
    kycdestino.setNote(nota);
    sicdjrdestino.setValue(importe);
    sicdjrdestino.setNote(nota);

};




//FUNCION AHORRO CON IMPUTS DE GOOGLE
function update(){
  //HOJAS USADAS
  
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CARTERA VIGENTE');
  var kyc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KYC');
  var sicdjr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SICDJR');
  
  // RANGOS A UTILIZAR ORIGEN
   var google =SpreadsheetApp.getUi();
   var importe = parseFloat(Browser.inputBox('A T T I --- SYSTEM', 'Ingresa el monto nuevo', google.ButtonSet.OK));
   var nota = Browser.inputBox('A T T I--- SYSTEM', 'Capturar nota', google.ButtonSet.OK); 
  
  //var importe = hoja.getRange('G62').getValue();
  //var nota = hoja.getRange('J62').getValue();
  var cliente = hoja.getRange('f62').getValue();
  var credito = hoja.getRange('d62').getValue();
  var filakyc = hoja.getRange('E62').getValue();
  var filasicdjr = hoja.getRange('E63').getValue();
  var columnakyc = hoja.getRange('H62').getValue();
  var columnasicdjr =hoja.getRange('H63').getValue();
  //Logger.log(sicdjr);
  
  //RANGOS DESTINO
  var kycdestino = kyc.getRange(filakyc, columnakyc);
  var sicdjrdestino = sicdjr.getRange(filasicdjr,columnasicdjr);
  //Logger.log(kycdestino.getValue());
  //Logger.log(sicdjrdestino.getValue());
   
  //EDITAR CAMPO AHORRO E INSERTAMOS UNA NOTA EN LA MISMA CELDA
   if(hoja.getName()== 'CARTERA VIGENTE' && nota != ""){
    kycdestino.setValue(importe);
    kycdestino.setNote(nota);
    sicdjrdestino.setValue(importe);
    sicdjrdestino.setNote(nota);
    let notificacion ='Se actualizo el campo ahorro de el credito n .' + '' + credito + '' +' por importe de :' + '' + importe + ' '+ 'pesos' + ' ' + 'A nombre de :'+ ''+ cliente;
    notify(notificacion);
    //var borrar1 = hoja.getRange('G62').clearContent();
    //var borrar2 = hoja.getRange('J62').clearContent();
    hoja.getRange('l43').activate();
 
  }else{
    Browser.msgBox('Debes capturar el campo NOTA ');
     hoja.getRange('L43').activate();
  } 
  
 }; 
  


//Funcion para alta de clientes Dajer

function formulario(){
  
 var formulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FORM'); 
 
 var BaseData =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KYC');
 var ultimafila = BaseData.getRange('AL2').getValue(); 

  //hoja de plantilla
  var libro = SpreadsheetApp.getActiveSpreadsheet();  
  var hoja = libro.getActiveSheet();
  var sheet = libro.getSheetByName('PLANTILLA_EMAIL');  
  var lastRow = sheet.getRange('f1').getValue(); 

  
  var rangoOrigen = formulario.getRange('A41:X41').getValues();
  var rangoDestino = BaseData.getRange(ultimafila +1, 4,1,24);
  var nombre = formulario.getRange('d4').getValue();
  var monto = formulario.getRange('k10').getValue();
  var plazo = formulario.getRange('k14').getValue();
  
  
  //Logger.log(rangoDestino);
  //Logger.log(ultimafila);
  //Para insertar una fila completa tambien podemos guaradr todas las variebles en un arreglo de arreglos
   var cliente =[[nombre,monto,plazo]]
  
  if(nombre != "" && monto>0){
    rangoDestino.setValues(rangoOrigen);
    var limpiar1 = formulario.getRange('D4:D24').clearContent();
    var limpiar2 = formulario.getRange('K4:K17').clearContent();
    var limpiar3 = formulario.getRange('K19:K20').clearContent();
    var limpiar4 = formulario.getRange('K23:K29').clearContent();
    sheet.getRange(lastRow,2,1,3).setValues(cliente);
    //envioEmail();
  
    Browser.msgBox('Alta exitosa !'+ ' ' + ' Cliente :'+ ' '+ nombre);
    var  libro = SpreadsheetApp.getUi();
    var respuesta =libro.alert('Deseas agregar a otro cliente ?', libro.ButtonSet.YES_NO);
    
      if(respuesta == 'YES'){
    
      formulario.getRange('D4').activate();
      
      } else if (respuesta == 'NO'){
       formulario.hideSheet();
       var Dajer = SpreadsheetApp.getActiveSpreadsheet();
       var hojaactiva =Dajer.setActiveSheet(Dajer.getSheetByName('LOGIN'),true);
      
      }           
    
     }else{
       
       Browser.msgBox('Revisa tu formulario, al parecer los datos estan incompletos');
       formulario.getRange('D4').activate();
    
    } 

};







//FUNCION PARA ENVIO DE CORREOS DESDE SHEETS
function envioEmail(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();  
  var hoja = libro.getActiveSheet();
  var sheet = libro.getSheetByName('PLANTILLA_EMAIL');  
  var lastRow = sheet.getRange('f1').getValue();
  
  var destinatario = sheet.getRange(lastRow,1).getValue();
  var nombre = sheet.getRange(lastRow,2).getValue();
  var monto = sheet.getRange(lastRow, 3).getValue();
  var plazo = sheet.getRange(lastRow,4).getValue();
  var alta = sheet.getRange(lastRow,5).getValue();  
  var plantilla = sheet.getRange('f5').getValue();
  var asunto = "Alta de caso a nombre de " + nombre;

  //en caso de envio de muchos obtener todos os contactos
  var contactos = sheet.getRange(372, 1, 2, 5).getValues();
  
  
  //PARA ENVIO DE MAILS A VARIOS DESTINATARIOS AL MISMO TIEMPO USAMOS CLICLO FOREACH
  //Logger.log(contactos);
  /*contactos.forEach(function(fila){
    Logger.log(fila[0]);
    GmailApp.sendEmail(fila[0], 'prueba','body')
  })*/
  
  
  //Con este manera el cuerpo del correo se va sin formato , todo pegado asi que haremos una plantilla, lo mismo podriamos hacer coin asunto 
  //var body = "Alta de credito a nombre de : " + nombre + "con fecha " + alta +"por un monto de :  "+ monto + " comentarios : " + comentarios;
  //plantilla con replace
  var body = plantilla.replace('{nombre}', nombre).replace('{fecha}', alta).replace('{monto}', monto).replace('{plazo}',plazo);
 
  //Logger.log(body);
  if(nombre != "" && monto != ""){
     var email = GmailApp.sendEmail(destinatario, asunto,body)
     Browser.msgBox('Email enviado con exito"');
  
  
  }else{
   Browser.msgBox('El mail no se ha enviado,hubo un error.');
  
  } 
 
};






//FUNCION PARA ACCESAR MEDIANTE USUARIO Y PASSWORD

function login(usuario,password){
   let tablero = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TABLERO CONSULTA');
   let lastRow = 10;
   var usuario = usuario;
   var password = password;  
   var message = "Pago de sistema PENDIENTE!, vencimiento 15/mes.";
    // var usuario = Browser.inputBox('Captura tu usuario');
    //var password =Browser.inputBox('Captura tu contraseña');
    if ( usuario == 'CERATTI' & password == 2702) {
    let notificacion ='Hola' + " " + usuario + " " + "Bienvenido ya tienes acceso al menu! ";
    notify(notificacion);
    crearMenu();
    //var libro =SpreadsheetApp.getActive();
    //var hoja = libro.setActiveSheet(libro.getSheetByName('TABLERO CONSULTA'),true);
    //hoja.getRange(4,16).setValue(usuario);
    //WHITCOLOR();
    //hoja.getRange('p4').activate().setFontFamily('Comfortaa').setFontSize(14).setFontColor('red');
    /*MAIL = GmailApp.sendEmail('robertoceratti@gmail.com', "Alerta seguriy CERATTI-PYTHON ","EL USUARIO : DARIO BARRIOS HA ACCESADO ALA SISTEMA")*/

    } else if(usuario == 'Adbr' & password == 1306){
    let notificacion ='HOLA' + " " + "DARIO BARRIOS" + " " + "YA TIENES ACCESO AL MENU ";
    notify(notificacion);
    //Browser.msgBox(" PAGO PENDIENTE SISTEMA SICA-DAJER  ");
    var libro =SpreadsheetApp.getActive();
    var hoja = libro.setActiveSheet(libro.getSheetByName('DASHBOARD DIGITAL'),true);
    hoja.getRange('k1').setValue(usuario);
    hoja.getRange('k1').activate().setFontFamily('Comfortaa').setFontSize(14).setFontColor('white');
    menuDario();
    /*Browser.msgBox("CERATTI TEC" + " " + " Le recordamos que esta pendiente el pago por el servicio del sistema, vencimiento 15/cada mes ");*/
    MAIL = GmailApp.sendEmail('robertoceratti@gmail.com', "Alerta seguriy CERATTI-PYTHON ","EL USUARIO : " + usuario + " HA ACCESADO AL SISTEMA ");
    } else {
    Browser.msgBox('datos incorrectos,si olvidaste tus claves contacta al admin del sistema o verifica tu captura (sensible a mayusculas y minusculas)');
    var libro =SpreadsheetApp.getActive();
    var hoja =libro.setActiveSheet(libro.getSheetByName('login'),true);
    hoja.getRange(1,12).setValue(usuario);
    hoja.getRange('A1').activate();
    }

 
};//termina funcion
  




//FUNCION PARA CONECTAR HTML MEDIANTE LA CREACION DE UN MODAL PARA EL LOGIN


function form(){
 
  var fileweb = "ATTI"  
  var html = HtmlService.createHtmlOutputFromFile(fileweb);
  var modal = SpreadsheetApp.getUi();
    
  modal.showModalDialog(html, 'DAJER ');
  // modal.showSidebar(html);
  
};



//MODAL DE GASTOS BOOTSTRAP Y JAVA SCRIPT

function formGastos(){
    //Browser.msgBox('¡ MANTENIMIENTO DE SISTEMA PENDIENTE !') ;
    var hojaGastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
    var listado = hojaGastos.getRange('b152:b179').getValues();
    var gadmon = hojaGastos.getRange('B59:B73').getValues();
    var gventas = hojaGastos.getRange('f59:f70').getValues(); 
    var ahorro = hojaGastos.getRange('j59:j70').getValues();  
    //Browser.msgBox("Pago De Sistema SICA DAJER ++PENDIENTE++ ");
        
    var fileweb = "GASTOS"  
    var html = HtmlService.createTemplateFromFile('gastosBootstrap'); 
      //html.listado = listado;
      
      
      const pagina = html.evaluate();  
            pagina.setHeight(350).setWidth(400);

      var modal = SpreadsheetApp.getUi();  
      modal.showModalDialog(pagina, 'Control Gastos ');
    //modal.showSidebar(pagina);
    // modal.showSidebar(html);
  
};






//FUNCION PARA PROCESAR Y GAURADR OS GASTOS 


function gastosModal(data){
   var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GC');
  var lastRow = gastos.getRange('q2').getValue();
  //var gastos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGIN');
  //var lastRow = gastos.getRange('q1').getValue();
  
  var gasto =[[data.fecha,data.importe,data.categoria,data.concepto,data.nota]];
  //Logger.log(categoria);
  //Logger.log(concepto);
  //var rd = gastos.getRange(lastRow + 1,18, 1,5);
  //var rd = gastos.getRange('B29:f29');
  //var fila = gastos.getRange(29, 2, 1, 5);
  
   var fila= gastos.getRange(lastRow + 1,18, 1,5);
    fila.setValues(gasto);
    //Browser.msgBox('Gasto registrado con exito'); 
    
    
  

};









