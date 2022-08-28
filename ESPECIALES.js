//OBTENEMOS DATOS DE LOS LIBROS (editores, sus emails , etc )
function dataFile(){
    let book = SpreadsheetApp.getActiveSpreadsheet();
    let hojas = book.getSheets();
    let libro = SpreadsheetApp.getActive();
    let hojaDatos = book.getSheetByName('INFO_SISTEMA');



    //console.log("El propietario de este archivo es :"+ " "+ book.getOwner().getEmail());
    //console.log("Nombre de sistema : "+ " " + libro.getName());
    let editores = book.getEditors();
    //console.log("Los editores autorizados son  : ");

    hojaDatos.getRange('A1').setValue("El propietario de este archivo es :"+ " "+ book.getOwner().getEmail());

    hojaDatos.getRange('a2').setValue("Nombre de sistema : "+ " " + libro.getName());
    hojaDatos.getRange('a3').setValue("Consta de   : " + hojas.length + " "+ "hojas");
    hojaDatos.getRange('a4').setValue("Los editores autorizados son  : ");

    editores.forEach(fila=>{
      console.log(fila.getEmail());
      const autorizados = [fila.getEmail()];
      hojaDatos.appendRow(autorizados);


    })

};




//OBTENEMOS TODAS LAS HOJAS CON SUSU NOMBRES Y INDEX
function getSheets(){
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = libro.getSheets();

  hojas.forEach((hoja,index)=> console.log(hoja.getName() + " " + "n de hoja :" + index ));

};


//OCULTA HOJAS
function ocultarSheets(){
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    //hojas.forEach((hoja,index) =>{console.log(hoja.getName()+"/ n hoja /"+ index)})

    /*hojas.forEach((hoja,index)=>{
      if(index > 0){
        hoja.hideSheet();
        
      }
    })*/

    hojas.forEach((hoja,index)=>{
      if(hoja.getName() != "DAJER MOVIL!" && hoja.getName() != "DB"){
        hoja.hideSheet();
        
      }
    })

};





//ELIMINA HOJAS 
function deleteHojas(){
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    //console.log(datoLibro);

    //hojas.forEach((hoja,index) =>{console.log(hoja.getName()+"/ n hoja /"+ index)})
    //console.log(libros);
    hojas.forEach((fila,index)=>{
      if(index === 4){
        libro.deleteSheet(fila);
      }
    })

};




//funcion para buscarv con rango dinamico
function formula() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGIN');
  var a = hoja.getRange('f1').getValue();
  var b = hoja.getRange('g1').getValue();
  var rango = "SICDJR"+'!'+a+b
  console.log(rango);
  var pegar =hoja.getRange('k34');
  pegar.setFormula('=vlookup(B32;'+ rango +';4;0)');
  
};




//Funcion para ocultar filas de SICDJR


function ocultar() {
  
  var hoja = SpreadsheetApp.getActive();
  
  var ultimaFila = hoja.getRange('A4').getValue();
  
  //Logger.log(ultimaFila);  
  
  hoja.getActiveSheet().hideRows(7, ultimaFila);
  
  hoja.getRange('G1').activate();
};


function display(){
  var hoja = SpreadsheetApp.getActive();
  var ultimaFila = hoja.getRange('A4').getValue();
  
  hoja.getActiveSheet().showRows(7, ultimaFila);
  hoja.getRange('G1').activate();
  
};




//Funcion para crear un intervalo con nombre dinamico

function Intervalo() {
  var hojaIntervalo = SpreadsheetApp.getActive().getSheetByName('LOGIN');
  
  spreadsheet.getRange('A385:A391').activate();
  spreadsheet.setNamedRange('lista', spreadsheet.getRange('A385:A391'));
};


//FUNCION PARA INSERTAR NOTAS

function NOTAS() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C17').activate();
  spreadsheet.getRange('D15').setNote('GFGFGFGF');
};



//FUNCIONES PARA LIMPIAR, REMPLAZAR ELIMINAR , INSERTAR FILAS MEDIANTE UN VALOR EXISTENTE EN CIERTA FILA

function clearRow(){

var hojaPruebas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pruebas');
var lastRow =   hojaPruebas.getLastRow();

//Logger.log(lastRow);
  
  for(i=2; i<= lastRow; i++){
    if(hojaPruebas.getRange(i, 1).getValue() == 'ROBERTO'){
    
      hojaPruebas.getRange('A'+ i + ':C'+ i).clear();
    }   
      
 }  



};


function deleteRow(){

var hojaPruebas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pruebas');
var lastRow =   hojaPruebas.getLastRow();

//Logger.log(lastRow);
  
  for(i=2; i<= lastRow; i++){
    if(hojaPruebas.getRange(i, 1).getValue() == 'ROBERTO'){
    
      hojaPruebas.deleteRow(i);
    }   
      
 }  



};


function insertRow(){

var hojaPruebas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pruebas');
var lastRow =   hojaPruebas.getLastRow();

//Logger.log(lastRow);
  
  for(i=2; i<= lastRow; i++){
    if(hojaPruebas.getRange(i, 1).getValue() == 'ROBERTO'){
      
      //Logger.log(i);
      hojaPruebas.insertRowAfter(i);
    }   
      
 }  



};


function valuesInRow(){

var hojaPruebas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pruebas');
var lastRow =   hojaPruebas.getLastRow();

//Logger.log(lastRow);
      var nombre = "ERCIK";
      var edad = 35;
      var sexo = "M"
  for(i=2; i<= lastRow; i++){
    if(hojaPruebas.getRange(i, 1).getValue() == 'ROBERTO'){
      
      //Logger.log(i);
      
      var cliente = [[nombre,edad,sexo]];
      //Logger.log(cliente);
      //hojaPruebas.getRange(i, 1, 1, 3).setValues(cliente);
      //con stringAnotaion
      //hojaPruebas.getRange('A'+ i + ':C' + i).setValues([['maria',40, 'f']]);
    }   
      
 }  



};

//METODOS DE PROGRMACION FUNCIONAL

function metodos(){


const libro = SpreadsheetApp.getActiveSpreadsheet();
const genrealesClientes = libro.getSheetByName("KYC");
var ultimaFila = genrealesClientes.getRange('AL2').getValue();
//let data = genrealesClientes.getRange('D7:H20').getValues();
let data = genrealesClientes.getRange('D' + 2 +':S'+ ultimaFila ).getValues();

//////////////////////////

const cartera = libro.getSheetByName('SICDJR');
let lastRow = cartera.getRange('A1').getValue();
let dataCartera = cartera.getRange('A'+ 6 +':V' + lastRow).getValues();


let reduced = (suma, monto) => suma + monto[2];

let filtrado = dataCartera.filter(filtro => filtro[2] > 15000);



let n = filtrado.map(nombre => nombre[1] + " " + nombre[2]);
let summ = filtrado.reduce(reduced,0);
console.log(n);
console.log(summ);



var personas =[ {
    nombre: "CERATTI",
    edad : 35,
    sexo: "HOMBRE",
    
  },
    {
    nombre: "pepe",
    edad : 30,
    sexo: "HOMBRE"
    },

    {
    nombre: "lili",
    edad : 30,
    sexo: "MUJER"
    },

    {
    nombre: "LORENA",
    edad : 40,
    sexo: "MUJER"
    }
    
   ];


   let customers =[
     ["roberto", 2000],
      ["lili", 10000],
       ["pedro", 9000],
        ["lola", 7000],
         ["javier", 2000]
   ];

let numeros = [10,30,50,30,90];
let minimo = 1000;
let maximo = 3000;
let valorbuscado = "SALVADOR LOPEZ SANCHEZ";

//ptrdicados que usan los metodos 
let rangoMontos = monto => monto[15] > minimo &  monto[15] < maximo;
let rangoTotal = monto => monto[15] > minimo ;
let totalxClient = cliente => cliente[0] == valorbuscado; 


//traer algunas columnas pro las trae en un solo valor 
//let nombres = data.map((nombre,index)=> nombre[0] +"n credito :"+ (index +1)  + " monto :" + nombre[15]);

//traer una  columnas segub en filtro
let nombres = data.map(nombre => nombre[0]);
let indice = nombres.indexOf(valorbuscado);
//Logger.log(indice);

let names = data[indice][0];
let importe = data[indice][15];
//Logger.log("Mi nombre es :" + names );
//Logger.log(importe);

//encuentra el primer elemento que coincida con el valor buscado
let nombre = data.find(nombre => nombre[0] == "JOSE PEREZ SANCHEZ");


//formas de hacerlo en un sola linea de formula filtrando primero el array original
let mayores5 = data.filter(rangoMontos).map(nombre => nombre[0]);

//usando dos metodos separados recomendado
let mayores6 = data.filter(rangoTotal);
let datosNombre_monto = mayores6.map(nombre => nombre[0]+ "monto:" + "" + nombre[15]);

//metodo reduce

let montoTotal = mayores6.reduce((contador,monto)=>{

return  contador + monto[15];


},0);


let Cliente = data.filter(totalxClient);
let customer = Cliente.find(nombre => nombre[0] == valorbuscado);
const totalCliente = Cliente.reduce((acumulado,monto)=> acumulado + monto[15],0);

nombres.sort();
datosNombre_monto.sort();
//console.log(datosNombre_monto);
//console.log(montoTotal);
//console.log("Nombre :" + customer);
//console.log("Prestamo Total $" + totalCliente);

//Recomendacion filtra primero con filter para reduce a menos que se quiera sumar todo


};

//CLASES MATEMATICAS MATH Y NUMBER

function calculos(){

//CLASE MATH
let valor = Math.PI;


//CLASE NUMBER

let numero = Number.MAX_VALUE;

//console.log(valor);
//console.log(numero);

//metodos especiales de Math

//round redondea segun la tendencia de valor  ejemplo 2.3 = 2 , 2.6 = 3
let redondeo = Math.round(2.3);

//rendodeo hacia arriba
let arriba = Math.ceil(2.4);
//rendodeo hacia abajo
let haciaAbajo = Math.floor(8.95);


//console.log(redondeo);
//console.log(arriba);
//console.log(haciaAbajo);

//rednodeo con decimales y cantidad de decimales

let decimal = 12/13;
let multiplicador = 1000;
//usando un multiplicador con dos decimales nos dara resultado de dos decimales,si usaramos 1000 daria 3 decimales

let a = Math.round(decimal * multiplicador)/ multiplicador;
let b = Math.ceil(decimal * multiplicador)/ multiplicador;
let c = Math.floor(decimal * multiplicador)/ multiplicador;

//template html y js
console.log(`${a} - ${b} - ${c}`);


//otro metodo para usar decimales desinar cuantos y redondea los decimales segun si tendencia

let sueldo = 1456.345;

console.log(sueldo.toFixed(2));

// metodo truncar simplemente quita decimales y deja en numero entero

console.log(Math.trunc(3.456));

//calcular potencias y raiz cuadrada

//potenciar un numero
console.log(Math.pow(3,3));

//raiz cuadrada
console.log(Math.sqrt(16));

//generar numeros aleatorios de un rango de numeros

console.log(Math.floor( Math.random() *10));


};

//FUNCIONA PARA CALCULAR EL MAX DE UNOS MONTOS

function numeroMayor(a,b,c){
  let mayor = 0;

  if(a>b){
    mayor = a;
  }else{
    mayor=b;
  }

  if(c > mayor){
 mayor = c;

  }
   
  

return mayor;

}

console.log(numeroMayor(3,5,1));



//funcion para obtener valores unicos 
function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index;

// usage example:
var a = ['a', 1, 'a', 2, '1'];
var unique = a.filter( onlyUnique ); // returns ['a', 1, 2, '1']

console.log(unique);

}

//FUNCION PRA RELLENAR UNA LISTA DESPLEGFABLE DE UN FORM DESDE SHEETS

function getDta(){
  var libro = SpreadsheetApp.getActiveSpreadsheet(); 
  var hojadatos = libro.getSheetByName('SCD'); 
  var data = hojadatos.getRange('BV2:BV').getValues();
  //console.log(data)
  let clientes = [];

  data.forEach(fila=> clientes.push(fila));

//console.log(clientes)

return clientes;
};

//funcion para accesar al formulario y sus items 
//[0].getId();
function getForm(){
 let formulario = FormApp.openById('1EqREiHRPkmYq_IPFQUOJnFb86D0PZ3DrlWG4WOfJPQQ'); 
let titulo = FormApp.openById('1EqREiHRPkmYq_IPFQUOJnFb86D0PZ3DrlWG4WOfJPQQ').getItems()[26].getTitle();
let id = FormApp.openById('1EqREiHRPkmYq_IPFQUOJnFb86D0PZ3DrlWG4WOfJPQQ').getItems()[26].getId();
//console.log(pregunta);

let item = formulario.getItemById(id).asListItem();
let opciones = getDta();
let clientes = ['roberto','maria','PEPE','lolo'];
//console.log(clientes);
item.setChoiceValues(clientes);


};


//FUNCION OARA OBTENER EL USUARUIO EDITOR

function getEditor(e){
 var range = e.range;
 var hoja = e.source;
 

 user = e.user;

 notify(user);

};



