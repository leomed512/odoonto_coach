    var hojaVista = ss.getSheetByName("Vista Previsiones");
    if (!hojaVista) {
        hojaVista = ss.insertSheet("Vista Previsiones");
        configurarVistaPrevisiones(hojaVista, ss);
    }
    return hojaVista;
}
// Configurar hoja de vista de previsiones
function configurarVistaPrevisiones(hojaVista, ss) {
    //hojaVista.clear();

    hojaVista.getRange("A1:B1").setValue("Previsiones");
    hojaVista.getRange("A2").setValue("Año:");
    hojaVista.getRange("A3").setValue("Mes:");
        // Configurar título "Previsiones"
    hojaVista.getRange("A1:B1")
        .setValue("Previsiones")
        .setFontSize(20)
        .setFontWeight("bold")
        .setBackground("#00c896")
        .setHorizontalAlignment("center")
        .merge();

    // Configurar "Año:"
    hojaVista.getRange("A2")
        .setValue("Año:")
        .setFontSize(12)
        .setFontWeight("bold")
        .setBackground("#999999")
        .setFontColor("white");

    // Configurar "Mes:"
    hojaVista.getRange("A3")
        .setValue("Mes:")
        .setFontSize(12)
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white");
        
    // Alinear el contenido de B3 a la derecha
    hojaVista.getRange("B3").setHorizontalAlignment("right");

 ////////////////////////////LEO
////INSTRUCCIONES
    hojaVista.getRange(1, 5, 1, 6).merge().setValue("INSTRUCCIONES PARA REGISTRAR PAGOS ")
        .setFontSize(12)
        .setBackground("#00c896")
        .setFontColor("#424242")
        .setFontWeight('bold')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center");
      hojaVista.getRange(2, 5, 1, 6).merge().setValue("Identifica la transacción que deseas ejecutar, ingresa el ABONO, TIPO DE PAGO. Si el pago será divido en cuotas ingresa la próxima fecha en que se realizará otro pago.")
      .setFontSize(10)
      .setBackground("#424242")
      .setFontColor("#FFFFFF")
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");

      hojaVista.getRange(3, 5, 1, 6).merge().setValue("Selecciona la fila completa de la transacción, ve a la barra de Menú > Cobros > Ejecutar cobro en fila seleccionada")
      .setFontSize(10)
      .setBackground("#b8b6b6")
      .setFontColor("#080808")
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");

// Validaciones de año y mes 
var hojaStaging = ss.getSheetByName("Staging Previsiones");
var datos = hojaStaging.getDataRange().getValues();
var annos = new Set();

// procesar correctamente todas las fechas
for (var i = 1; i < datos.length; i++) {
    if (datos[i][1]) {
        var fecha;
        // Manejar tanto objetos Date como strings de fecha
        if (datos[i][1] instanceof Date) {
            fecha = datos[i][1];
        } else {
            fecha = new Date(datos[i][1]);
        }
        // Verificar que la fecha sea válida antes de extraer el año
        if (!isNaN(fecha.getTime())) {
            annos.add(fecha.getFullYear());
        } else {
            Logger.log("Fecha inválida encontrada en la fila " + (i + 1));
        }
    }
}

// Asegurarnos de que el Set se convierte correctamente a array
var annosArray = Array.from(annos).sort((a, b) => a - b);
var annosArrayString = annosArray.map(String);
console.log(typeof annosArrayString);
var validacionAnno = SpreadsheetApp.newDataValidation()
    .requireValueInList(annosArrayString)
    .setAllowInvalid(false)
    .build();
hojaVista.getRange("B2").setDataValidation(validacionAnno);


var meses = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    "Ver todo el año"  // Añade esta nueva opción
];

var validacionMes = SpreadsheetApp.newDataValidation()
    .requireValueInList(meses)
    .setAllowInvalid(false)
    .build();
hojaVista.getRange("B3").setDataValidation(validacionMes);

hojaVista.getRange("P1").setValue("Fechas de control");
hojaVista.getRange("P2:P3").setValues([["Fecha inicio"], ["Fecha fin"]]);
hojaVista.getRange("P2").setFormula('=IF(B3="Ver todo el año",DATE(B2,1,1),DATE(B2,MATCH(B3,{"Enero";"Febrero";"Marzo";"Abril";"Mayo";"Junio";"Julio";"Agosto";"Septiembre";"Octubre";"Noviembre";"Diciembre"},0),1))');
hojaVista.getRange("P3").setFormula('=IF(B3="Ver todo el año",DATE(B2,12,31),EOMONTH(P2,0))');

// Ocultar las columnas de control
hojaVista.hideColumns(16, 1); // Oculta la columna P


// Encabezados de la tabla
var encabezados = [
    "ID TRANSACCIÓN", "FECHA", "PACIENTE", "DOCTOR", "IMPORTE TOTAL",
    "ABONO", "SALDO PENDIENTE", "TIPO DE PAGO", "PRÓXIMO PAGO", "TRATAMIENTO", "ESTADO"
];

hojaVista.getRange(5, 1, 1, encabezados.length).setValues([encabezados])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("white")
    .setHorizontalAlignment("center");
    // Agregar filtros a los encabezados
    hojaVista.getRange(5, 1, 1, encabezados.length).createFilter();

// Configurar formato de columnas
hojaVista.getRange("E:G").setNumberFormat("€#,##0.00");

// Validación para tipo de pago
var tipoPagoOpciones = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
var validacionTipoPago = SpreadsheetApp.newDataValidation()
    .requireValueInList(tipoPagoOpciones, true)
    .setAllowInvalid(false)
    .build();

// Validación para fecha
var validacionFecha = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Por favor, ingrese una fecha válida')
    .build();


// Obtener el rango de datos desde la hoja Staging
var hojaStaging = ss.getSheetByName("Staging Previsiones");
var datosStaging = hojaStaging.getDataRange().getValues();
var numFilasConDatos = datosStaging.length - 1; // Restar 1 por la fila de encabezado


    // Añadir tabla resumen (igual que antes)
    configurarTablaResumen(hojaVista);
}

function actualizarDropdownAnos() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    
    // Obtener todos los datos de Staging Previsiones
    var datos = hojaStaging.getDataRange().getValues();
    var annos = new Set();
    
    // Procesar las fechas para obtener años únicos
    for (var i = 1; i < datos.length; i++) {
        if (datos[i][1]) {  // Columna B (índice 1) contiene las fechas
            var fecha;
            // Manejar tanto objetos Date como strings de fecha
            if (datos[i][1] instanceof Date) {
                fecha = datos[i][1];
            } else {
                fecha = new Date(datos[i][1]);
            }
            // Verificar que la fecha sea válida antes de extraer el año
            if (!isNaN(fecha.getTime())) {
                annos.add(fecha.getFullYear());
            }
        }
    }
    
    // Convertir Set a array ordenado
    var annosArray = Array.from(annos).sort((a, b) => a - b);
    var annosArrayString = annosArray.map(String);
    
    // Actualizar la validación del dropdown
    if (annosArrayString.length > 0) {
        var validacionAnno = SpreadsheetApp.newDataValidation()
            .requireValueInList(annosArrayString)
            .setAllowInvalid(false)
            .build();
        
        hojaVista.getRange("B2").setDataValidation(validacionAnno);
    }
}
function configurarTablaResumen(hojaVista) {
    var resumenEncabezados = [
        ["RESUMEN", "", ""],
        ["Total Importe", "=SUM(E6:E)", ""],
        ["Total Abonado", "=SUM(F6:F)", ""],
        ["Total Pendiente", "=SUM(G6:G)", ""]
    ];
  ////////////////////////////////////////////////////////LEO  
    var rangoResumen = hojaVista.getRange(1, 12, resumenEncabezados.length, 3);
    rangoResumen.setValues(resumenEncabezados);
  ////////////////////////////////////////////////////////LEO  
    // Aplicar formatos
    hojaVista.getRange("L1:M1")
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold")
        .merge();
    
    hojaVista.getRange("M2:M4").setNumberFormat("€#,##0.00");
    
    // Estilos alternados
    hojaVista.getRange("L2:M2").setBackground("#f6f6f6");
    hojaVista.getRange("L3:M3").setBackground("#e2e2e2");
    hojaVista.getRange("L4:M4").setBackground("#f6f6f6");
    
    // Bordes
    hojaVista.getRange("L1:M4").setBorder(true, true, true, true, true, true);
    ////////////////////////////////////////////////////////////////////////////////
}
function actualizarVistaPrevisiones() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    
    var anno = hojaVista.getRange("B2").getValue();
    var mes = hojaVista.getRange("B3").getValue();
    
if (!anno) return; // Solo requerimos el año

    // Definir fechas de inicio y fin según si hay mes seleccionado
    var fechaInicio, fechaFin;

        if (mes && mes !== "Ver todo el año") {
        // Si hay mes seleccionado y no es "Ver todo el año", usar las fechas calculadas en P2 y P3
        fechaInicio = hojaVista.getRange("P2").getValue();
        fechaFin = hojaVista.getRange("P3").getValue();
    } else {
        // Si solo hay año o se seleccionó "Ver todo el año"
        fechaInicio = new Date(anno, 0, 1); // 1 de enero del año seleccionado
        fechaFin = new Date(anno, 11, 31); // 31 de diciembre del año seleccionado
    }
  // Primero, limpiar todos los datos existentes
    var ultimaFila = hojaVista.getLastRow();
    if (ultimaFila > 5) { // 5 es la fila del encabezado
        hojaVista.getRange(6, 1, ultimaFila - 5, hojaVista.getLastColumn()).clearContent();
    }

    var datosStaging = hojaStaging.getDataRange().getValues();
    var datosFiltrados = datosStaging.filter((row, index) => {
        if (index === 0) return false; // Skip header
        var fecha = new Date(row[1]);
        return fecha >= fechaInicio && fecha <= fechaFin;
    });
    
    if (datosFiltrados.length > 0) {
        hojaVista.getRange(6, 1, datosFiltrados.length, datosFiltrados[0].length)
            .setValues(datosFiltrados);
///////////////////////////////////////////////LEO NUEVO
       hojaVista.getRange(6, 8, datosFiltrados.length).setDataValidation(validacionTipoPago)  // Tipo de Pago
            .setBackground("#f0f0f0");  // Light gray background
        hojaVista.getRange(6, 9, datosFiltrados.length).setDataValidation(validacionFecha)     // Próximo Pago
            .setBackground("#f0f0f0");  // Light gray background
        hojaVista.getRange(6, 6, datosFiltrados.length).setBackground("#f0f0f0");  // Abono column
        hojaVista.getRange(6, 10, datosFiltrados.length).setBackground("#f0f0f0");  // TRATAMIENTO column
            //////////////////////////LEO
            // Añadir este bloque para aplicar validaciones
        var validacionTipoPago = SpreadsheetApp.newDataValidation()
            .requireValueInList(["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"], true)
            .setAllowInvalid(false)
            .build();

        var validacionFecha = SpreadsheetApp.newDataValidation()
            .requireDate()
            .setAllowInvalid(false)
            .setHelpText('Por favor, ingrese una fecha válida')
            .build();
        // Aplicar validaciones a las columnas correspondientes
        hojaVista.getRange(6, 8, datosFiltrados.length).setDataValidation(validacionTipoPago);  // Tipo de Pago
        hojaVista.getRange(6, 9, datosFiltrados.length).setDataValidation(validacionFecha);     // Próximo Pago
    } else {
        hojaVista.getRange(6, 1).setValue("No hay datos para mostrar");
    }
}

// Función helper para verificar si ya existe el ID en una hoja
function existeIdEnHoja(hoja, id) {
    if (!id) return false;
    var datos = hoja.getRange("A:A").getValues();
    var duplicado = datos.some(row => row[0] === id);
    
    if (duplicado) {
        Browser.msgBox("Error", `El ID de transacción ${id} ya existe en la hoja ${hoja.getName()}`, Browser.Buttons.OK);
    }
    
    return duplicado;
}


//// Actualizar lo que sucede al cambiar de estado en parrilla ppal
function actualizarFormatoFila(hoja, fila, estado) {
    var rangoFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn());
    var colores = { 
        "Aceptado": "#54c772", 
        "Pendiente sin cita": "#FF9D23",
        "Pendiente con cita": "#f7f73e", 
        "No aceptado": "#fc4c3d" 
    };
    rangoFila.setBackground(colores[estado] || null);

    var reglaValidacion = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Aceptado", "Pendiente sin cita", "Pendiente con cita", "No aceptado"], true)
        .setAllowInvalid(false)
        .build();
    hoja.getRange(fila, 10).setDataValidation(reglaValidacion);
}

/// tabla resumen de parrilla ppal
function actualizarTablaResumen(hojaMes) {
   // Verificar si la tabla ya existe en la hoja
    var celdaCheck = hojaMes.getRange("C4").getValue();
    var tablaExiste = celdaCheck && celdaCheck.toString().trim().toUpperCase().includes("TOTAL PRESUPUESTADO");

    if (!tablaExiste) {
        var resumenEncabezados = [
            ["ENVIAR SEMANALMENTE", "", ""],  
            ["Gerencia@odontologycoach.cr", "", ""],  
            ["", "IMPORTES", "N° PACIENTES"],  
            ["TOTAL PRESUPUESTADO", "", ""],  
            ["TOTAL ACEPTADO", "", ""],  
            ["TOTAL COBRADO", "", ""],  
            ["PTO MEDIO", "", ""]   
        ];

        hojaMes.getRange(1, 2, resumenEncabezados.length, 3).setValues(resumenEncabezados);
        hojaMes.autoResizeColumns(2, 3); 

        var estilos = [
            ["B1:C1", "#00c896", true], ["B2:C2", "#f2ecff", false], ["C3:D3", "#424242", true, "#FFFFFF"],
            ["B4", "#e2e2e2", true], ["B5", "#f6f6f6", true], ["B6", "#e2e2e2", true], ["B7", "#f6f6f6", true],
            ["C4", "#f6f6f6", true], ["C5", "#e2e2e2", true], ["C6", "#f6f6f6", true], ["C7", "#e2e2e2", true],
            ["D4", "#e2e2e2"], ["D5", "#f6f6f6"], ["D6", "#e2e2e2"]
        ];
        
        estilos.forEach(item => {
            var celda = hojaMes.getRange(item[0]);
            celda.setBackground(item[1]);
            if (item[2]) celda.setFontWeight("bold");
            if (item.length === 4) celda.setFontColor(item[3]);
        });
    }
    var filaInicio = 18;
    var ultimaFila = hojaMes.getLastRow();

    var rangoTotalPresupuestado = hojaMes.getRange(4, 3);
    var rangoTotalAceptado = hojaMes.getRange(5, 3);
    var rangoTotalCobrado = hojaMes.getRange(6, 3);
    var rangoPtoMedio = hojaMes.getRange(7, 3);
    var rangoPacientesPresupuestados = hojaMes.getRange(4, 4);
    var rangoPacientesAceptados = hojaMes.getRange(5, 4);

    rangoTotalPresupuestado.setFormula(`=SUMIF(J${filaInicio}:J${ultimaFila}, "<>No aceptado", K${filaInicio}:K${ultimaFila})`);
    rangoTotalAceptado.setFormula(`=SUMIF(J${filaInicio}:J${ultimaFila}, "Aceptado", L${filaInicio}:L${ultimaFila})`);
    rangoTotalCobrado.setFormula(`=SUM('Cobros ${hojaMes.getName()}'!H:H)`); // Ajustar si es necesario
    rangoPtoMedio.setFormula(`=IF(COUNTA(K${filaInicio}:K${ultimaFila})>0, C4/COUNTA(K${filaInicio}:K${ultimaFila}), 0)`);
    rangoPacientesPresupuestados.setFormula(`=COUNTIF(J${filaInicio}:J${ultimaFila}, "<>No aceptado")`);
    rangoPacientesAceptados.setFormula(`=COUNTIF(J${filaInicio}:J${ultimaFila}, "Aceptado")`);

    [rangoTotalPresupuestado, rangoTotalAceptado, rangoTotalCobrado, rangoPtoMedio].forEach(celda => {
        celda.setNumberFormat("€#,##0.00");
    });

    hojaMes.autoResizeColumns(2, 4);
}

/// detectar cambios en las hojas y ejecutar tareas
function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;
  if (hoja.getName() === "Vista Previsiones") {
        if ((rango.getRow() === 2 || rango.getRow() === 3) && rango.getColumn() === 2) {
            actualizarVistaPrevisiones();
        } else if (rango.getRow() >= 6) {
            // Si se edita un dato en las filas de datos, actualizar en Staging
            actualizarDatoEnStaging(hoja, rango);
        }
    }


    if (rango.getColumn() === 10 && !hoja.getName().startsWith("Cobros")) {
        var estadoNuevo = rango.getValue();
        var fila = rango.getRow();
        if (fila < 18) return;

        var estadoAnterior = e.oldValue;

        if ((estadoAnterior === "Pendiente con cita" || estadoAnterior === "No aceptado" || estadoAnterior === "Pendiente sin cita") && estadoNuevo === "Aceptado") {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var hojaPrevisiones = ss.getSheetByName("Staging Previsiones") || crearHojaPrevisiones(ss);

                        // Asegurar que existe Vista Previsiones
            var hojaVista = ss.getSheetByName("Vista Previsiones");
            if (!hojaVista) {
                hojaVista = crearVistaPrevisiones(ss);
            }

            // Obtener todos los datos necesarios incluyendo el ID
            var transactionId = hoja.getRange(fila, 1).getValue(); // ID está en la primera columna
            var paciente = hoja.getRange(fila, 3).getValue();
            var fecha = new Date();
            var doctor = hoja.getRange(fila, 5).getValue();
            var importe = hoja.getRange(fila, 12).getValue(); 

            // Agregar a Cobros y Previsiones con el ID
            agregarAStagingPrevisiones(hojaPrevisiones, transactionId, fecha, paciente, doctor, importe);
          
        }

        actualizarFormatoFila(hoja, fila, estadoNuevo);
    }
}


// limpiar formulario de registro de transacción
function limpiarFormulario(hoja) {
    var celdas = ["C3", "C5", "C7","B12", "C12", "D12", "E12", "F12", "G12", "B17", "C17", "D17", "E17", "F17", "B21"];
    celdas.forEach(celda => hoja.getRange(celda).setValue(""));
}
 
///////////////////////////crear botón en menú para cobro
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Cobros')
        .addItem('Ejecutar cobro en fila seleccionada', 'obtenerFilaActiva')
        .addToUi();
}

////////////////////////EJECUTAR COBROS 
function crearStagingCobros(ss) {
    var hojaCobros = ss.insertSheet("Staging Cobros");
    
    // Tabla principal de staging de cobros
    var encabezados = ["ID TRANSACCIÓN", "FECHA DE COBRO", "PACIENTE", "DOCTOR", "TIPO DE PAGO", "MONTO"];

    hojaCobros.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaCobros.autoResizeColumns(1, hojaCobros.getLastColumn());
    hojaCobros.hideSheet();
    return hojaCobros;
}

function obtenerFechaActual() {
  var fecha = new Date();
  var fechaFormateada = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  Logger.log(fechaFormateada);
  return fechaFormateada;
}

//vista previsiones
function obtenerFilaActiva() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var fila = hoja.getActiveCell().getRow(); // Obtiene el número de fila activa
  var datosFila = hoja.getRange(fila, 1, 1, 10).getValues(); // Obtiene los datos de la fila activa
  var abono = datosFila[0][5];
  var saldoPendiente = datosFila[0][6];

  if (saldoPendiente > 0 && abono === "") {
    Browser.msgBox("Error: El campo 'Abono' es obligatorio.");
    return;
  }/////////////////////////////////////////////////////////////////////////////////////////LEO
  if (saldoPendiente > 0 && datosFila[0][8] === "") {
    Browser.msgBox("Error: El campo 'Proxima Fecha' es obligatorio.");
    return;
  }
  if (datosFila[0][7] === "") {
  Browser.msgBox("Error: El campo 'Tipo de pago' es obligatorio.");
  return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (abono) {
      
      var hojaStagingCobros = ss.getSheetByName("Staging Cobros") || crearStagingCobros(ss);

      var newData = [
        datosFila[0][0],
        obtenerFechaActual(),
        datosFila[0][2],
        datosFila[0][3],
        datosFila[0][7],
        datosFila[0][5]
      ];

      hojaStagingCobros.appendRow(newData);

  if (saldoPendiente > 0) {
    var newData = [
      datosFila[0][0],
      datosFila[0][8],
      datosFila[0][2],
      datosFila[0][3],
      datosFila[0][4],
      "",
      datosFila[0][6],
      datosFila[0][7],
      "",
      datosFila[0][10]
    ];
    var hojaStagingPrevisiones = ss.getSheetByName("Staging Previsiones");
    hojaStagingPrevisiones.appendRow(newData);
        // Actualizar el dropdown de años después de agregar el nuevo registro
    actualizarDropdownAnos();
  }
      var ui = SpreadsheetApp.getUi();
    ui.alert('¡Operación exitosa!', 'El cobro se ha registrado apropiadamente', ui.ButtonSet.OK);
}
}

//////////////BALANCE GENERAL/////////////////////
    function obtenerHojasPorMesYAnio(anio) {
      // Lista de meses válidos
      const mesesValidos = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
      ];
      
      // Obtener todas las hojas
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var todasLasHojas = ss.getSheets();
    
      // Filtrar las hojas que cumplan el patrón
      var hojasValidas = todasLasHojas.filter(hoja => {
        var nombreHoja = hoja.getName();
    
        // Patrón esperado: "Mes de Año"
        var patronMes = new RegExp(`^(${mesesValidos.join('|')}) de ${anio}$`);
        var coincide = patronMes.test(nombreHoja);
    
        return coincide;
      });
      
      // Ordenar las hojas según el orden de los meses
      hojasValidas.sort((a, b) => {
        // Extraer solo el nombre del mes (todo antes de " de ")
        var mesA = a.getName().split(" de ")[0];
        var mesB = b.getName().split(" de ")[0];
        Logger.log("Comparando: " + mesA + " con " + mesB);
        return mesesValidos.indexOf(mesA) - mesesValidos.indexOf(mesB);
      });
    
      hojasValidas.forEach(hoja => {
        Logger.log(hoja.getName());
      });
      
      return hojasValidas;
    }
  
function obtenerSumasHojas(hojas) {
  var resultados = Object.create(null);
  
  // Verifica el mapeo de filas
  const mesAFila = {
    "Enero": 5,
    "Febrero": 6,
    "Marzo": 7,
    "Abril": 8,
    "Mayo": 9,
    "Junio": 10,
    "Julio": 11,
    "Agosto": 12,
    "Septiembre": 13,
    "Octubre": 14,
    "Noviembre": 15,
    "Diciembre": 16
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBalance = ss.getSheetByName('BALANCE GENERAL');
  hojaBalance.getRange("B5:H20").clearContent();

  for (let hoja of hojas) {
    try {
      var nombreHoja = hoja.getName();
      var mes = nombreHoja.split(" de ")[0];
      Logger.log('Procesando: ' + nombreHoja + ' (mes: ' + mes + ')');

      var suma = 0;
      var suma_pre = 0;
      var pac_acep = 0;

      var lastRow = hoja.getLastRow(); // Última fila con datos en la hoja
      var startRow = 18; // Primera fila de interés
      if (lastRow >= startRow) {
        var valores = hoja.getRange(startRow, 12, lastRow - startRow + 1, 1).getValues();
        var pacientes = hoja.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();
        var presupuestos = hoja.getRange(startRow, 11, lastRow - startRow + 1, 1).getValues();
        var aceptados = hoja.getRange(startRow, 10, lastRow - startRow + 1, 1).getValues();

        var n_pacientes = pacientes.filter(String).length;
        var n_presupuesto = presupuestos.filter(String).length;
      } else {
        var valores = []; // Si no hay datos, asignamos un array vacío
      }

      var abonoMes = sumarAbonosPorMes(mes, hoja);

      for (let aceptado of aceptados) {
        if (aceptado[0] && aceptado[0] === 'Aceptado') {
          pac_acep += 1;
        }
      }

      for (let presupuesto of presupuestos) {
        if (presupuesto[0] && typeof presupuesto[0] === 'number') {
          suma_pre += presupuesto[0];
        }
      }

      for (let valor of valores) {
        if (valor[0] && typeof valor[0] === 'number') {
          suma += valor[0];
        }
      }

      resultados[nombreHoja] = Number(suma);
      
      var filaDestino = mesAFila[mes];
      Logger.log(`Mes: ${mes}, Fila destino: ${filaDestino}, Suma: ${suma}`);
      
      if (filaDestino) {
        hojaBalance.getRange(filaDestino, 7).setValue(suma);
        hojaBalance.getRange(filaDestino, 3).setValue(n_pacientes);
        hojaBalance.getRange(filaDestino, 5).setValue(n_presupuesto);
        hojaBalance.getRange(filaDestino, 4).setValue(suma_pre);
        hojaBalance.getRange(filaDestino, 6).setValue(suma_pre/n_presupuesto);
        hojaBalance.getRange(filaDestino, 8).setValue(pac_acep);
        hojaBalance.getRange(filaDestino, 2).setValue(abonoMes);
        Logger.log(`Valor ${suma} escrito en fila ${filaDestino} para ${mes}`);
      } else {
        Logger.log(`No se encontró fila destino para el mes: ${mes}`);
      }

    } catch (error) {
      Logger.log(`Error en ${nombreHoja}: ${error}`);
      resultados[nombreHoja] = 0;
    }
  }

  return resultados;
}

function sumarAbonosPorMes(nombreMes) {

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staging Cobros");
  if (!hoja) {
    return "La hoja 'Staging Cobros' no existe.";
  }

  var datos = hoja.getDataRange().getValues();
  
  // Mapeo de nombres de meses a números (Enero = 1, Febrero = 2, etc.)
  var meses = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
    "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
  };

  var mesBuscado = meses[nombreMes]; // Obtener el número del mes

  if (!mesBuscado) {
    return "Mes no válido. Usa nombres como 'Enero', 'Febrero', etc.";
  }

  var suma = 0;

  for (var i = 1; i < datos.length; i++) { // Empezamos en 1 para ignorar la cabecera
    var fecha = new Date(datos[i][1]); // FECHA DE COBRO está en la columna B (índice 1)
    var monto = parseFloat(datos[i][5]); // MONTO está en la columna F (índice 5)

    if (!isNaN(fecha.getTime()) && fecha.getMonth() + 1 === mesBuscado) { 
      suma += isNaN(monto) ? 0 : monto;
    }
  }

  return suma;
}

function obtenerAniosDeHojas() {
  // Lista de meses válidos
  const mesesValidos = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];

  // Expresión regular para buscar hojas con el patrón "Mes de Año"
  const patron = new RegExp(`^(${mesesValidos.join('|')}) de (\\d{4})$`);

  // Obtener todas las hojas
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var todasLasHojas = ss.getSheets();

  // Set para almacenar años únicos
  var aniosUnicos = new Set();

  // Recorrer las hojas y extraer los años
  todasLasHojas.forEach(hoja => {
    var nombreHoja = hoja.getName();
    var coincidencia = nombreHoja.match(patron);
    if (coincidencia) {
      aniosUnicos.add(parseInt(coincidencia[2])); // Convertir a número
    }
  });

  // Convertir el Set a lista, ordenarla de menor a mayor y devolverla
  var listaAnios = Array.from(aniosUnicos).sort((a, b) => a - b);
  Logger.log(listaAnios);
  return listaAnios;
}

function actualizarFiltroDeAnios() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BALANCE GENERAL");
  if (!hoja) {
    Logger.log("La hoja 'BALANCE GENERAL' no existe.");
    return;
  }

  // Obtener la lista de años únicos ordenados
  var anios = obtenerAniosDeHojas(); // Esta es la función que hicimos antes

  if (anios.length === 0) {
    Logger.log("No hay años para agregar al filtro.");
    return;
  }

  // Crear la validación de datos (menú desplegable)
  var rango = hoja.getRange("A2");
  var reglaValidacion = SpreadsheetApp.newDataValidation()
    .requireValueInList(anios, true)
    .setAllowInvalid(false) // No permite valores fuera de la lista
    .build();

  // Aplicar la validación a la celda A2
  rango.setDataValidation(reglaValidacion);
  Logger.log("Filtro de años actualizado con: " + anios);
}



function aplicarFiltroPorAnio(anio) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BALANCE GENERAL");
  Logger.log("Aplicando filtro con el año: " + anio);
  
}

    
function balanceGeneral(annio) {
  //var anios = obtenerAniosDeHojas()
  var listaHojas = obtenerHojasPorMesYAnio(annio)
  var sumas = obtenerSumasHojas(listaHojas)
}