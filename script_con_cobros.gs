// Función para mantener el contador secuencial
function getNextTransactionNumber() {
  const properties = PropertiesService.getScriptProperties();
  let currentNumber = parseInt(properties.getProperty('LAST_TRANSACTION_NUMBER') || '0');
  currentNumber++;
  properties.setProperty('LAST_TRANSACTION_NUMBER', currentNumber.toString());
  return currentNumber.toString().padStart(6, '0');
}

// Generador de ID secuencial
function generateSequentialTransactionId() {
  const year = new Date().getFullYear().toString().slice(-2);
  const month = (new Date().getMonth() + 1).toString().padStart(2, '0');
  const sequence = getNextTransactionNumber();
  return `${year}${month}${sequence}`;
}

////// Registro de pacientes
function guardarDatosEnTabla2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFormulario = ss.getSheetByName("Registro de presupuesto");

    if (!hojaFormulario) {
        //Logger.log("Error: No se encontró la hoja 'Registro de presupuesto'");
        return;
    }

    var datos = hojaFormulario.getRange("B3:H21").getValues();

    if (!datos[0][1]) {
        Browser.msgBox("Error: El campo 'Paciente' es obligatorio.");
        return;
    }

    var fechaIngresada = datos[2][1];
    if (!fechaIngresada) {
        Browser.msgBox("Error: Debes ingresar una fecha.");
        return;
    }
    if (datos[14][2] === "Aceptado" && !datos[14][3]) {
    Browser.msgBox("Error: Para pacientes con estado 'Aceptado', la fecha de inicio es obligatoria.");
    return;
    }

    // Generar ID de transacción
    const transactionId = generateSequentialTransactionId();

    var fecha = new Date(fechaIngresada);
    var nombreMes = fecha.toLocaleDateString("es-ES", { year: "numeric", month: "long" });
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    // Crear o obtener la hoja staging de cobros
    var hojaMes = ss.getSheetByName(nombreMes) || crearHojaMes(ss, nombreMes);
    var hojaPrevisiones = ss.getSheetByName("Staging Previsiones");
    if (!hojaPrevisiones) {
        hojaPrevisiones = crearHojaPrevisiones(ss);
    }
    
    // Crear o obtener la hoja Vista Previsiones
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    if (!hojaVista) {
        hojaVista = crearVistaPrevisiones(ss);
    }
    // Crear o obtener la hoja Vista Cobros
    var hojaVistaCobros = ss.getSheetByName("Vista Cobros");
    if (!hojaVistaCobros) {
    hojaVistaCobros = crearVistaCobros(ss);
    }
    // Crear o obtener la hoja staging de cobros
    var hojaStagingCobros = ss.getSheetByName("Staging Cobros") || crearStagingCobros(ss);
      if (!hojaStagingCobros) {
    hojaStagingCobros = crearVistaCobros(ss);
    }

    var filaEscribir = hojaMes.getLastRow() < 17 ? 18 : hojaMes.getLastRow() + 1;

    var nuevaFila = [
        transactionId,     // ID Transacción
        fechaIngresada,    // FECHA DE PRESUPUESTO
        datos[0][1],       // PACIENTE
        datos[4][1],       // TELÉFONO
        datos[9][0],       // DOCTOR/A
        datos[9][1],       // ATP
        datos[9][2],       // TIPOLOGÍA PV
        datos[9][3],       // SUBTIPOLOGÍA
        datos[9][4],       // PLAN DE CITAS
        datos[14][2],      // ESTADO
        datos[14][0],      // IMPORTE PRESUPUESTADO
        datos[14][1],      // IMPORTE ACEPTADO
        datos[14][3],      // FECHA DE INICIO
        datos[18][0]       // OBSERVACIONES
    ];
    
    hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    actualizarFormatoFila(hojaMes, filaEscribir, datos[14][2]);
    hojaMes.getRange(filaEscribir, 11).setNumberFormat("€#,##0.00");
    hojaMes.getRange(filaEscribir, 12).setNumberFormat("€#,##0.00");

    if (datos[14][2] === "Aceptado") {
        agregarAStagingPrevisiones(hojaPrevisiones, transactionId, datos[14][3], datos[0][1], datos[9][0], datos[14][1]);
    }
    actualizarTablaResumen(hojaMes);
    limpiarFormulario(hojaFormulario);

     //actualiza filtro de años en la hoja de Balance General
    actualizarFiltroDeAnios();

   // Logger.log("Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");
}
// Crear hojaMes (parrilla ppal)
function crearHojaMes(ss, nombreMes) {
    var hojaMes = ss.insertSheet(nombreMes);
    var encabezados = [
        "ID TRANSACCIÓN", "FECHA PRESUPUESTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", 
        "ATP", "TIPOLOGÍA PV", "SUBTIPOLOGÍA", "PLAN DE CITAS", "ESTADO", 
        "IMPORTE PRESUPUESTADO", "IMPORTE ACEPTADO", "FECHA INICIO / CONCRETAR", "OBSERVACIONES"
    ];

    hojaMes.getRange(17, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center");
    hojaMes.getRange(17, 1, 1, encabezados.length).createFilter();
    hojaMes.autoResizeColumns(1, hojaMes.getLastColumn());
    return hojaMes;
}

/// PREVISIONES 

/// Crear hoja STAGING PREVISIONES 
function crearHojaPrevisiones(ss) {
    var hojaPrevisiones = ss.insertSheet("Staging Previsiones");
    var encabezados = [
        "ID TRANSACCIÓN", 
        "FECHA ACTUAL", 
        "PACIENTE", 
        "DOCTOR", 
        "PREV TOTAL", 
        "PREV ESPERADA", 
        "PREV PAGADA", 
        "SALDO PENDIENTE", 
        "TIPO DE PAGO", 
        "CITA", 
        "TRATAMIENTO",
        "ESTADO / € TOTALES"
    ];

    hojaPrevisiones.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaPrevisiones.autoResizeColumns(1, hojaPrevisiones.getLastColumn());
    hojaPrevisiones.hideSheet();
    return hojaPrevisiones;
}
/// Agregar a paciente a Staging previsiones
function agregarAStagingPrevisiones(hojaPrevisiones, transactionId, fechaInicio, paciente, doctor, importeAceptado) {
    //Verificar si el ID ya existe
    if (existeIdEnHoja(hojaPrevisiones, transactionId)) {
      //  Logger.log("ID ya existe en Staging Previsiones: " + transactionId);
        Browser.msgBox("Error", `El ID de transacción ${id} ya existe en la hoja ${hoja.getName()}`, Browser.Buttons.OK);
        return;
    }

    var ultimaFila = hojaPrevisiones.getLastRow() + 1;
    var tipoPagoOpciones = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    // Determinar el estado de pago
    var estadoPago = importeAceptado === 0 ? "PAGADO" : "PENDIENTE";

    var nuevaFila = [
        transactionId,
        fechaInicio,
        paciente,
        doctor,
        importeAceptado,
        "", 
        "",
        "", 
        "",
        "",
        "",
        estadoPago
    ];

    hojaPrevisiones.getRange(ultimaFila, 1, 1, nuevaFila.length).setValues([nuevaFila]);
    
   // Aplicar formatos
    hojaPrevisiones.getRange(ultimaFila, 5).setNumberFormat("€#,##0.00");
    hojaPrevisiones.getRange(ultimaFila, 6).setNumberFormat("€#,##0.00"); // Formato para PREV ESPERADA
    hojaPrevisiones.getRange(ultimaFila, 7).setNumberFormat("€#,##0.00"); // Ahora PREV PAGADA
    hojaPrevisiones.getRange(ultimaFila, 8).setNumberFormat("€#,##0.00"); // Ahora SALDO PENDIENTE
    
    hojaPrevisiones.getRange(ultimaFila, 9).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(tipoPagoOpciones, true).setAllowInvalid(false).build()
    );
    hojaPrevisiones.getRange(ultimaFila, 10).setDataValidation(
        SpreadsheetApp.newDataValidation().requireDate().build()
    );
    actualizarDropdownAnos();
}

/// Actualizar datos en el Staging de previsiones
function actualizarDatoEnStaging(hojaVista, rango) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    
    var fila = rango.getRow();
    var columna = rango.getColumn();
    var idTransaccion = hojaVista.getRange(fila, 1).getValue();
    var nuevoValor = rango.getValue();
    
    // Buscar la fila correspondiente en Staging
    var datosStaging = hojaStaging.getDataRange().getValues();
    

  for (var i = 1; i < datosStaging.length; i++) {
        if (datosStaging[i][0] === idTransaccion) {
            hojaStaging.getRange(i + 1, columna).setValue(nuevoValor);
          
        }
    }

}

// crear hoja vista de previsiones
function crearVistaPrevisiones(ss) {
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    if (!hojaVista) {
        hojaVista = ss.insertSheet("Vista Previsiones");
        configurarVistaPrevisiones(hojaVista, ss);
    }
    return hojaVista;
}
// Configurar hoja de vista de previsiones
function configurarVistaPrevisiones(hojaVista, ss) {

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

////INSTRUCCIONES
    hojaVista.getRange(1, 5, 1, 6).merge().setValue("Actualizar Previsión: Cuándo? al registrar un presupuesto NUEVO como 'ACEPTADO' o cuando cambia de ESTADO a 'ACEPTADO' en la parrilla mensual. Qué hacer? modifique las columnas PREV ESPERADA y CITA, TRATAMIENTO y TIPO DE PAGO. Paso final: seleccione toda la fila, vaya a MENÚ > PREVISIONES > ACTUALIZAR PREVISIÓN.")
        .setFontSize(9)
        .setBackground("#00c896")
        .setFontColor("#424242")
        .setWrap(true)
        .setVerticalAlignment("middle");
      hojaVista.getRange(2, 5, 1, 6).merge().setValue("Agregar Previsión: Cuándo? para crear varias citas/pagos para el mismo cliente / presupuesto. Qué hacer? copie y pegue TODA la fila deseada, modifique los datos en las columnas PREV ESPERADA y CITA, Tratamiento (opcional). Paso final: seleccione toda la fila modificada, vaya a MENÚ > PREVISIONES > AGREGAR PREVISIÓN.")
      .setFontSize(9)
      .setBackground("#424242")
      .setFontColor("#FFFFFF")
      .setWrap(true)
      .setVerticalAlignment("middle");

      hojaVista.getRange(3, 5, 1, 6).merge().setValue("Cobros: Cuándo? para registrar pago de una previsión. Qué hacer? ingrese en PREV PAGADA el monto pagado. Paso final: seleccione toda la fila modificada, en el MENÚ > COBROS > EJECUTAR COBRO.")
      .setFontSize(9)
      .setBackground("#98e0fa")
      .setFontColor("#080808")
      .setWrap(true)
      .setVerticalAlignment("middle");


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
            //Logger.log("Fecha inválida encontrada en la fila " + (i + 1));
        }
    }
}

// Asegurarsede que el Set se convierte correctamente a array
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

/// columnas de control de fecha
hojaVista.getRange("Q1").setValue("Fechas de control");
hojaVista.getRange("Q2:Q3").setValues([["Fecha inicio"], ["Fecha fin"]]);
hojaVista.getRange("Q2").setFormula('=IF(B3="Ver todo el año",DATE(B2,1,1),DATE(B2,MATCH(B3,{"Enero";"Febrero";"Marzo";"Abril";"Mayo";"Junio";"Julio";"Agosto";"Septiembre";"Octubre";"Noviembre";"Diciembre"},0),1))');
hojaVista.getRange("Q3").setFormula('=IF(B3="Ver todo el año",DATE(B2,12,31),EOMONTH(Q2,0))');

// Ocultar las columnas de control
hojaVista.hideColumns(17, 1); // Oculta la columna de control


// Encabezados de la tabla
var encabezados = [
    "ID TRANSACCIÓN", "FECHA", "PACIENTE", "DOCTOR", "PREV TOTAL",
    "PREV ESPERADA", 
    "PREV PAGADA",
    "SALDO PENDIENTE", "TIPO DE PAGO", "CITA", 
    "TRATAMIENTO", "ESTADO / € TOTALES" 
];

hojaVista.getRange(5, 1, 1, encabezados.length).setValues([encabezados])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("white")
    .setHorizontalAlignment("center");
    // Agregar filtros a los encabezados
    hojaVista.getRange(5, 1, 1, encabezados.length).createFilter();
    

// Configurar formato de columnas
hojaVista.getRange("E:H").setNumberFormat("€#,##0.00");

// Validación para tipo de pago ///////////////////////////////////////////////////////////////////////////////////////// variables sin usar
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


    // Añadir tabla resumen 
    configurarTablaResumen(hojaVista);

/// Ajustar ancho de columnas de tabla y sección de instrucciones
   for (var i = 5; i <= 10; i++) {
      hojaVista.setColumnWidth(i, 150); // Ancho fijo para columnas de instrucciones
    }
    
    //  autoResize solo para las demás columnas de la tabla
    hojaVista.autoResizeColumns(1, 4); // Columnas A-D
    hojaVista.autoResizeColumn(12);    // Columna K
    
    // margen adicional a cada columna excepto las de instrucciones
    for (var i = 1; i <= 12; i++) {
      // Saltamos las columnas E a J (5 a 10)
      if (i < 5 || i > 10) {
        var anchoActual = hojaVista.getColumnWidth(i);
        var margenExtra = 25; // Píxeles 
        hojaVista.setColumnWidth(i, anchoActual + margenExtra);
      }
    }
    
}

///// Actualizar filtro de fechas en Vista Previsiones
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
///// Tabla resumen de Vista Previsiones
function configurarTablaResumen(hojaVista) {
    var resumenEncabezados = [
        ["RESUMEN", "", ""],
        ["Total Importe", "=SUM(E6:E)", ""],
        ["Previsión mes actual", "=SUM(F6:F)", ""], // Nueva línea para PREV ESPERADA
        ["Previsión abonada", "=SUM(G6:G)", ""] // Ahora referencia a columna G
    ];
 
    var rangoResumen = hojaVista.getRange(1, 14, resumenEncabezados.length, 3);
    rangoResumen.setValues(resumenEncabezados);
 
    // Aplicar formatos
    hojaVista.getRange("N1:O1")
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold")
        .merge();
    
    hojaVista.getRange("O2:O4").setNumberFormat("€#,##0.00"); // Ahora 5 filas para incluir PREV ESPERADA
    
    // Estilos alternados
    hojaVista.getRange("N2:O2").setBackground("#f6f6f6");
    hojaVista.getRange("N3:O3").setBackground("#e2e2e2");
    hojaVista.getRange("N4:O4").setBackground("#f6f6f6");
    
    // Bordes
    hojaVista.getRange("N1:O4").setBorder(true, true, true, true, true, true); 
    hojaVista.autoResizeColumns(14, 3);

}
///// Actualizar Vista de Previsiones luego de filtrar por fecha
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
        // Si hay mes seleccionado y no es "Ver todo el año", usar las fechas calculadas en Q2 y Q3
        fechaInicio = hojaVista.getRange("Q2").getValue();
        fechaFin = hojaVista.getRange("Q3").getValue();
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
       hojaVista.getRange(6, 9, datosFiltrados.length).setDataValidation(validacionTipoPago)  // Tipo de Pago
            .setBackground("#f0f0f0");  // Light gray background
        hojaVista.getRange(6, 10, datosFiltrados.length).setDataValidation(validacionFecha)     // Próximo Pago O CITA
            .setBackground("#f0f0f0");  // Light gray background
        hojaVista.getRange(6, 7, datosFiltrados.length).setBackground("#98e0fa");  // Abono column O PREV ABONADA
        hojaVista.getRange(6, 6, datosFiltrados.length).setBackground("#f0f0f0");  //  PREV ESPERADA

        hojaVista.getRange(6, 11, datosFiltrados.length).setBackground("#f0f0f0");  // TRATAMIENTO column

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
        hojaVista.getRange(6, 9, datosFiltrados.length).setDataValidation(validacionTipoPago);  // Tipo de Pago
        hojaVista.getRange(6, 10, datosFiltrados.length).setDataValidation(validacionFecha);     // Próximo Pago O CITA
    } else {
        hojaVista.getRange(6, 1).setValue("No hay datos para mostrar");
    }
}

// Función helper para verificar si ya existe el ID en una hoja
function existeIdEnHoja(hoja, id) {
    if (!id) return false;
    var datos = hoja.getRange("A:A").getValues();
    var duplicado = datos.some(row => row[0] === id);
    
    if (duplicado && hojaActiva.getName() !== "Vista Previsiones") {
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
    var rangoPacientesCobrados = hojaMes.getRange(6, 4);

    // Extraer el nombre del mes y año de la hoja
    var nombreHoja_cobros = hojaMes.getName();
    var nombreMes_cobros = nombreHoja_cobros.split(" de ")[0];
    var anio_cobros = nombreHoja_cobros.includes(" de ") ? nombreHoja_cobros.split(" de ")[1] : new Date().getFullYear().toString();

    // Mapeo de nombres de meses a números
    var mesesMap_cobros = {
        "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, 
        "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, 
        "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
    };

    var mesNum_cobros = mesesMap_cobros[nombreMes_cobros];

    // Crear las fechas de inicio y fin del mes
    var fechaInicio_cobros = `DATE(${anio_cobros},${mesNum_cobros},1)`;
    var fechaFin_cobros = `EOMONTH(DATE(${anio_cobros},${mesNum_cobros},1),0)`;

// Fórmulas de tabla resumen

    rangoTotalPresupuestado.setFormula(`=SUMIF(J${filaInicio}:J${ultimaFila}, "<>No aceptado", K${filaInicio}:K${ultimaFila})`);
    rangoTotalAceptado.setFormula(`=SUMIF(J${filaInicio}:J${ultimaFila}, "Aceptado", L${filaInicio}:L${ultimaFila})`);
    rangoPtoMedio.setFormula(`=IF(COUNTA(K${filaInicio}:K${ultimaFila})>0, C4/COUNTA(K${filaInicio}:K${ultimaFila}), 0)`);
    rangoPacientesPresupuestados.setFormula(`=COUNTIF(J${filaInicio}:J${ultimaFila}, "<>No aceptado")`);
    rangoPacientesAceptados.setFormula(`=COUNTIF(J${filaInicio}:J${ultimaFila}, "Aceptado")`);
    /// cuenta lo que se cobre en ese mes, si se cobra luego de ese mes, no se cuenta aquí 

    rangoPacientesCobrados.setFormula(`=COUNTIFS('Staging Cobros'!B:B,">="&${fechaInicio_cobros},'Staging Cobros'!B:B,"<="&${fechaFin_cobros})`);
    rangoTotalCobrado.setFormula(`=SUMIFS('Staging Cobros'!F:F,'Staging Cobros'!B:B,">="&${fechaInicio_cobros},'Staging Cobros'!B:B,"<="&${fechaFin_cobros})`);

    ///// Cuenta lo que se ha pagado que corresponde con los presupuestos de ese mes 
    // rangoPacientesCobrados.setFormula(`=COUNTIFS('Staging Cobros'!A:A,TRANSPOSE(UNIQUE(A${filaInicio}:A${ultimaFila})))`);
    // rangoTotalCobrado.setFormula(`=SUMPRODUCT(SUMIF('Staging Cobros'!A:A,A${filaInicio}:A${ultimaFila},'Staging Cobros'!F:F))`);


    [rangoTotalPresupuestado, rangoTotalAceptado, rangoTotalCobrado, rangoPtoMedio].forEach(celda => {
        celda.setNumberFormat("€#,##0.00");
    });

    hojaMes.autoResizeColumns(2, 4);
}

/// detectar cambios en las hojas y ejecutar tareas
function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;


  // Detectar cambios en la hoja "BALANCE GENERAL"
    if (hoja.getName() === "BALANCE GENERAL") {
        if (rango.getA1Notation() === "A2") { // Solo si editan A2
            var anioSeleccionado = e.value; // Captura el nuevo valor
            
            if (anioSeleccionado) {
                var anioEntero = parseInt(anioSeleccionado, 10); // Convertir a entero
                if (!isNaN(anioEntero)) { // Verificar que sea un número válido
                   // Logger.log("Año seleccionado en BALANCE GENERAL: " + anioEntero);
                    balanceGeneral(anioEntero);
                }
            }
        }

        /////////////////
        // Nuevo código para detectar cambios en A26
        if (rango.getA1Notation() === "A26") {
            var anioComparacion = e.value;
            
            if (anioComparacion) {
                var anioEnteroComp = parseInt(anioComparacion, 10);
                if (!isNaN(anioEnteroComp)) {
                    // Logger.log("Año seleccionado para comparación: " + anioEnteroComp);
                    balanceGeneralComparacion(anioEnteroComp);
                }
            }
        }
        ////////////////
  }
// Detectar cambios en la hoja en VISTA PREVISIONES
  if (hoja.getName() === "Vista Previsiones") {
        if ((rango.getRow() === 2 || rango.getRow() === 3) && rango.getColumn() === 2) {
            actualizarVistaPrevisiones();
        } 

    }

    ////////VISTA COBROS
    if (hoja.getName() === "Vista Cobros") {
        if ((rango.getRow() === 2 || rango.getRow() === 3) && rango.getColumn() === 2) {
            actualizarVistaCobros();
        }
    }


    if (rango.getColumn() === 10 && !hoja.getName().startsWith("Cobros")&& hoja.getName() !== "Vista Previsiones") {
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

/////// COBROS

//crear botón en menú para cobro
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Crear un nuevo menú para ACTUALIZAR y CREAR previsiones
    ui.createMenu('Previsiones')
        .addItem('Actualizar Previsión', 'actualizarPrevisionManual')
        .addItem('Agregar Previsión', 'agregarPrevisionManual')
        .addToUi();

    ui.createMenu('Cobros')
        .addItem('Ejecutar cobro', 'obtenerFilaActiva')
        .addToUi();

      // Agregar nuevo menú para el Balance General
  ui.createMenu('Análisis')
    .addItem('Actualizar Balance General', 'actualizarBalanceGeneral')
    .addItem('Actualizar KPIs', 'actualizarKPIs') // Nueva opción
    .addToUi(); 

    // Añadir esta línea para inicializar Vista Cobros
    actualizarDropdownAnosCobros();
}
///agregar previsión desde vista previsiones (crear varias citas)
function agregarPrevisionManual() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaVista = ss.getActiveSheet();
    
    // Verificar que estamos en la hoja correcta
    if (hojaVista.getName() !== "Vista Previsiones") {
        Browser.msgBox("Error", "Por favor, seleccione una fila en la hoja 'Vista Previsiones'", Browser.Buttons.OK);
        return;
    }
    
    var fila = hojaVista.getActiveCell().getRow();
    
    // Verificar que la fila seleccionada está en la zona de datos (después de la fila 5)
    if (fila <= 5) {
        Browser.msgBox("Error", "Por favor, seleccione una fila de datos (después de la fila 5)", Browser.Buttons.OK);
        return;
    }
    
    // Obtener los datos de la fila seleccionada
    var datosFila = hojaVista.getRange(fila, 1, 1, 12).getValues()[0];
    
    // Verificar que los datos mínimos necesarios estén presentes
    if (!datosFila[0] || !datosFila[2] || !datosFila[3] || !datosFila[4] || !datosFila[5] || !datosFila[9]) {
        Browser.msgBox("Error", "Datos incompletos", Browser.Buttons.OK);
        return;
    }
    
    // Verificar si el ID ya existe en la hoja Staging Previsiones
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    if (!hojaStaging) {
        hojaStaging = crearHojaPrevisiones(ss);
    }
    
   
   // Extraer los datos necesarios para llamar a agregarAStagingPrevisiones
    var transactionId = datosFila[0]; // ID Transacción
    var cita = datosFila[9]
    var fechaInicio = cita || new Date(); // Fecha, usar la fecha actual si está vacía
    var paciente = datosFila[2]; // Paciente
    var doctor = datosFila[3]; // Doctor
    var importeAceptado = datosFila[4]; // Importe Total
    var prevEsperada = datosFila[5];
    var prevPagada = datosFila[6];
    var saldoPendiente = datosFila[7];
    var tipo_pago = datosFila[8];
    var treatment = datosFila[10];

    var ultimaFila = hojaStaging.getLastRow() + 1;
    var tipoPagoOpciones = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    // Determinar el estado de pago
    var estadoPago = saldoPendiente === 0 ? "PAGADO" : "PENDIENTE";

    var nuevaFila = [
        transactionId,
        fechaInicio,
        paciente,
        doctor,
        importeAceptado,
        prevEsperada, 
        prevPagada,
        saldoPendiente, 
        tipo_pago,
        cita,
        treatment,
        estadoPago
    ];


    hojaStaging.getRange(ultimaFila, 1, 1, nuevaFila.length).setValues([nuevaFila]);
    
    // Aplicar formatos
    hojaStaging.getRange(ultimaFila, 5).setNumberFormat("€#,##0.00");
    hojaStaging.getRange(ultimaFila, 6).setNumberFormat("€#,##0.00"); // Formato para PREV ESPERADA
    hojaStaging.getRange(ultimaFila, 7).setNumberFormat("€#,##0.00"); // Ahora PREV PAGADA
    hojaStaging.getRange(ultimaFila, 8).setNumberFormat("€#,##0.00"); // Ahora SALDO PENDIENTE
    
    hojaStaging.getRange(ultimaFila, 9).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(tipoPagoOpciones, true).setAllowInvalid(false).build()
    );
    hojaStaging.getRange(ultimaFila, 10).setDataValidation(
        SpreadsheetApp.newDataValidation().requireDate().build()
    );
    actualizarDropdownAnos();
    
    // Actualizar la vista después de añadir el nuevo registro
    actualizarVistaPrevisiones();

    // Mostrar mensaje de éxito
    Browser.msgBox("Éxito", "La Previsión adicional se ha agregado correctamente", Browser.Buttons.OK);
}

////// actualizar previsión (para fijar cita de previsión apropiadamente luego de registrar paciente aceptado)
function actualizarPrevisionManual() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaVista = ss.getActiveSheet();
    
    // Verificar que estamos en la hoja correcta
    if (hojaVista.getName() !== "Vista Previsiones") {
        Browser.msgBox("Error", "Por favor, seleccione una fila en la hoja 'Vista Previsiones'", Browser.Buttons.OK);
        return;
    }
    
    var fila = hojaVista.getActiveCell().getRow();
    
    // Verificar que la fila seleccionada está en la zona de datos (después de la fila 5)
    if (fila <= 5) {
        Browser.msgBox("Error", "Por favor, seleccione una fila de datos (después de la fila 5)", Browser.Buttons.OK);
        return;
    }
    
    // Obtener los datos de la fila seleccionada
    var datosFila = hojaVista.getRange(fila, 1, 1, 12).getValues()[0];
    
    // Verificar que los datos mínimos necesarios estén presentes
    if (!datosFila[0] || !datosFila[2] || !datosFila[3] || !datosFila[4] || !datosFila[5] || !datosFila[9]) {
        Browser.msgBox("Error", "Datos incompletos", Browser.Buttons.OK);
        return;
    }
    
    // Verificar si el ID ya existe en la hoja Staging Previsiones
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    if (!hojaStaging) {
        hojaStaging = crearHojaPrevisiones(ss);
    }

    // Datos para identificar el registro específico
    var transactionId = datosFila[0]; // ID Transacción
    var dataStaging = hojaStaging.getDataRange().getValues();
    var filaEncontrada = -1;

 for (var i = 1; i < dataStaging.length; i++) {
        // Primera verificación por ID
        if (dataStaging[i][0] === transactionId) {
            // Verificaciones adicionales para encontrar el registro exacto
            // Comparamos paciente (índice 2) y doctor (índice 3)
            if (dataStaging[i][2] === datosFila[2] && 
                dataStaging[i][3] === datosFila[3] && 
                dataStaging[i][4] === datosFila[4]) {
                
                // Verificación adicional con la fecha (si es consistente entre vistas)
                var fechaStaging = dataStaging[i][1];
                var fechaVista = datosFila[1];
                
                // Compara fechas como strings si son objetos Date
                var fechaStagingStr = fechaStaging instanceof Date ? 
                    Utilities.formatDate(fechaStaging, Session.getScriptTimeZone(), "yyyy-MM-dd") : 
                    String(fechaStaging);
                    
                var fechaVistaStr = fechaVista instanceof Date ? 
                    Utilities.formatDate(fechaVista, Session.getScriptTimeZone(), "yyyy-MM-dd") : 
                    String(fechaVista);
                
                if (fechaStagingStr === fechaVistaStr) {
                    filaEncontrada = i + 1; // +1 porque las filas empiezan en 1, no en 0
                    break;
                }
            }
        }
    }
    
    // Si no se ha encontrado con criterios estrictos, preguntar al usuario si desea actualizar el primer registro con ese ID
    if (filaEncontrada === -1) {
        // Buscar al menos el primer registro con ese ID
        for (var i = 1; i < dataStaging.length; i++) {
            if (dataStaging[i][0] === transactionId) {
                var respuesta = Browser.msgBox("Confirmar actualización", 
                    "No se encontró una coincidencia exacta. ¿Desea actualizar el primer registro con ID: " + 
                    transactionId + " para el paciente " + dataStaging[i][2] + "?", 
                    Browser.Buttons.YES_NO);
                
                if (respuesta === Browser.Buttons.YES) {
                    filaEncontrada = i + 1;
                }
                break;
            }
        }
    }
    
    if (filaEncontrada === -1) {
        Browser.msgBox("Operación cancelada", "No se actualizó ningún registro", Browser.Buttons.OK);
        return;
    }
   
    // Extraer los datos necesarios para llamar a agregarAStagingPrevisiones
    // var transactionId = datosFila[0]; // ID Transacción
    var cita = datosFila[9]
    var fechaInicio = cita || dataStaging[filaEncontrada-1][1]; // Mantener fecha original si no hay cita
    var paciente = datosFila[2]; // Paciente
    var doctor = datosFila[3]; // Doctor
    var importeAceptado = datosFila[4]; // Importe Total
    var prevEsperada = datosFila[5];
    var prevPagada = datosFila[6];
    var saldoPendiente = datosFila[7];
    var tipo_pago = datosFila[8];
    var treatment = datosFila[10];


    var ultimaFila = hojaStaging.getLastRow() + 1;
    var tipoPagoOpciones = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    // Determinar el estado de pago
    var estadoPago = saldoPendiente === 0 ? "PAGADO" : "PENDIENTE";

    // Crear el array con los datos actualizados
    var datosActualizados = [
        transactionId,
        fechaInicio,
        paciente,
        doctor,
        importeAceptado,
        prevEsperada,
        prevPagada,
        saldoPendiente,
        tipo_pago,
        cita,
        treatment,
        estadoPago
    ];
    // hojaStaging.getRange(ultimaFila, 1, 1, nuevaFila.length).setValues([nuevaFila]);
        hojaStaging.getRange(filaEncontrada, 1, 1, datosActualizados.length).setValues([datosActualizados]);

    // Aplicar formatos
    hojaStaging.getRange(ultimaFila, 5).setNumberFormat("€#,##0.00");
    hojaStaging.getRange(ultimaFila, 6).setNumberFormat("€#,##0.00"); // Formato para PREV ESPERADA
    hojaStaging.getRange(ultimaFila, 7).setNumberFormat("€#,##0.00"); // Ahora PREV PAGADA
    hojaStaging.getRange(ultimaFila, 8).setNumberFormat("€#,##0.00"); // Ahora SALDO PENDIENTE
    
    hojaStaging.getRange(ultimaFila, 9).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(tipoPagoOpciones, true).setAllowInvalid(false).build()
    );
    hojaStaging.getRange(ultimaFila, 10).setDataValidation(
        SpreadsheetApp.newDataValidation().requireDate().build()
    );
    actualizarDropdownAnos();
    
    // Actualizar la vista después de añadir el nuevo registro
    actualizarVistaPrevisiones();

    // Mostrar mensaje de éxito
    Browser.msgBox("Éxito", "La Previsión se ha actualizado correctamente", Browser.Buttons.OK);
}

///Crear Staging Cobros
function crearStagingCobros(ss) {
    var hojaCobros = ss.insertSheet("Staging Cobros");
    
    // Tabla principal de staging de cobros
    var encabezados = ["ID TRANSACCIÓN", "FECHA DE COBRO", "PACIENTE", "DOCTOR", "TIPO DE PAGO", "MONTO", "TRATAMIENTO"];

    hojaCobros.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaCobros.autoResizeColumns(1, hojaCobros.getLastColumn());
    hojaCobros.hideSheet();
    return hojaCobros;
}
// validar fecha
function obtenerFechaActual() {
  var fecha = new Date();
  var fechaFormateada = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  //Logger.log(fechaFormateada);
  return fechaFormateada;
}

//contar ocurrencias de la transacción
function contarOcurrenciasID(idBuscado, datos) {
  
  // Encontrar el índice de la columna "ID TRANSACCIÓN"
  const headers = datos[0];
  const idColumnaIndex = headers.findIndex(header => header === "ID TRANSACCIÓN");
  
  // Encontrar el índice de la columna "PREV PAGADA"
  const prevPagadaIndex = headers.findIndex(header => header === "PREV PAGADA");
  
  // Verificar si las columnas existen
  if (idColumnaIndex === -1) {
    throw new Error("No se encontró la columna 'ID TRANSACCIÓN'");
  }
  
  if (prevPagadaIndex === -1) {
    throw new Error("No se encontró la columna 'PREV PAGADA'");
  }
  
  // Filtrar solo las filas que tienen el ID buscado y datos en PREV PAGADA (excluyendo la fila de encabezados)
  const coincidencias = datos.slice(1).filter(fila => {
    return fila[idColumnaIndex] == idBuscado && 
           fila[prevPagadaIndex] !== undefined && 
           fila[prevPagadaIndex] !== null && 
           fila[prevPagadaIndex] !== "";
  });
  
  // Retornar la cantidad de coincidencias
  return coincidencias.length;
}

///// Ejecutar cobro
function obtenerFilaActiva() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var fila = hoja.getActiveCell().getRow(); 
  var datosFila = hoja.getRange(fila, 1, 1, 12).getValues()[0]; 
  var prevAbonada = Number(datosFila[6]); // PREV PAGADA
  var saldoPendAnterior = Number(datosFila[7]);
  
  var importeTotal = Number(datosFila[4]); // PREV TOTAL
  var transactionId = datosFila[0];
  var fechaExcluir = datosFila[1];
  
  // Validar datos
  if (isNaN(prevAbonada) || prevAbonada <= 0) {
    Browser.msgBox("Error: El campo 'PREV PAGADA' debe ser un número mayor que cero.");
    return;
  }

  if (datosFila[8] === "") {
    Browser.msgBox("Error: El campo 'Tipo de pago' es obligatorio.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Registrar el cobro en Staging Cobros
  var hojaStagingCobros = ss.getSheetByName("Staging Cobros") || crearStagingCobros(ss);
  var newData = [
    transactionId,
    obtenerFechaActual(),
    datosFila[2], // paciente
    datosFila[3], // doctor
    datosFila[8], // tipo de pago
    prevAbonada, // monto
    datosFila[10], // tratamiento
  ];
  hojaStagingCobros.appendRow(newData);

  // Actualizar en Staging Previsiones
  var hojaStaging = ss.getSheetByName("Staging Previsiones");
  if (hojaStaging) {
    //Hace append de la fila que se está agrergando

    var dataRange = hojaStaging.getDataRange();
    var valores = dataRange.getValues();

    var columnaID = 0;          // Columna A - ID TRANSACCION
    var columnaFechaActual = 1; // Columna B - FECHA ACTUAL
    var columnaH = 7;
    var fechaExcluirObj = new Date(fechaExcluir);
    // Solo conservamos año, mes y día para la comparación
    var fechaExcluirStr = fechaExcluirObj.toDateString()

    var rows = contarOcurrenciasID(transactionId, valores);

    var filasActualizadas = []
    if (rows === 0) {
        var saldoPendActual = importeTotal -prevAbonada;
      }else if (rows > 0) {
        var saldoPendActual = saldoPendAnterior - prevAbonada;
      }
    var cont = 0;
    for (var i = 1; i < valores.length; i++) {
        var fila = valores[i];
        var idActual = fila[columnaID].toString();
        
        // Verificar si el ID coincide con el ID buscado
        if (idActual === transactionId.toString()) {
          cont += 1;
          var fechaActual = fila[columnaFechaActual];
          
          // Convertir a string para comparar solo año, mes y día
          var fechaFilaStr = fechaActual.toDateString();
          hojaStaging.getRange(i + 1, columnaH + 1).setValue(saldoPendActual);

          // Actualizar el estado basado en el nuevo saldo pendiente
if (saldoPendActual === 0) {
    hojaStaging.getRange(i + 1, 12).setValue("PAGADO");
    // Opcionalmente, puedes aplicar un formato especial
    hojaStaging.getRange(i + 1, 1, 1, 12).setBackground("#98e0fa");
} else {
    hojaStaging.getRange(i + 1, 12).setValue("PENDIENTE");
}
          if (fechaFilaStr === fechaExcluirStr) {
              hojaStaging.getRange(i + 1, 6 + 1).setValue(prevAbonada);
          }
        }
  }}
  
  
  // Actualizar las vistas
  actualizarDropdownAnosCobros();
  actualizarVistaCobros();
  actualizarVistaPrevisiones();

  var ui = SpreadsheetApp.getUi();
  ui.alert('¡Operación exitosa!', 'El cobro se ha registrado apropiadamente', ui.ButtonSet.OK);
}

//////////////BALANCE GENERAL/////////////////////

/// Trackear las hojas que cumplan con fecha de filtro 
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
        //Logger.log("Comparando: " + mesA + " con " + mesB);
        return mesesValidos.indexOf(mesA) - mesesValidos.indexOf(mesB);
      });
    
      hojasValidas.forEach(hoja => {
       // Logger.log(hoja.getName());
      });
      
      return hojasValidas;
    }

/// Calcular valores de cada variable
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
  hojaBalance.getRange("B5:H16").clearContent();

  for (let hoja of hojas) {
    try {
      var nombreHoja = hoja.getName();
      var mes = nombreHoja.split(" de ")[0];
     // Logger.log('Procesando: ' + nombreHoja + ' (mes: ' + mes + ')');

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
      //Logger.log(`Mes: ${mes}, Fila destino: ${filaDestino}, Suma: ${suma}`);
      
      if (filaDestino) {
        hojaBalance.getRange(filaDestino, 7).setValue(suma);
        hojaBalance.getRange(filaDestino, 3).setValue(n_pacientes);
        hojaBalance.getRange(filaDestino, 5).setValue(n_presupuesto);
        hojaBalance.getRange(filaDestino, 4).setValue(suma_pre);
        hojaBalance.getRange(filaDestino, 6).setValue(suma_pre/n_presupuesto);
        hojaBalance.getRange(filaDestino, 8).setValue(pac_acep);
        hojaBalance.getRange(filaDestino, 2).setValue(abonoMes);
       // Logger.log(`Valor ${suma} escrito en fila ${filaDestino} para ${mes}`);
      } else {
        //Logger.log(`No se encontró fila destino para el mes: ${mes}`);
      }

    } catch (error) {
      //Logger.log(`Error en ${nombreHoja}: ${error}`);
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
  //Logger.log(listaAnios);
  return listaAnios;
}
/// Actualizar filtro en balance general
function actualizarFiltroDeAnios() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BALANCE GENERAL");
  if (!hoja) {
    //Logger.log("La hoja 'BALANCE GENERAL' no existe.");
    return;
  }

  // Obtener la lista de años únicos ordenados
  var anios = obtenerAniosDeHojas(); // Esta es la función que hicimos antes

  if (anios.length === 0) {
    //Logger.log("No hay años para agregar al filtro.");
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
  //Logger.log("Filtro de años actualizado con: " + anios);
}



// function aplicarFiltroPorAnio(anio) {
//   var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BALANCE GENERAL");
//   Logger.log("Aplicando filtro con el año: " + anio);

// }
    
// function balanceGeneral(annio) {
//   //var anios = obtenerAniosDeHojas()
//   var listaHojas = obtenerHojasPorMesYAnio(annio)
//   var sumas = obtenerSumasHojas(listaHojas)
// }
// Reemplaza la función balanceGeneral para que use la misma lógica que balanceGeneralComparacion
function balanceGeneral(annio) {
  var listaHojas = obtenerHojasPorMesYAnio(annio);
  var mesAFila = {
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
  hojaBalance.getRange("B5:H16").clearContent();

  for (var i = 0; i < listaHojas.length; i++) {
    var hoja = listaHojas[i];
    try {
      var nombreHoja = hoja.getName();
      var mes = nombreHoja.split(" de ")[0];
      //Logger.log('Procesando tabla principal: ' + nombreHoja + ' (mes: ' + mes + ')');

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

        var n_pacientes = 0;
        for (var j = 0; j < pacientes.length; j++) {
          if (pacientes[j][0]) n_pacientes++;
        }
        
        var n_presupuesto = 0;
        for (var j = 0; j < presupuestos.length; j++) {
          if (presupuestos[j][0]) n_presupuesto++;
        }
      } else {
        var n_pacientes = 0;
        var n_presupuesto = 0;
      }

      var abonoMes = sumarAbonosPorMes(mes);

      for (var j = 0; j < aceptados.length; j++) {
        if (aceptados[j][0] && aceptados[j][0] === 'Aceptado') {
          pac_acep += 1;
        }
      }

      for (var j = 0; j < presupuestos.length; j++) {
        if (presupuestos[j][0] && typeof presupuestos[j][0] === 'number') {
          suma_pre += presupuestos[j][0];
        }
      }

      for (var j = 0; j < valores.length; j++) {
        if (valores[j][0] && typeof valores[j][0] === 'number') {
          suma += valores[j][0];
        }
      }
      
      var filaDestino = mesAFila[mes];
     // Logger.log(`Mes: ${mes}, Fila destino para tabla principal: ${filaDestino}, Suma: ${suma}`);
      
      if (filaDestino) {
        hojaBalance.getRange(filaDestino, 7).setValue(suma);
        hojaBalance.getRange(filaDestino, 3).setValue(n_pacientes);
        hojaBalance.getRange(filaDestino, 5).setValue(n_presupuesto);
        hojaBalance.getRange(filaDestino, 4).setValue(suma_pre);
        hojaBalance.getRange(filaDestino, 6).setValue(n_presupuesto > 0 ? suma_pre/n_presupuesto : 0);
        hojaBalance.getRange(filaDestino, 8).setValue(pac_acep);
        hojaBalance.getRange(filaDestino, 2).setValue(abonoMes);
        //Logger.log(`Valor ${suma} escrito en fila ${filaDestino} para ${mes} (tabla principal)`);
      }
    } catch (error) {
      Logger.log(`Error en ${nombreHoja} (tabla principal): ${error}`);
    }
  }
    // Agregar la nueva funcionalidad de distribución de presupuestos
  obtenerDistribucionPresupuestos(listaHojas, false);
}

////////////////////////////////////////////////////////// BALANCE GENERAL COMPARACIÓN
// Función para el balance general de la tabla de comparación
function balanceGeneralComparacion(annio) {
  var listaHojas = obtenerHojasPorMesYAnio(annio);
  var mesAFila = {
    "Enero": 29,
    "Febrero": 30,
    "Marzo": 31,
    "Abril": 32,
    "Mayo": 33,
    "Junio": 34,
    "Julio": 35,
    "Agosto": 36,
    "Septiembre": 37,
    "Octubre": 38,
    "Noviembre": 39,
    "Diciembre": 40
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBalance = ss.getSheetByName('BALANCE GENERAL');
  hojaBalance.getRange("B29:H40").clearContent();

  for (var i = 0; i < listaHojas.length; i++) {
    var hoja = listaHojas[i];
    try {
      var nombreHoja = hoja.getName();
      var mes = nombreHoja.split(" de ")[0];
      //Logger.log('Procesando comparación: ' + nombreHoja + ' (mes: ' + mes + ')');

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

        var n_pacientes = 0;
        for (var j = 0; j < pacientes.length; j++) {
          if (pacientes[j][0]) n_pacientes++;
        }
        
        var n_presupuesto = 0;
        for (var j = 0; j < presupuestos.length; j++) {
          if (presupuestos[j][0]) n_presupuesto++;
        }
      } else {
        var n_pacientes = 0;
        var n_presupuesto = 0;
      }

      var abonoMes = sumarAbonosPorMes(mes);

      for (var j = 0; j < aceptados.length; j++) {
        if (aceptados[j][0] && aceptados[j][0] === 'Aceptado') {
          pac_acep += 1;
        }
      }

      for (var j = 0; j < presupuestos.length; j++) {
        if (presupuestos[j][0] && typeof presupuestos[j][0] === 'number') {
          suma_pre += presupuestos[j][0];
        }
      }

      for (var j = 0; j < valores.length; j++) {
        if (valores[j][0] && typeof valores[j][0] === 'number') {
          suma += valores[j][0];
        }
      }
      
      var filaDestino = mesAFila[mes];
      //Logger.log(`Mes: ${mes}, Fila destino para comparación: ${filaDestino}, Suma: ${suma}`);
      
      if (filaDestino) {
        hojaBalance.getRange(filaDestino, 7).setValue(suma);
        hojaBalance.getRange(filaDestino, 3).setValue(n_pacientes);
        hojaBalance.getRange(filaDestino, 5).setValue(n_presupuesto);
        hojaBalance.getRange(filaDestino, 4).setValue(suma_pre);
        hojaBalance.getRange(filaDestino, 6).setValue(n_presupuesto > 0 ? suma_pre/n_presupuesto : 0);
        hojaBalance.getRange(filaDestino, 8).setValue(pac_acep);
        hojaBalance.getRange(filaDestino, 2).setValue(abonoMes);
       // Logger.log(`Valor ${suma} escrito en fila ${filaDestino} para ${mes} (comparación)`);
      }
    } catch (error) {
     // Logger.log(`Error en ${nombreHoja} (comparación): ${error}`);
    }
  }
      // Agregar la nueva funcionalidad de distribución de presupuestos
  obtenerDistribucionPresupuestos(listaHojas, true);
}


// Nueva función dedicada para actualizar el Balance General
function actualizarBalanceGeneral() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBalance = ss.getSheetByName('BALANCE GENERAL');
  
  if (!hojaBalance) {
    Browser.msgBox("Error", "No se encontró la hoja 'BALANCE GENERAL'", Browser.Buttons.OK);
    return;
  }
  
  // Mostrar mensaje de inicio
  SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando actualización del Balance General...", "Actualizando");
  
  // Obtener los años seleccionados
  var anioTabla1 = hojaBalance.getRange("A2").getValue();
  var anioTabla2 = hojaBalance.getRange("A26").getValue();
  
  // Actualizar las tablas según los años seleccionados
  if (anioTabla1) {
    var anioEntero1 = parseInt(anioTabla1, 10);
    if (!isNaN(anioEntero1)) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Actualizando tabla principal para año " + anioEntero1, "Actualizando");
      balanceGeneral(anioEntero1);
    }
  }
  
  if (anioTabla2) {
    var anioEntero2 = parseInt(anioTabla2, 10);
    if (!isNaN(anioEntero2)) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Actualizando tabla de comparación para año " + anioEntero2, "Actualizando");
      balanceGeneralComparacion(anioEntero2);
    }
  }
  
  // Mostrar mensaje de finalización
  Browser.msgBox("Actualización Completa", "Las tablas del Balance General han sido actualizadas correctamente.", Browser.Buttons.OK);
}

// Función para forzar la actualización de ambas tablas
function actualizarTodasLasTablas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBalance = ss.getSheetByName('BALANCE GENERAL');
  
  if (!hojaBalance) {
    //Logger.log("La hoja 'BALANCE GENERAL' no existe.");
    return;
  }
  
  // Obtener los años seleccionados actualmente
  var anioTabla1 = hojaBalance.getRange("A2").getValue();
  var anioTabla2 = hojaBalance.getRange("A26").getValue();
  
  // Actualizar ambas tablas
  if (anioTabla1) {
    var anioEntero1 = parseInt(anioTabla1, 10);
    if (!isNaN(anioEntero1)) {
      balanceGeneral(anioEntero1);
    }
  }
  
  if (anioTabla2) {
    var anioEntero2 = parseInt(anioTabla2, 10);
    if (!isNaN(anioEntero2)) {
      balanceGeneralComparacion(anioEntero2);
         balanceGeneralComparacion(anioEntero2);
    }
  }
}

// Agregar esta función para actualizar ambas tablas después de guardar nuevos datos
function guardarDatosYActualizarTablas() {
  // Llamar primero a la función original
  guardarDatosEnTabla2();
  
  // Luego actualizar las tablas
  actualizarTodasLasTablas();
}






/////////////////////////////////////////////////////////////////////

////////VISTA COBROS///////
function crearVistaCobros(ss) {
    var hojaVista = ss.getSheetByName("Vista Cobros");
    if (!hojaVista) {
        hojaVista = ss.insertSheet("Vista Cobros");
        configurarVistaCobros(hojaVista, ss);
    }
    return hojaVista;
}

function configurarVistaCobros(hojaVista, ss) {
    // Configurar título
    hojaVista.getRange("A1:B1").setValue("Cobros");
    hojaVista.getRange("A2").setValue("Año:");
    hojaVista.getRange("A3").setValue("Mes:");
    
    // Formato para título y filtros
    hojaVista.getRange("A1:B1")
        .setValue("Cobros")
        .setFontSize(20)
        .setFontWeight("bold")
        .setBackground("#00c896")
        .setHorizontalAlignment("center")
        .merge();

    hojaVista.getRange("A2")
        .setValue("Año:")
        .setFontSize(12)
        .setFontWeight("bold")
        .setBackground("#999999")
        .setFontColor("white");

    hojaVista.getRange("A3")
        .setValue("Mes:")
        .setFontSize(12)
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white");
        
    hojaVista.getRange("B3").setHorizontalAlignment("right");

    // Fechas de control (ocultas)
    hojaVista.getRange("P1").setValue("Fechas de control");
    hojaVista.getRange("P2:P3").setValues([["Fecha inicio"], ["Fecha fin"]]);
    hojaVista.getRange("P2").setFormula('=IF(B3="Ver todo el año",DATE(B2,1,1),DATE(B2,MATCH(B3,{"Enero";"Febrero";"Marzo";"Abril";"Mayo";"Junio";"Julio";"Agosto";"Septiembre";"Octubre";"Noviembre";"Diciembre"},0),1))');
    hojaVista.getRange("P3").setFormula('=IF(B3="Ver todo el año",DATE(B2,12,31),EOMONTH(P2,0))');
    hojaVista.hideColumns(16, 1);

    // Encabezados de la tabla
    var encabezados = [
        "ID TRANSACCIÓN", "FECHA DE COBRO", "PACIENTE", "DOCTOR", "TIPO DE PAGO", "MONTO", "TRATAMIENTO"
    ];

    hojaVista.getRange(5, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Agregar filtros a los encabezados
    hojaVista.getRange(5, 1, 1, encabezados.length).createFilter();

    // Configurar formato de columnas
    hojaVista.getRange("F:F").setNumberFormat("€#,##0.00");

    // Configurar validaciones para filtros
    actualizarDropdownAnosCobros(); // Necesitas crear esta función
    
    // Configurar meses
    var meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        "Ver todo el año"
    ];

    var validacionMes = SpreadsheetApp.newDataValidation()
        .requireValueInList(meses)
        .setAllowInvalid(false)
        .build();
    hojaVista.getRange("B3").setDataValidation(validacionMes);

    // Añadir tabla resumen
    configurarTablaResumenCobros(hojaVista);
}

function configurarTablaResumenCobros(hojaVista) {
    var resumenEncabezados = [
        ["RESUMEN COBROS", "", ""],
        ["Total Cobrado", "=SUM(F6:F)", ""]
    ];
 
    var rangoResumen = hojaVista.getRange(1, 12, resumenEncabezados.length, 3);
    rangoResumen.setValues(resumenEncabezados);
 
    // Aplicar formatos
    hojaVista.getRange("L1:M1")
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold")
        .merge();
    
    hojaVista.getRange("M2").setNumberFormat("€#,##0.00");
    hojaVista.getRange("L2:M2").setBackground("#f6f6f6");
    
    // Bordes
    hojaVista.getRange("L1:M2").setBorder(true, true, true, true, true, true);
}

function actualizarDropdownAnosCobros() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaStaging = ss.getSheetByName("Staging Cobros");
    var hojaVista = ss.getSheetByName("Vista Cobros");
    
    if (!hojaStaging || !hojaVista) return;
    
    // Obtener todos los datos de Staging Cobros
    var datos = hojaStaging.getDataRange().getValues();
    var annos = new Set();
    
    // Procesar las fechas para obtener años únicos
    for (var i = 1; i < datos.length; i++) {
        if (datos[i][1]) {  // Columna B (índice 1) contiene las fechas de cobro
            var fecha;
            if (datos[i][1] instanceof Date) {
                fecha = datos[i][1];
            } else {
                fecha = new Date(datos[i][1]);
            }
            
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

function actualizarVistaCobros() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaVista = ss.getSheetByName("Vista Cobros");
    var hojaStaging = ss.getSheetByName("Staging Cobros");
    
    if (!hojaStaging || !hojaVista) return;
    
    var anno = hojaVista.getRange("B2").getValue();
    var mes = hojaVista.getRange("B3").getValue();
    
    if (!anno) return; // Solo requerimos el año

    // Definir fechas de inicio y fin según si hay mes seleccionado
    var fechaInicio, fechaFin;

    if (mes && mes !== "Ver todo el año") {
        fechaInicio = hojaVista.getRange("P2").getValue();
        fechaFin = hojaVista.getRange("P3").getValue();
    } else {
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
    } else {
        hojaVista.getRange(6, 1).setValue("No hay datos para mostrar");
    }
}

///////////////Balance general extra
// Función para obtener la distribución de presupuestos por rangos de montos
function obtenerDistribucionPresupuestos(hojas,  esComparacion = false) {
  // Mapeo de meses a filas (tabla principal o comparación)
  const mesAFila = esComparacion ? {
    "Enero": 29, "Febrero": 30, "Marzo": 31, "Abril": 32,
    "Mayo": 33, "Junio": 34, "Julio": 35, "Agosto": 36,
    "Septiembre": 37, "Octubre": 38, "Noviembre": 39, "Diciembre": 40
  } : {
    "Enero": 5, "Febrero": 6, "Marzo": 7, "Abril": 8,
    "Mayo": 9, "Junio": 10, "Julio": 11, "Agosto": 12,
    "Septiembre": 13, "Octubre": 14, "Noviembre": 15, "Diciembre": 16
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBalance = ss.getSheetByName('BALANCE GENERAL');
  
  // Limpiar la tabla de distribución correspondiente
  if (esComparacion) {
    hojaBalance.getRange("K29:O40").clearContent();
    hojaBalance.getRange("R29:V40").clearContent();
  } else {
    hojaBalance.getRange("K5:O16").clearContent();
    hojaBalance.getRange("R5:V16").clearContent();
  }
  
  // Definir los rangos de montos
  const rangos = [
    { min: 0, max: 1000 },      // Columna K (0-1000)
    { min: 1000, max: 3000 },   // Columna L (1000-3000)
    { min: 3000, max: 6000 },   // Columna M (3000-6000)
    { min: 6000, max: 10000 },  // Columna N (6000-10000)
    { min: 10000, max: Infinity } // Columna O (>10000)
  ];
  
  // Procesar cada hoja mensual
  for (let hoja of hojas) {
    try {
      var nombreHoja = hoja.getName();
      var mes = nombreHoja.split(" de ")[0];
      //Logger.log('Procesando distribución para: ' + nombreHoja + ' (mes: ' + mes + ')');
      
      var lastRow = hoja.getLastRow();
      var startRow = 18; // Primera fila de datos
      
      // Inicializar contadores para cada rango
      var contadores = [0, 0, 0, 0, 0];

      var count_tipol = [0, 0, 0, 0, 0];
      
      if (lastRow >= startRow) {
        // Obtener los importes presupuestados (columna 11)
        var presupuestos = hoja.getRange(startRow, 11, lastRow - startRow + 1, 1).getValues();

        var tipologias = hoja.getRange(18, 7, lastRow - 17, 1).getValues();
        var subtipologias = hoja.getRange(18, 8, lastRow - 17, 1).getValues();

        for (let i = 0; i < tipologias.length; i++) {
          if (tipologias[i][0] === "1VTA") {
            count_tipol[0]++;
          } else if(tipologias[i][0] ==="AMP") {
            count_tipol[1]++;
          } else if (tipologias[i][0] ==="Criba") {
            count_tipol[2]++;
          } else if (tipologias[i][0] ==="OC") {
              if (subtipologias[i][0] === "OC Llamado"){
                count_tipol[3]++;
              } else if (subtipologias[i][0] === "OC Vino él"){
                count_tipol[4]++;
              }
          }
        }
        
        // Contar presupuestos por rango
        for (let i = 0; i < presupuestos.length; i++) {
          var importe = presupuestos[i][0];
          
          // Solo considerar valores numéricos válidos
          if (typeof importe === 'number' && !isNaN(importe)) {
            // Clasificar en el rango correspondiente
            for (let j = 0; j < rangos.length; j++) {
              if (importe > rangos[j].min && importe <= rangos[j].max) {
                contadores[j]++;
                break;
              }
            }
          }
        }
      }
      
      // Escribir resultados en la hoja de balance
      var filaDestino = mesAFila[mes];
      if (filaDestino) {
        for (let i = 0; i < contadores.length; i++) {
          hojaBalance.getRange(filaDestino, 11 + i).setValue(contadores[i]);
        }
       // Logger.log(`Contadores para ${mes}: ${contadores.join(', ')}`);
      }
      for (let i = 0; i < count_tipol.length; i++) {
        hojaBalance.getRange(filaDestino, 18 + i).setValue(count_tipol[i]);
      }
      
    } catch (error) {
      Logger.log(`Error al procesar distribución para ${nombreHoja}: ${error}`);
    }
  }
  
}
// function actualizarKPIs() {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var hojaKPIs = ss.getSheetByName('Análisis de KPIs');
    
//     if (!hojaKPIs) {
//         Browser.msgBox("Error", "No se encontró la hoja 'Análisis de KPIs'", Browser.Buttons.OK);
//         return;
//     }
    
//     // Mostrar mensaje de inicio
//     SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando actualización de KPIs...", "Actualizando");
    
//     // Obtener las fechas seleccionadas actualmente
//     var fechaInicio = hojaKPIs.getRange('D2').getValue();
//     var fechaFin = hojaKPIs.getRange('D3').getValue();
    
//     if (!fechaInicio || !fechaFin) {
//         Browser.msgBox("Error", "No hay fechas seleccionadas para actualizar KPIs. Por favor seleccione un rango de fechas primero.", Browser.Buttons.OK);
//         return;
//     }
    
//     // Verificar que las fechas sean del mismo mes
//     if (!sonDelMismoMes(fechaInicio, fechaFin)) {
//         Browser.msgBox("Error", "Las fechas deben ser del mismo mes para actualizar KPIs.", Browser.Buttons.OK);
//         return;
//     }
    
//     // Ejecutar la función que actualiza los KPIs con las fechas actuales
//     saveDateRange(fechaInicio, fechaFin);
    
//     // Mostrar mensaje de finalización
//     Browser.msgBox("Actualización Completa", "Los KPIs han sido actualizados correctamente para el rango de fechas seleccionado.", Browser.Buttons.OK);
// }
function actualizarKPIs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaKPIs = ss.getSheetByName('Análisis de KPIs');
    
    if (!hojaKPIs) {
      Browser.msgBox("Error", "No se encontró la hoja 'Análisis de KPIs'. Por favor, asegúrese de que existe esta hoja en el libro.", Browser.Buttons.OK);
      return;
    }
    
    // Mostrar mensaje de inicio
    SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando verificación de estructura...", "Preparando");
    
    // Obtener las fechas seleccionadas actualmente
    var fechaInicio = hojaKPIs.getRange('D2').getValue();
    var fechaFin = hojaKPIs.getRange('D3').getValue();
    
    // Si no hay fechas seleccionadas, limpiar todos los valores de KPIs
    if (!fechaInicio || !fechaFin) {
      // Limpiar todas las áreas de datos de KPIs
      SpreadsheetApp.getActiveSpreadsheet().toast("No hay fechas seleccionadas. Limpiando valores de KPIs...", "Limpiando");
      
      // Limpiar análisis por rango de importes
      hojaKPIs.getRange("B5:C9").setValue(0);
      
      // Limpiar análisis de previsiones y cobros
      hojaKPIs.getRange("C15:C17").setValue(0);
      hojaKPIs.getRange("C19").setValue(0);
      
      // Limpiar análisis por estado
      hojaKPIs.getRange("B23:C26").setValue(0);
      
      // Limpiar análisis por tipología
      hojaKPIs.getRange("B30:C36").setValue(0);
      
      // Limpiar análisis por tipo de pago
      hojaKPIs.getRange("C40:C43").setValue(0);
      
      // Limpiar tabla de análisis por doctor
      var ultimaFila = hojaKPIs.getLastRow();
      if (ultimaFila >= 49) {
        hojaKPIs.getRange(49, 1, ultimaFila - 48, hojaKPIs.getLastColumn()).clearContent();
      }
      
      Browser.msgBox("Limpieza Completa", "Se han limpiado todos los valores de KPIs. Para ver datos nuevos, seleccione un rango de fechas usando el botón 'SELECCIONAR'.", Browser.Buttons.OK);
      return;
    }
    
    // Verificar que las hojas necesarias existan
    var hojaStagingPrevisiones = ss.getSheetByName("Staging Previsiones");
    var hojaStagingCobros = ss.getSheetByName("Staging Cobros");
    
    if (!hojaStagingPrevisiones) {
      Browser.msgBox("Error", "No se encontró la hoja 'Staging Previsiones' que es necesaria para el cálculo de KPIs.", Browser.Buttons.OK);
      return;
    }
    
    if (!hojaStagingCobros) {
      Browser.msgBox("Error", "No se encontró la hoja 'Staging Cobros' que es necesaria para el cálculo de KPIs.", Browser.Buttons.OK);
      return;
    }
    
    // Verificar que las fechas sean del mismo mes
    if (!sonDelMismoMes(fechaInicio, fechaFin)) {
      Browser.msgBox("Error", "Las fechas deben ser del mismo mes para actualizar KPIs.", Browser.Buttons.OK);
      return;
    }
    
    // Verificar si existe la hoja mensual correspondiente
    var hojaMes = formatearFecha(fechaInicio);
    var hojaMensual = ss.getSheetByName(hojaMes);
    
    if (!hojaMensual) {
      Browser.msgBox("Error", `No se encontró la hoja '${hojaMes}' necesaria para el análisis. Es posible que aún no haya datos para este período.`, Browser.Buttons.OK);
      return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Iniciando actualización de KPIs...", "Actualizando");
    
    // Ejecutar la función que actualiza los KPIs con las fechas actuales
    try {
      saveDateRange(fechaInicio, fechaFin);
      // Mostrar mensaje de finalización
      Browser.msgBox("Actualización Completa", "Los KPIs han sido actualizados correctamente para el rango de fechas seleccionado.", Browser.Buttons.OK);
    } catch (e) {
      Browser.msgBox("Error en la actualización", "Se produjo un error al actualizar los KPIs: " + e.toString(), Browser.Buttons.OK);
      Logger.log("Error en saveDateRange: " + e.toString());
    }
  } catch (e) {
    Browser.msgBox("Error general", "Se produjo un error inesperado: " + e.toString(), Browser.Buttons.OK);
    Logger.log("Error general en actualizarKPIs: " + e.toString());
  }
}