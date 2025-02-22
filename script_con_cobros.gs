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

////// función principal
function guardarDatosEnTabla2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFormulario = ss.getSheetByName("Registro de transacciones");

    if (!hojaFormulario) {
        Logger.log("Error: No se encontró la hoja 'Registro de transacciones'");
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

    // Generar ID de transacción
    const transactionId = generateSequentialTransactionId();

    var fecha = new Date(fechaIngresada);
    var nombreMes = fecha.toLocaleDateString("es-ES", { year: "numeric", month: "long" });
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    var hojaMes = ss.getSheetByName(nombreMes) || crearHojaMes(ss, nombreMes);
    var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);
    var hojaPrevisiones = ss.getSheetByName("Staging Previsiones");
    if (!hojaPrevisiones) {
        hojaPrevisiones = crearHojaPrevisiones(ss);
    }
        // Crear o obtener la hoja Vista Previsiones
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    if (!hojaVista) {
        hojaVista = crearVistaPrevisiones(ss);
    }

    var filaEscribir = hojaMes.getLastRow() < 17 ? 18 : hojaMes.getLastRow() + 1;

    var nuevaFila = [
        transactionId,     // ID Transacción
        fechaIngresada,    // FECHA DE CONTACTO
        datos[0][1],       // PACIENTE
        datos[4][1],       // TELÉFONO
        datos[9][0],       // DOCTOR/A
        datos[9][1],       // AUXILIAR
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
        agregarAPacientesAceptados(hojaCobros, transactionId, datos[0][1], datos[14][3], datos[14][1], datos[9][0]);
        agregarAStagingPrevisiones(hojaPrevisiones, transactionId, datos[14][3], datos[0][1], datos[9][0], datos[14][1]);
    }
    actualizarTablaResumen(hojaMes);
    limpiarFormulario(hojaFormulario);
    Logger.log("Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");
}

function crearHojaMes(ss, nombreMes) {
    var hojaMes = ss.insertSheet(nombreMes);
    var encabezados = [
        "ID TRANSACCIÓN", "FECHA DE CONTACTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", 
        "AUXILIAR", "TIPOLOGÍA PV", "SUBTIPOLOGÍA", "PLAN DE CITAS", "ESTADO", 
        "IMPORTE PRESUPUESTADO", "IMPORTE ACEPTADO", "FECHA DE INICIO", "OBSERVACIONES"
    ];

    hojaMes.getRange(17, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center");
    hojaMes.getRange(17, 1, 1, encabezados.length).createFilter();
    hojaMes.autoResizeColumns(1, hojaMes.getLastColumn());
    return hojaMes;
}

function crearHojaCobros(ss, nombreMes) {
    var hojaCobros = ss.insertSheet("Cobros " + nombreMes);
    
    // Tabla principal de cobros con nueva columna ID
    var encabezados = ["ID TRANSACCIÓN", "FECHA DE COBRO", "PACIENTE", "DOCTOR", "IMPORTE TOTAL", "TIPO DE PAGO", "ESTADO DEL COBRO", "TOTAL PAGADO", "FECHA FINAL"];
    hojaCobros.getRange(4, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaCobros.getRange(4, 1, 1, encabezados.length).createFilter();

    var titulo = "Caja de " + nombreMes;
    hojaCobros.getRange(1, 1).setValue(titulo)
        .setFontWeight("bold")
        .setFontSize(18);
    
    // Mover el texto "Modifica manualmente estas columnas" dos columnas
    hojaCobros.getRange(3, 5, 1, 4).merge().setValue("Modifica manualmente estas columnas")
        .setFontSize(10)
        .setBackground("#00c896")
        .setFontColor("#424242")
        .setHorizontalAlignment("center");
    
    // Mover la tabla de resumen dos columnas a la derecha (de columna 10 a columna 12)
    var encabezadosResumen = ["TIPO DE PAGO", "N° PACIENTES", "TOTAL / TIPO"];
    hojaCobros.getRange(4, 12, 1, 3).setValues([encabezadosResumen])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Tipos de pago
    var tiposPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    
    // Insertar tipos de pago y fórmulas con nuevas posiciones de columna
    tiposPago.forEach((tipo, index) => {
        var fila = 5 + index;
        
        // Tipo de pago (columna 12 en lugar de 10)
        hojaCobros.getRange(fila, 12).setValue(tipo)
            .setBackground("#f6f6f6")
            .setHorizontalAlignment("left");
        
        // Fórmula para contar pacientes (columna 13 en lugar de 11)
        var formulaConteo = `=COUNTIFS(F:F,"${tipo}")`;
        hojaCobros.getRange(fila, 13).setFormula(formulaConteo)
            .setBackground("#e2e2e2")
            .setHorizontalAlignment("center");
        
        // Fórmula para sumar montos (columna 14 en lugar de 12)
        var formulaSuma = `=SUMIF(F:F,"${tipo}",H:H)`;
        hojaCobros.getRange(fila, 14).setFormula(formulaSuma)
            .setBackground("#f6f6f6")
            .setHorizontalAlignment("right")
            .setNumberFormat("€#,##0.00");
    });
    
    // Agregar totales con nuevas posiciones
    var filaTotales = 5 + tiposPago.length;
    hojaCobros.getRange(filaTotales, 12).setValue("TOTAL PREVISTO")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white");
    
    // Fórmula para total de pacientes (columna 13)
    hojaCobros.getRange(filaTotales, 13).setFormula("=SUM(M5:M8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Fórmula para total de montos (columna 14)
    hojaCobros.getRange(filaTotales, 14).setFormula("=SUM(N5:N8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("right")
        .setNumberFormat("€#,##0.00");

    // Fila "TOTAL COBRADO" con nuevas posiciones
    var filaTotalCobrado = filaTotales + 1;
    hojaCobros.getRange(filaTotalCobrado, 12).setValue("TOTAL COBRADO")
        .setFontWeight("bold")
        .setBackground("#999999")
        .setFontColor("black");

    // Columna del medio vacía (columna 13)
    hojaCobros.getRange(filaTotalCobrado, 13).setValue("")
        .setBackground("#999999");

    // Total cobrado con nueva posición (columna 14)
    hojaCobros.getRange(filaTotalCobrado, 14).setFormula('=SUM(H:H)')
        .setFontWeight("bold")
        .setBackground("#999999")
        .setFontColor("black")
        .setHorizontalAlignment("right")
        .setNumberFormat("€#,##0.00");

    // Validación para método de pago
    var reglaValidacion_cobro = SpreadsheetApp.newDataValidation()
        .requireValueInList(tiposPago, true)
        .setAllowInvalid(false)
        .build();

    // Ajustar ancho de columnas (ahora hasta la columna 14)
    hojaCobros.autoResizeColumns(1, 14);
    
    return hojaCobros;
}

///////PREVISIONES STAGING
function crearHojaPrevisiones(ss) {
    var hojaPrevisiones = ss.insertSheet("Staging Previsiones");
    var encabezados = [
        "ID TRANSACCIÓN",  // Nueva columna ID
        "FECHA ACTUAL", 
        "PACIENTE", 
        "DOCTOR", 
        "IMPORTE TOTAL", 
        "ABONO", 
        "SALDO PENDIENTE", 
        "TIPO DE PAGO", 
        "PRÓXIMO PAGO"
    ];

    hojaPrevisiones.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaPrevisiones.autoResizeColumns(1, hojaPrevisiones.getLastColumn());
    return hojaPrevisiones;
}

function agregarAStagingPrevisiones(hojaPrevisiones, transactionId, fechaInicio, paciente, doctor, importeAceptado) {
    // Verificar si el ID ya existe
    if (existeIdEnHoja(hojaPrevisiones, transactionId)) {
        Logger.log("ID ya existe en Staging Previsiones: " + transactionId);
        Browser.msgBox("Error", `El ID de transacción ${id} ya existe en la hoja ${hoja.getName()}`, Browser.Buttons.OK);
        return;
    }

    var ultimaFila = hojaPrevisiones.getLastRow() + 1;
    var tipoPagoOpciones = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];

    var nuevaFila = [
        transactionId,
        fechaInicio,
        paciente,
        doctor,
        importeAceptado,
        "",
        `=IF(ISBLANK(F${ultimaFila}), E${ultimaFila}, E${ultimaFila}-F${ultimaFila})`,
        "",
        ""
    ];

    hojaPrevisiones.getRange(ultimaFila, 1, 1, nuevaFila.length).setValues([nuevaFila]);
    
    // Aplicar formatos
    hojaPrevisiones.getRange(ultimaFila, 5).setNumberFormat("€#,##0.00");
    hojaPrevisiones.getRange(ultimaFila, 6).setNumberFormat("€#,##0.00");
    hojaPrevisiones.getRange(ultimaFila, 7).setNumberFormat("€#,##0.00");
    hojaPrevisiones.getRange(ultimaFila, 8).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(tipoPagoOpciones, true).setAllowInvalid(false).build()
    );
    hojaPrevisiones.getRange(ultimaFila, 9).setDataValidation(
        SpreadsheetApp.newDataValidation().requireDate().build()
    );
}


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
            break;
        }
    }
    
    // Actualizar saldo pendiente si es necesario
    if (columna === 6) { // Si se modificó el abono
        var importeTotal = hojaVista.getRange(fila, 5).getValue();
        var saldoPendiente = importeTotal - nuevoValor;
        hojaVista.getRange(fila, 7).setValue(saldoPendiente);
    }
}

// crear hoja de visualizaciones de previsiones
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
    hojaVista.clear();

    // Configuración inicial igual que antes...
    hojaVista.getRange("A1").setValue("Filtros");
    hojaVista.getRange("A2").setValue("Año:");
    hojaVista.getRange("A3").setValue("Mes:");
    
    // Validaciones de año y mes igual que antes...
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    var datos = hojaStaging.getDataRange().getValues();
    var annos = new Set();
    
    for (var i = 1; i < datos.length; i++) {
        if (datos[i][1]) {
            var fecha = new Date(datos[i][1]);
            annos.add(fecha.getFullYear());
        }
    }
    
    var annosArray = Array.from(annos).sort();
    var validacionAnno = SpreadsheetApp.newDataValidation()
        .requireValueInList(annosArray)
        .setAllowInvalid(false)
        .build();
    hojaVista.getRange("B2").setDataValidation(validacionAnno);
    
    var meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ];
    var validacionMes = SpreadsheetApp.newDataValidation()
        .requireValueInList(meses)
        .setAllowInvalid(false)
        .build();
    hojaVista.getRange("B3").setDataValidation(validacionMes);

    // Celdas auxiliares para fechas
    hojaVista.getRange("D1").setFormula('=IF(B2="","",B2)');
    hojaVista.getRange("D2").setFormula(`=IF(OR(B2="",B3=""),"",DATE(B2,
        SWITCH(B3,
        "Enero",1,"Febrero",2,"Marzo",3,"Abril",4,"Mayo",5,"Junio",6,
        "Julio",7,"Agosto",8,"Septiembre",9,"Octubre",10,"Noviembre",11,"Diciembre",12),1))`);
    hojaVista.getRange("D3").setFormula('=IF(OR(B2="",B3=""),"",EOMONTH(D2,0))');
    hojaVista.hideColumn(hojaVista.getRange("D:D"));

    // Encabezados de la tabla
    var encabezados = [
        "ID TRANSACCIÓN", "FECHA", "PACIENTE", "DOCTOR", "IMPORTE TOTAL",
        "ABONO", "SALDO PENDIENTE", "TIPO DE PAGO", "PRÓXIMO PAGO"
    ];
    
    hojaVista.getRange(5, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");

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

// Aplicar validaciones a las columnas completas desde la fila 6
// hojaVista.getRange("H6:H").setDataValidation(validacionTipoPago);  // Tipo de Pago
// hojaVista.getRange("I6:I").setDataValidation(validacionFecha);     // Próximo Pago
// Obtener el rango de datos desde la hoja Staging
var hojaStaging = ss.getSheetByName("Staging Previsiones");
var datosStaging = hojaStaging.getDataRange().getValues();
var numFilasConDatos = datosStaging.length - 1; // Restar 1 por la fila de encabezado

// Aplicar validaciones solo al rango esperado de datos
if (numFilasConDatos > 0) {
    hojaVista.getRange(6, 8, numFilasConDatos).setDataValidation(validacionTipoPago);  // Tipo de Pago
    hojaVista.getRange(6, 9, numFilasConDatos).setDataValidation(validacionFecha);     // Próximo Pago
}

    // Añadir tabla resumen (igual que antes)
    configurarTablaResumen(hojaVista);
}
function configurarTablaResumen(hojaVista) {
    var resumenEncabezados = [
        ["RESUMEN", "", ""],
        ["Total Importe", "=SUM(E6:E)", ""],
        ["Total Abonado", "=SUM(F6:F)", ""],
        ["Total Pendiente", "=SUM(G6:G)", ""]
    ];

    var rangoResumen = hojaVista.getRange(1, 11, resumenEncabezados.length, 3);
    rangoResumen.setValues(resumenEncabezados);
    
    // Aplicar formatos
    hojaVista.getRange("K1:M1")
        .setBackground("#424242")
        .setFontColor("white")
        .setFontWeight("bold")
        .merge();
    
    hojaVista.getRange("L2:L4").setNumberFormat("€#,##0.00");
    
    // Estilos alternados
    hojaVista.getRange("K2:M2").setBackground("#f6f6f6");
    hojaVista.getRange("K3:M3").setBackground("#e2e2e2");
    hojaVista.getRange("K4:M4").setBackground("#f6f6f6");
    
    // Bordes
    hojaVista.getRange("K1:M4").setBorder(true, true, true, true, true, true);
}
function actualizarVistaPrevisiones() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaVista = ss.getSheetByName("Vista Previsiones");
    var hojaStaging = ss.getSheetByName("Staging Previsiones");
    
    var anno = hojaVista.getRange("B2").getValue();
    var mes = hojaVista.getRange("B3").getValue();
    
    if (!anno || !mes) return;
    
    var fechaInicio = hojaVista.getRange("D2").getValue();
    var fechaFin = hojaVista.getRange("D3").getValue();
    
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


///// agregar pacientes a staging de previsiones y cobros
function agregarAPacientesAceptados(hojaCobros, transactionId, paciente, fecha, importe, doctor) {
    // Verificar si el ID ya existe
    if (existeIdEnHoja(hojaCobros, transactionId)) {
        Logger.log("ID ya existe en hoja Cobros: " + transactionId);
        Browser.msgBox("Error", `El ID de transacción ${id} ya existe en la hoja ${hoja.getName()}`, Browser.Buttons.OK);
        return;
    }

    var ultimaFila = 4;
    var datos = hojaCobros.getRange("A5:A").getValues();
    
    for (var i = 0; i < datos.length; i++) {
        if (datos[i][0] !== "") {
            ultimaFila = i + 5;
        }
    }
    
    var filaEscribir = ultimaFila === 4 ? 5 : ultimaFila + 1;
    
    hojaCobros.getRange(filaEscribir, 1, 1, 9).setValues([[
        transactionId, fecha, paciente, doctor, importe, "", "", "", ""
    ]]);
    
    hojaCobros.getRange(filaEscribir, 5).setNumberFormat("€#,##0.00");
    hojaCobros.getRange(filaEscribir, 8).setNumberFormat("€#,##0.00");

    var formulaEstadoCobro = `=IF(H${filaEscribir} < E${filaEscribir}, "Pendiente de pago", 
                              IF(H${filaEscribir} = E${filaEscribir}, "PAGADO", ""))`;
                                  
    hojaCobros.getRange(filaEscribir, 7).setFormula(formulaEstadoCobro);
    
    var opcionesMetodoPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    var reglaValidacion_cobro = SpreadsheetApp.newDataValidation()
        .requireValueInList(opcionesMetodoPago, true)
        .setAllowInvalid(false)
        .build();

    hojaCobros.getRange(filaEscribir, 6).setDataValidation(reglaValidacion_cobro);
    hojaCobros.getRange(filaEscribir, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
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
            var nombreMes = hoja.getName();
            var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);
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
            agregarAPacientesAceptados(hojaCobros, transactionId, paciente, fecha, importe, doctor);
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

//vista previsiones
function obtenerFilaActiva() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var fila = hoja.getActiveCell().getRow(); // Obtiene el número de fila activa
  var datosFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues(); // Obtiene los datos de la fila activa
  Logger.log(datosFila); // Muestra los datos en el registro
}
 
