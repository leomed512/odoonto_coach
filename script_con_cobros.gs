function guardarDatosEnTabla() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFormulario = ss.getSheetByName("Registro de transacciones");

    if (!hojaFormulario) {
        Logger.log("Error: No se encontró la hoja 'Registro de transacciones'");
        return;
    }
    // SpreadsheetApp.flush(); // 🛑 Forzar la actualización de valores antes de leerlos


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

    var fecha = new Date(fechaIngresada);
    var nombreMes = fecha.toLocaleDateString("es-ES", { year: "numeric", month: "long" });
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    var hojaMes = ss.getSheetByName(nombreMes) || crearHojaMes(ss, nombreMes);
    var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);

    var filaEscribir = hojaMes.getLastRow() < 17 ? 18 : hojaMes.getLastRow() + 1;

    var nuevaFila = [
        fechaIngresada, datos[0][1], datos[4][1], datos[9][0], datos[9][1], datos[9][2], datos[9][3], datos[9][4], datos[14][4], 
        datos[14][0], datos[14][1], datos[14][2], datos[14][3], datos[9][5], datos[18][0]
    ];
    hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    // Aplicar formatos y validaciones
    actualizarFormatoFila(hojaMes, filaEscribir, datos[14][4]);
    hojaMes.getRange(filaEscribir, 10).setNumberFormat("€#,##0.00");
    hojaMes.getRange(filaEscribir, 12).setNumberFormat("€#,##0.00");

    if (datos[14][4] === "Aceptado") {
        agregarAPacientesAceptados(hojaCobros, datos[0][1], fechaIngresada, datos[14][2], datos[9][0]);
    }
    actualizarTablaResumen(hojaMes);
    limpiarFormulario(hojaFormulario);
    Logger.log("Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");
}

function crearHojaMes(ss, nombreMes) {
    var hojaMes = ss.insertSheet(nombreMes);
    var encabezados = ["FECHA DE CONTACTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", "AUXILIAR", "TIPOLOGÍA PV", "SUBTIPOLOGÍA", "CAMPAÑA", "ESTADO", "IMPORTE PRESUPUESTADO", "N° PTOS", "IMPORTE ACEPTADO", "FECHA DE INICIO", "PLAN DE CITAS", "OBSERVACIONES"];
    hojaMes.getRange(17, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight("bold").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center");
    hojaMes.getRange(17, 1, 1, encabezados.length).createFilter();
    hojaMes.autoResizeColumns(1, hojaMes.getLastColumn());
    return hojaMes;
}

function crearHojaCobros(ss, nombreMes) {
    var hojaCobros = ss.insertSheet("Cobros " + nombreMes);
    
    // Tabla principal de cobros
    var encabezados = ["FECHA DE COBRO", "PACIENTE", "DOCTOR", "IMPORTE TOTAL", "TIPO DE PAGO", "ESTADO DEL COBRO"];
    hojaCobros.getRange(4, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaCobros.getRange(4, 1, 1, encabezados.length).createFilter();
    
    // Nueva tabla de resumen
    var encabezadosResumen = ["TIPO DE PAGO", "N° PACIENTES", "MONTO"];
    hojaCobros.getRange(4, 10, 1, 3).setValues([encabezadosResumen])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Tipos de pago
    var tiposPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    
    // Insertar tipos de pago y fórmulas
    tiposPago.forEach((tipo, index) => {
        var fila = 5 + index;
        
        // Tipo de pago
        hojaCobros.getRange(fila, 10).setValue(tipo)
            .setBackground("#f6f6f6")
            .setHorizontalAlignment("left");
        
        // Fórmula para contar pacientes
        var formulaConteo = `=COUNTIFS(E:E,"${tipo}")`;
        hojaCobros.getRange(fila, 11).setFormula(formulaConteo)
            .setBackground("#e2e2e2")
            .setHorizontalAlignment("center");
        
        // Fórmula para sumar montos
        var formulaSuma = `=SUMIF(E:E,"${tipo}",D:D)`;
        hojaCobros.getRange(fila, 12).setFormula(formulaSuma)
            .setBackground("#f6f6f6")
            .setHorizontalAlignment("right")
            .setNumberFormat("€#,##0.00");
    });
    
    // Agregar totales
    var filaTotales = 5 + tiposPago.length;
    hojaCobros.getRange(filaTotales, 10).setValue("TOTAL")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white");
    
    // Fórmula para total de pacientes
    hojaCobros.getRange(filaTotales, 11).setFormula("=SUM(I5:I8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Fórmula para total de montos
    hojaCobros.getRange(filaTotales, 12).setFormula("=SUM(J5:J8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("right")
        .setNumberFormat("€#,##0.00");

    // Validación para método de pago
    var reglaValidacion_cobro = SpreadsheetApp.newDataValidation()
        .requireValueInList(tiposPago, true)
        .setAllowInvalid(false)
        .build();

    // Ajustar ancho de columnas
    hojaCobros.autoResizeColumns(1, 12);
    
    return hojaCobros;
}
function actualizarFormatoFila(hoja, fila, estado) {
    var rangoFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn());
    var colores = { "Aceptado": "#54c772", "Pendiente": "#FF9D23", "No aceptado": "#fc4c3d" };
    rangoFila.setBackground(colores[estado] || null);

    var reglaValidacion = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Aceptado", "Pendiente", "No aceptado"], true)
        .setAllowInvalid(false)
        .build();
    hoja.getRange(fila, 9).setDataValidation(reglaValidacion);
}


function agregarAPacientesAceptados(hojaCobros, paciente, fecha, importe, doctor) {
    var ultimaFila = 4; // Empezamos desde la fila de encabezados
    var datos = hojaCobros.getRange("A5:A").getValues();
    
    // la última fila con datos en la columna A
    for (var i = 0; i < datos.length; i++) {
        if (datos[i][0] !== "") {
            ultimaFila = i + 5; // Sumamos 5 porque empezamos desde A5
        }
    }
    
    var filaEscribir = ultimaFila === 4 ? 5 : ultimaFila + 1;
    
    // Insertar los nuevos datos
    hojaCobros.getRange(filaEscribir, 1, 1, 6).setValues([[fecha, paciente, doctor, importe, "", "Y"]]);
    
    // Aplicar formato de moneda en euros a la columna IMPORTE TOTAL (columna 4)
    hojaCobros.getRange(filaEscribir, 4).setNumberFormat("€#,##0.00");
    
    // Definir la regla de validación
    var opcionesMetodoPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    var reglaValidacion_cobro = SpreadsheetApp.newDataValidation()
        .requireValueInList(opcionesMetodoPago, true)
        .setAllowInvalid(false)
        .build();

    // Aplicar la validación a la celda de método de pago
    hojaCobros.getRange(filaEscribir, 5).setDataValidation(reglaValidacion_cobro);
}

function limpiarFormulario(hoja) {
    var celdas = ["C3", "C5", "C7","B12", "C12", "D12", "E12", "F12", "G12", "B17", "C17", "D17", "E17", "F17", "B21"];
    celdas.forEach(celda => hoja.getRange(celda).setValue(""));
}


function actualizarTablaResumen(hojaMes) {
    var totalPresupuestado = 0, totalAceptado = 0, pacientesAceptados = 0, totalPacientes = 0;

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

    // Obtener los datos de la tabla desde la fila 18 en adelante
    var datosCompletos = hojaMes.getRange("A18:Z" + hojaMes.getLastRow()).getValues();

    datosCompletos.forEach(fila => {
        var estado = fila[8] ? fila[8].toString().trim().toLowerCase() : "";
        var importePresupuestado = parseFloat(fila[9]) || 0;
        var importeAceptado = parseFloat(fila[11]) || 0;

        if (estado === "aceptado") pacientesAceptados++;
        if (importePresupuestado > 0) {
            totalPacientes++;
            totalPresupuestado += importePresupuestado;
        }
        if (estado === "aceptado") totalAceptado += importeAceptado;
    });

    var pto_medio = totalPacientes > 0 ? totalPresupuestado / totalPacientes : 0;

    // Actualizar todos los valores en una sola llamada
    var valoresResumen = [
        [totalPresupuestado.toFixed(2) + " €", totalPacientes],
        [totalAceptado.toFixed(2) + " €", pacientesAceptados],
        ["", ""], 
        [pto_medio.toFixed(2) + " €", ""]
    ];
    hojaMes.getRange(4, 3, valoresResumen.length, 2).setValues(valoresResumen);
    hojaMes.autoResizeColumns(2, 4); 

}

// Actualizar valores de calculo en parrilla por mes, el color de la fila y agregar a hoja de cobro si cambia a aceptado


function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;

    // Verificar si la edición se hizo en la columna 9 ("Estado") y no en la tabla de cobros
    if (rango.getColumn() === 9 && !hoja.getName().startsWith("Cobros")) {
        var estadoNuevo = rango.getValue();
        var fila = rango.getRow();

        // Verificar que no sea una celda de encabezado
        if (fila < 18) return;

        // Obtener el estado anterior antes del cambio
        var estadoAnterior = e.oldValue;

        // Si el estado anterior era "Pendiente" y el nuevo estado es "Aceptado", agregar a Cobros
        if (estadoAnterior === "Pendiente" || estadoAnterior === "No Aceptado"  && estadoNuevo === "Aceptado") {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var nombreMes = hoja.getName();
            var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);
        ///// variables que van a cobros
            var paciente = hoja.getRange(fila, 2).getValue(); // Columna PACIENTE
            var fecha = hoja.getRange(fila, 1).getValue(); // Columna FECHA
            var doctor = hoja.getRange(fila, 4).getValue(); // Columna DOCTOR 
            var importe = hoja.getRange(fila, 12).getValue(); // Columna IMPORTE ACEPTADO

            agregarAPacientesAceptados(hojaCobros, paciente, fecha, importe, doctor);

        }

        // Actualizar formato de la fila
        actualizarFormatoFila(hoja, fila, estadoNuevo);

        // Llamar a actualizarTablaResumen
        actualizarTablaResumen(hoja);
    }
}
