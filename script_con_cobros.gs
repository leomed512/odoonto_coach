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

    var fecha = new Date(fechaIngresada);
    var nombreMes = fecha.toLocaleDateString("es-ES", { year: "numeric", month: "long" });
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    var hojaMes = ss.getSheetByName(nombreMes) || crearHojaMes(ss, nombreMes);
    var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);

    var filaEscribir = hojaMes.getLastRow() < 17 ? 18 : hojaMes.getLastRow() + 1;

    var nuevaFila = [
        fechaIngresada,    // FECHA DE CONTACTO
        datos[0][1],       // PACIENTE
        datos[4][1],       // TELÉFONO
        datos[9][0],       // DOCTOR/A
        datos[9][1],       // AUXILIAR
        datos[9][2],       // TIPOLOGÍA PV
        datos[9][3],       // SUBTIPOLOGÍA
        datos[9][4],       // PLAN DE CITAS (Antes era Campaña)
        datos[14][2],      // ESTADO (Antes estaba en datos[14][4])
        datos[14][0],      // IMPORTE PRESUPUESTADO
        datos[14][1],      // IMPORTE ACEPTADO (Antes en datos[14][1], se movió)
        datos[14][3],      // FECHA DE INICIO (Antes en datos[14][3], posición ajustada)
        datos[18][0]       // OBSERVACIONES
    ];
    
    hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    // Aplicar formatos y validaciones
    actualizarFormatoFila(hojaMes, filaEscribir, datos[14][2]);
    hojaMes.getRange(filaEscribir, 10).setNumberFormat("€#,##0.00");
    hojaMes.getRange(filaEscribir, 11).setNumberFormat("€#,##0.00"); // Ajustado porque N° PTOS fue eliminado

    if (datos[14][2] === "Aceptado") {
        agregarAPacientesAceptados(hojaCobros, datos[0][1], datos[14][3], datos[14][1], datos[9][0]);
    }
    actualizarTablaResumen(hojaMes);
    limpiarFormulario(hojaFormulario);
    Logger.log("Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");
}

function crearHojaMes(ss, nombreMes) {
    var hojaMes = ss.insertSheet(nombreMes);
    var encabezados = [
        "FECHA DE CONTACTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", "AUXILIAR", "TIPOLOGÍA PV", "SUBTIPOLOGÍA",
        "PLAN DE CITAS", "ESTADO", "IMPORTE PRESUPUESTADO", "IMPORTE ACEPTADO", "FECHA DE INICIO", "OBSERVACIONES"
    ];

    hojaMes.getRange(17, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center");
    hojaMes.getRange(17, 1, 1, encabezados.length).createFilter();
    hojaMes.autoResizeColumns(1, hojaMes.getLastColumn());
    return hojaMes;
}



function crearHojaCobros(ss, nombreMes) {
    var hojaCobros = ss.insertSheet("Cobros " + nombreMes);
    
    // Tabla principal de cobros
    var encabezados = ["FECHA DE COBRO", "PACIENTE", "DOCTOR", "IMPORTE TOTAL", "TIPO DE PAGO", "ESTADO DEL COBRO", "TOTAL PAGADO", "FECHA FINAL"];
    hojaCobros.getRange(4, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    hojaCobros.getRange(4, 1, 1, encabezados.length).createFilter();

    //////////////////////////
    // var fechaActual = new Date();
    var titulo = "Caja de " + nombreMes ;
    hojaCobros.getRange(1, 1).setValue(titulo)
        .setFontWeight("bold")
        .setFontSize(18);
    
    // Insertar texto en la fila 3, fusionando las celdas E3:H3
    hojaCobros.getRange(3, 5, 1, 4).merge().setValue("Modifica manualmente estas columnas")
        .setFontSize(10)
        .setBackground("#00c896")
        .setFontColor("#424242") // Gris plomo oscuro
        .setHorizontalAlignment("center");
    
    // Nueva tabla de resumen
    var encabezadosResumen = ["TIPO DE PAGO", "N° PACIENTES", "TOTAL / TIPO"];
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
        var formulaSuma = `=SUMIF(E:E,"${tipo}",G:G)`;
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
    hojaCobros.getRange(filaTotales, 11).setFormula("=SUM(K5:K8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("center");
    
    // Fórmula para total de montos
    hojaCobros.getRange(filaTotales, 12).setFormula("=SUM(L5:L8)")
        .setFontWeight("bold")
        .setBackground("#424242")
        .setFontColor("white")
        .setHorizontalAlignment("right")
        .setNumberFormat("€#,##0.00");

         // Agregar fila "TOTAL COBRADO" justo debajo del total
    var filaTotalCobrado = filaTotales + 1;
    hojaCobros.getRange(filaTotalCobrado, 10).setValue("TOTAL COBRADO")
        .setFontWeight("bold")
        .setBackground("#999999")
        .setFontColor("black");

    // Dejar la columna del medio vacía
    hojaCobros.getRange(filaTotalCobrado, 11).setValue("")
        .setBackground("#999999");

    // Fórmula para total cobrado condicionalmente de la columna "TOTAL PAGADO"
    hojaCobros.getRange(filaTotalCobrado, 12).setFormula('=SUM(G:G)') 
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
    var ultimaFila = 4; // Empieza desde la fila de encabezados
    var datos = hojaCobros.getRange("A5:A").getValues();
    
    // la última fila con datos en la columna A
    for (var i = 0; i < datos.length; i++) {
        if (datos[i][0] !== "") {
            ultimaFila = i + 5; // Sumamr 5 porque empieza desde A5
        }
    }
    
    var filaEscribir = ultimaFila === 4 ? 5 : ultimaFila + 1;
    
    // Insertar los nuevos datos
    hojaCobros.getRange(filaEscribir, 1, 1, 8).setValues([[fecha, paciente, doctor, importe, "", "", "", ""]]);
    
    // Aplicar formato de moneda en euros a la columna IMPORTE TOTAL (columna 4 Y 7)
    hojaCobros.getRange(filaEscribir, 4).setNumberFormat("€#,##0.00");
    hojaCobros.getRange(filaEscribir, 7).setNumberFormat("€#,##0.00");

    var formulaEstadoCobro = `=IF(G${filaEscribir} < D${filaEscribir}, "Pendiente de pago", 
                              IF(G${filaEscribir} = D${filaEscribir}, "PAGADO", ""))`;

                                  
    hojaCobros.getRange(filaEscribir, 6).setFormula(formulaEstadoCobro);
    
    // Definir la regla de validación
    var opcionesMetodoPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
    var reglaValidacion_cobro = SpreadsheetApp.newDataValidation()
        .requireValueInList(opcionesMetodoPago, true)
        .setAllowInvalid(false)
        .build();

    // Aplicar la validación a la celda de método de pago
    hojaCobros.getRange(filaEscribir, 5).setDataValidation(reglaValidacion_cobro);

    /////////
      hojaCobros.getRange(filaEscribir, 8).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());

}


function limpiarFormulario(hoja) {
    var celdas = ["C3", "C5", "C7","B12", "C12", "D12", "E12", "F12", "G12", "B17", "C17", "D17", "E17", "F17", "B21"];
    celdas.forEach(celda => hoja.getRange(celda).setValue(""));
}
//////////////

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

    rangoTotalPresupuestado.setFormula(`=SUMIF(I${filaInicio}:I${ultimaFila}, "<>No aceptado", J${filaInicio}:J${ultimaFila})`);
    rangoTotalAceptado.setFormula(`=SUMIF(I${filaInicio}:I${ultimaFila}, "Aceptado", K${filaInicio}:K${ultimaFila})`);
    rangoTotalCobrado.setFormula(`=SUM('Cobros ${hojaMes.getName()}'!G:G)`);
    rangoPtoMedio.setFormula(`=IF(COUNTA(J${filaInicio}:J${ultimaFila})>0, C4/COUNTA(J${filaInicio}:J${ultimaFila}), 0)`);
    rangoPacientesPresupuestados.setFormula(`=COUNTIF(I${filaInicio}:I${ultimaFila}, "<>No aceptado")`);
    rangoPacientesAceptados.setFormula(`=COUNTIF(I${filaInicio}:I${ultimaFila}, "Aceptado")`);

    [rangoTotalPresupuestado, rangoTotalAceptado, rangoTotalCobrado, rangoPtoMedio].forEach(celda => {
        celda.setNumberFormat("€#,##0.00");
    });

    hojaMes.autoResizeColumns(2, 4);
}

function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;

    if (rango.getColumn() === 9 && !hoja.getName().startsWith("Cobros")) {
        var estadoNuevo = rango.getValue();
        var fila = rango.getRow();
        if (fila < 18) return;

        var estadoAnterior = e.oldValue;

        if ((estadoAnterior === "Pendiente" || estadoAnterior === "No Aceptado") && estadoNuevo === "Aceptado") {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var nombreMes = hoja.getName();
            var hojaCobros = ss.getSheetByName("Cobros " + nombreMes) || crearHojaCobros(ss, nombreMes);

            var paciente = hoja.getRange(fila, 2).getValue();
            var fecha = new Date();
            var doctor = hoja.getRange(fila, 4).getValue();
            var importe = hoja.getRange(fila, 11).getValue(); 

            agregarAPacientesAceptados(hojaCobros, paciente, fecha, importe, doctor);
        }

        actualizarFormatoFila(hoja, fila, estadoNuevo);
    }
}

