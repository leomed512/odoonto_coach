function guardarDatosEnTabla2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFormulario = ss.getSheetByName("Registro de transacciones");

    if (!hojaFormulario) {
        Logger.log("❌ Error: No se encontró la hoja 'Registro de transacciones'");
        return;
    }

    var datos = hojaFormulario.getRange("B3:H21").getValues();

    if (!datos[0][1]) {
        Browser.msgBox("Error: El campo 'Paciente' es obligatorio.");
        return;
    }

    var fechaIngresada = datos[2][1];

    if (!fechaIngresada || fechaIngresada == "") {
        Browser.msgBox("Error: Debes ingresar una fecha.");
        return;
    }

    var fecha = new Date(fechaIngresada);
    var opciones = { year: "numeric", month: "long" };
    var nombreMes = fecha.toLocaleDateString("es-ES", opciones);
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    var hojaMes = ss.getSheetByName(nombreMes);
    var filaInicio = 17; // Fila donde comienzan los encabezados

    ///////////////////////////////////////////////////////////COBRO
    var hojaCobros = ss.getSheetByName("Cobros " + nombreMes);
    var filaInicio_cobro = 4; // Fila donde comienzan los encabezados de COBRO
    ///////////////////////////////////////////////////////////

    if (!hojaMes) {
        hojaMes = ss.insertSheet(nombreMes);

        var encabezados = ["FECHA DE CONTACTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", "AUXILIAR", "TIPOLOGÍA PV", "SUBTIPOLOGÍA", "CAMPAÑA", "ESTADO", "IMPORTE PRESUPUESTADO", "N° PTOS", "IMPORTE ACEPTADO", "FECHA DE INICIO", "PLAN DE CITAS", "OBSERVACIONES"];
        hojaMes.getRange(filaInicio, 1, 1, encabezados.length).setValues([encabezados]);

        var rangoEncabezados = hojaMes.getRange(filaInicio, 1, 1, encabezados.length);
        rangoEncabezados.setFontWeight("bold").setBackground("#424242").setFontColor("white").setHorizontalAlignment("center");

        //  Activar filtros en los encabezados
        hojaMes.getRange(filaInicio, 1, 1, encabezados.length).createFilter();
    }

    //////////////////////////////////COBRO
        // Si no existe la hoja de "Cobros [Mes Año]", crearla automáticamente
      if (!hojaCobros) {
        hojaCobros = ss.insertSheet("Cobros " + nombreMes);

        // **Agregar encabezados en la hoja de Cobros**
        var encabezadosCobros = ["FECHA DE COBRO", "PACIENTE", "IMPORTE COBRADO", "MÉTODO DE PAGO", "ESTADO DEL COBRO", "OBSERVACIONES"];
        hojaCobros.getRange(filaInicio_cobro, 1, 1, encabezadosCobros.length).setValues([encabezadosCobros]);
        //hojaCobros.appendRow(encabezadosCobros);

        // **Aplicar estilos a los encabezados**
        var rangoEncabezadosCobros = hojaCobros.getRange(filaInicio_cobro, 1, 1, encabezadosCobros.length);
        rangoEncabezadosCobros.setFontWeight("bold");
        rangoEncabezadosCobros.setBackground("#424242");
        rangoEncabezadosCobros.setFontColor("white");
        rangoEncabezadosCobros.setHorizontalAlignment("center");


    }
    ////////////////////COBROS
    actualizarFiltroTablaCobros(hojaCobros, filaInicio_cobro);
    hojaCobros.autoResizeColumns(1, hojaCobros.getLastColumn());
    ////////////////////////////////////////////////

    var ultimaFila = hojaMes.getLastRow();
    var filaEscribir = (ultimaFila < filaInicio) ? filaInicio + 1 : ultimaFila + 1;

    //  Si la hoja está completamente vacía, asegurarnos de que empiece en la fila 18
    if (ultimaFila === 0) {
        filaEscribir = filaInicio + 1;
    }

    var nuevaFila = [
        fechaIngresada, datos[0][1], datos[4][1], datos[9][0], datos[9][1], datos[9][2], datos[9][3], datos[9][4], datos[14][4], 
        datos[14][0], datos[14][1], datos[14][2], datos[14][3], datos[9][5], datos[18][0]
    ];

    hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    // Aplicar formato de tabla y filtros
    actualizarFiltroTabla(hojaMes, filaInicio);

    //  Ajustar ancho de columnas automáticamente
    hojaMes.autoResizeColumns(1, hojaMes.getLastColumn());

    hojaMes.getRange(filaEscribir, 10).setNumberFormat("#,##0.00€");
    hojaMes.getRange(filaEscribir, 12).setNumberFormat("#,##0.00€");

    var estado = datos[14][4];
    var rangoFila = hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length);

    if (estado == "Aceptado") {
        rangoFila.setBackground("#54c772");
    } else if (estado == "Pendiente") {
        rangoFila.setBackground("#FF9D23");
    } else if (estado == "No aceptado") {
        rangoFila.setBackground("#fc4c3d");
    }

    var reglaValidacion = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Aceptado", "Pendiente", "No aceptado"], true)
        .setAllowInvalid(false)
        .build();
    hojaMes.getRange(filaEscribir, 9).setDataValidation(reglaValidacion);

    actualizarTablaResumen(hojaMes);

    ////////////////////////COBROS ADD
        // Si el estado es "Aceptado", agregar a hojaCobros
    if (datos[14][4] === "Aceptado") {
        agregarAPacientesAceptados(hojaCobros, datos[0][1], fechaIngresada, datos[14][2], datos[18][0]);
    }

// Limpiar el formulario después de guardar
    hojaFormulario.getRange(5, 3).setValue("");  // Fecha ingresada (C3)
    hojaFormulario.getRange(3, 3).setValue("");  // Paciente (B1)
    hojaFormulario.getRange(7, 3).setValue("");  // Teléfono (B5)
    hojaFormulario.getRange(12, 2).setValue(""); // Doctor/a (A10)
    hojaFormulario.getRange(12, 3).setValue(""); // Auxiliar (B10)
    hojaFormulario.getRange(12, 4).setValue(""); // Tipología PV (C10)
    hojaFormulario.getRange(12, 5).setValue(""); // Subtipología (D10)
    hojaFormulario.getRange(12, 6).setValue(""); // Campaña (E10)
    hojaFormulario.getRange(17, 6).setValue(""); // Estado (E15)
    hojaFormulario.getRange(17, 2).setValue(""); // Importe presupuestado (A15)
    hojaFormulario.getRange(17, 3).setValue(""); // Número de presupuestos (B15)
    hojaFormulario.getRange(17, 4).setValue(""); // Importe aceptado (C15)
    hojaFormulario.getRange(17, 5).setValue(""); // Fecha de empezar (D15)
    hojaFormulario.getRange(12, 7).setValue(""); // Plan de citas (F10)
    hojaFormulario.getRange(21, 2).setValue(""); // Observaciones (A19)

    Logger.log("✅ Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");
}

// /**
// Función para simular la función de filtros de las tablas de excel
//  */
function actualizarFiltroTabla(hojaMes, filaInicio) {
    var rangoFiltro = hojaMes.getRange(filaInicio, 1, hojaMes.getLastRow(), hojaMes.getLastColumn());
    if (hojaMes.getFilter()) {
        hojaMes.getFilter().remove(); // Eliminar filtro antiguo
    }
    rangoFiltro.createFilter(); // Crear nuevo filtro que abarque todas las filas
}


function actualizarFiltroTablaCobros(hojaCobros, filaInicio_cobro) {
    var ultimaFilaCobro = hojaCobros.getLastRow();
    var ultimaColumnaCobro = hojaCobros.getLastColumn();

    // Asegurar que al menos haya espacio para aplicar el filtro
    var filasDisponibles = hojaCobros.getMaxRows();
    if (ultimaFilaCobro < filaInicio_cobro + 1) {
        ultimaFilaCobro = filaInicio_cobro + 1; // Asegurar que al menos haya una fila después de los encabezados
    }

    var rangoFiltroCobros = hojaCobros.getRange(filaInicio_cobro, 1, filasDisponibles - filaInicio_cobro + 1, ultimaColumnaCobro);

    // Eliminar filtro antiguo si existe
    if (hojaCobros.getFilter()) {
        hojaCobros.getFilter().remove();
    }

    // Aplicar el filtro en la tabla de cobros
    rangoFiltroCobros.createFilter();
}

//  Función que agrega un paciente a la hoja de cobros si su estado es "Aceptado".

function agregarAPacientesAceptados(hojaCobros, paciente, fechaContacto, importeAceptado, observaciones) {
    var ultimaFilaCobros = hojaCobros.getLastRow();
    var filaEscribirCobro = ultimaFilaCobros < 4 ? 5 : ultimaFilaCobros + 1;

    var nuevaFilaCobro = [paciente, fechaContacto, importeAceptado, "Pendiente", observaciones];
    hojaCobros.getRange(filaEscribirCobro, 1, 1, nuevaFilaCobro.length).setValues([nuevaFilaCobro]);
}




function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;
    var columnaEstado = 9; // La columna "Estado"

    // Verificar que el cambio ocurre en la columna "Estado" y que la hoja no sea de cobros
    if (!hoja.getName().startsWith("Cobros") && rango.getColumn() === columnaEstado) {
        var estadoNuevo = rango.getValue();
        var estadoPrevio = e.oldValue; // Captura el estado anterior
        var fila = rango.getRow(); // Obtener la fila editada
        var ultimaColumna = hoja.getLastColumn(); // Obtener la última columna con datos
        var rangoFila = hoja.getRange(fila, 1, 1, ultimaColumna); // Selecciona toda la fila

        // **Actualizar el color de la fila según el estado nuevo**
        if (estadoNuevo == "Aceptado") {
            rangoFila.setBackground("#54c772"); // Verde
        } else if (estadoNuevo == "Pendiente") {
            rangoFila.setBackground("#fc9221"); // Naranja
        } else if (estadoNuevo == "No aceptado") {
            rangoFila.setBackground("#fc4c3d"); // Rojo
        } else {
            rangoFila.setBackground(null); // Restaurar color si no es un valor válido
        }

        // **Si el estado cambia a "Aceptado" desde cualquier otro estado, agregar a hojaCobros**
        if (estadoPrevio !== "Aceptado" && estadoNuevo === "Aceptado") {
            var paciente = hoja.getRange(fila, 2).getValue(); // Columna "Paciente"
            var fechaContacto = hoja.getRange(fila, 1).getValue(); // Columna "Fecha de Contacto"
            var importeAceptado = hoja.getRange(fila, 12).getValue(); // Columna "Importe Aceptado"
            var observaciones = hoja.getRange(fila, 15).getValue(); // Columna "Observaciones"

            var nombreMes = hoja.getName();
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var hojaCobros = ss.getSheetByName("Cobros " + nombreMes);

            if (hojaCobros) {
                agregarAPacientesAceptados(hojaCobros, paciente, fechaContacto, importeAceptado, observaciones);
            }
        }
    }
}

// limpiar datos del formulario si se desea cancelar la operación antes de guardar
function borrarDatosRegistro(){
      var ss_del = SpreadsheetApp.getActiveSpreadsheet();
      var hojaFormulario_del = ss_del.getSheetByName("Registro de transacciones"); // Hoja de entrada
         hojaFormulario_del.getRange(5, 3).setValue("");  // Fecha ingresada (C3)
    hojaFormulario_del.getRange(3, 3).setValue("");  // Paciente (B1)
    hojaFormulario_del.getRange(7, 3).setValue("");  // Teléfono (B5)
    hojaFormulario_del.getRange(12, 2).setValue(""); // Doctor/a (A10)
    hojaFormulario_del.getRange(12, 3).setValue(""); // Auxiliar (B10)
    hojaFormulario_del.getRange(12, 4).setValue(""); // Tipología PV (C10)
    hojaFormulario_del.getRange(12, 5).setValue(""); // Subtipología (D10)
    hojaFormulario_del.getRange(12, 6).setValue(""); // Campaña (E10)
    hojaFormulario_del.getRange(17, 6).setValue(""); // Estado (E15)
    hojaFormulario_del.getRange(17, 2).setValue(""); // Importe presupuestado (A15)
    hojaFormulario_del.getRange(17, 3).setValue(""); // Número de presupuestos (B15)
    hojaFormulario_del.getRange(17, 4).setValue(""); // Importe aceptado (C15)
    hojaFormulario_del.getRange(17, 5).setValue(""); // Fecha de empezar (D15)
    hojaFormulario_del.getRange(12, 7).setValue(""); // Plan de citas (F10)
    hojaFormulario_del.getRange(21, 2).setValue(""); // Observaciones (A19)
      Browser.msgBox("Datos borrados del formulario correctamente");
}

function actualizarTablaResumen(hojaMes) {
    var totalPresupuestado = 0;
    var totalAceptado = 0;
    var pacientesAceptados = 0;
    var totalPacientes = 0;
    // Verificar si la tabla ya existe en la hoja
    var celdaCheck = hojaMes.getRange("C4").getValue();
    var tablaExiste = (celdaCheck && celdaCheck.toString().trim().toUpperCase() === "TOTAL PRESUPUESTADO");

    if (!tablaExiste) {
        // Definir la estructura de la tabla con exactamente 3 columnas en todas las filas
        var resumenEncabezados = [
            ["ENVIAR SEMANALMENTE", "", ""],  // B1
            ["Gerencia@odontologycoach.cr", "", ""],  // B2
            ["", "IMPORTES", "N° PACIENTES"],  // C3:E3
            ["TOTAL PRESUPUESTADO", "", ""],  // C4:E4
            ["TOTAL ACEPTADO", "", ""],  // C5:E5
            ["TOTAL COBRADO", "", ""],  // C6:E6
            ["PTO MEDIO", "", ""]   // C7:E7
        ];

        // Insertar la tabla en las posiciones correctas
        hojaMes.getRange(1, 2, resumenEncabezados.length, 3).setValues(resumenEncabezados);
        hojaMes.autoResizeColumns(2,3); // Ajustar tamaño de columnas al texto

        // Aplicar colores a los encabezados
        hojaMes.getRange("B1:C1").setBackground("#00c896").setFontWeight("bold"); // Fondo lila claro
        hojaMes.getRange("B2:C2").setBackground("#f2ecff")
        hojaMes.getRange("C3:D3").setBackground("#424242").setFontColor("#FFFFFF").setFontWeight("bold"); // Gris con letras blancas
        /////
        hojaMes.getRange("B4").setBackground("#e2e2e2").setFontWeight("bold"); 
        hojaMes.getRange("B5").setBackground("#f6f6f6").setFontWeight("bold"); 
        hojaMes.getRange("B6").setBackground("#e2e2e2").setFontWeight("bold"); 
        hojaMes.getRange("B7").setBackground("#f6f6f6").setFontWeight("bold"); 

        hojaMes.getRange("C4").setBackground("#f6f6f6").setFontWeight("bold"); 
        hojaMes.getRange("C5").setBackground("#e2e2e2").setFontWeight("bold"); 
        hojaMes.getRange("C6").setBackground("#f6f6f6").setFontWeight("bold");
        hojaMes.getRange("C7").setBackground("#e2e2e2").setFontWeight("bold"); 
        ///////////////
        hojaMes.getRange("D4").setBackground("#e2e2e2"); 
        hojaMes.getRange("D5").setBackground("#f6f6f6"); 
        hojaMes.getRange("D6").setBackground("#e2e2e2"); 
    }

    // Obtener los datos de la tabla principal (desde fila 17 en adelante)
    var datosCompletos = hojaMes.getRange("A18:Z").getValues();
    
    for (var i = 0; i < datosCompletos.length; i++) { 
        var estado = datosCompletos[i][8] ? datosCompletos[i][8].toString().trim().toLowerCase() : ""; // Estado en columna "I"

        if (estado === "aceptado") {
            pacientesAceptados++;
        }
        if (datosCompletos[i][9]) { // Importe Presupuestado en columna "J"
            totalPacientes++;
        }
        if (datosCompletos[i][9]) { // Importe Presupuestado en columna "J"
            totalPresupuestado += parseFloat(datosCompletos[i][9]) || 0;
        }
        if (datosCompletos[i][11] && (estado === "aceptado") ){ // Importe Aceptado en columna "L"
            totalAceptado += parseFloat(datosCompletos[i][11]) || 0;
        }
    }

    // ✅ Actualizar la tabla resumen con valores en tiempo real
    hojaMes.getRange(4, 3).setValue(totalPresupuestado.toFixed(2) + " €");
    hojaMes.getRange(4, 4).setValue(totalPacientes).setNumberFormat("0");

    hojaMes.getRange(5, 3).setValue(totalAceptado.toFixed(2) + " €");
    hojaMes.getRange(5, 4).setValue(pacientesAceptados).setNumberFormat("0");
    pto_medio = totalPresupuestado/totalPacientes
    hojaMes.getRange(7, 3).setValue(pto_medio.toFixed(2) + " €");
    

}

