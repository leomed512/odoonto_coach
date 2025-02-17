function guardarDatosEnTabla2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFormulario = ss.getSheetByName("Registro de transacciones"); // Hoja de entrada

    if (!hojaFormulario) {
        Logger.log("❌ Error: No se encontró la hoja 'Registro de transacciones'");
        return;
    }

    // Obtener datos del formulario (ajustar rango según la cantidad de columnas)
    var datos = hojaFormulario.getRange("B3:H21").getValues();

    // Validar que el paciente tenga nombre
    if (!datos[0][1]) {
        Browser.msgBox("Error: El campo 'Paciente' es obligatorio.");
        return;
    }

    // Obtener la fecha ingresada en la celda
    var fechaIngresada = datos[2][1];

    if (!fechaIngresada || fechaIngresada == "") {
        Browser.msgBox("Error: Debes ingresar una fecha.");
        return;
    }

    // Convertir la fecha en objeto Date
    var fecha = new Date(fechaIngresada);
    
    // Obtener el nombre del mes y año de la fecha ingresada
    var opciones = { year: "numeric", month: "long" };
    var nombreMes = fecha.toLocaleDateString("es-ES", opciones); // Ejemplo: "enero de 2025"
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1); // Capitalizar

    var hojaMes = ss.getSheetByName(nombreMes);
    var filaInicio = 17; // Fila donde comienzan los encabezados

    // Si no existe la hoja del mes, crearla automáticamente
    if (!hojaMes) {
        hojaMes = ss.insertSheet(nombreMes);

        // **Agregar encabezados en la fila 17**
        var encabezados = ["FECHA DE CONTACTO", "PACIENTE", "TELÉFONO", "DOCTOR/A", "AUXILIAR", "TIPOLOGÍA PV", "SUBTIPOLOGÍA", "CAMPAÑA", "ESTADO", "IMPORTE PRESUPUESTADO", "N° PTOS", "IMPORTE ACEPTADO", "FECHA DE INICIO", "PLAN DE CITAS", "OBSERVACIONES"];
        hojaMes.getRange(filaInicio, 1, 1, encabezados.length).setValues([encabezados]);

        // **Aplicar estilos a los encabezados**
        var rangoEncabezados = hojaMes.getRange(filaInicio, 1, 1, encabezados.length);
        rangoEncabezados.setFontWeight("bold"); // Texto en negrita
        rangoEncabezados.setBackground("#1E90FF"); // Fondo azul (color azul similar al de Google Sheets)
        rangoEncabezados.setFontColor("white"); // Texto en color blanco para mayor contraste
        rangoEncabezados.setHorizontalAlignment("center"); // Centrar texto
    }

    // Determinar la fila donde escribir los datos
    var ultimaFila = hojaMes.getLastRow();
    var filaEscribir = ultimaFila < filaInicio ? filaInicio + 1 : ultimaFila + 1;

    // Formatear los datos para la tabla
    var nuevaFila = [
        fechaIngresada, // Fecha ingresada en el formulario
        datos[0][1], // Paciente
        datos[4][1], // Teléfono
        datos[9][0], // Doctor/a
        datos[9][1], // Auxiliar
        datos[9][2], // Tipología PV
        datos[9][3], // Subtipología
        datos[9][4], // Campaña
        datos[14][4], // Estado
        datos[14][0], // Importe presupuestado
        datos[14][1], // Número de presupuestos
        datos[14][2], // Importe aceptado
        datos[14][3], // Fecha de empezar
        datos[9][5], // Plan de citas
        datos[18][0], // Observaciones
    ];

    // Insertar los datos en la fila determinada
    hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length).setValues([nuevaFila]);

    hojaMes.getRange(filaEscribir, 10).setNumberFormat("#,##0.00€"); // Columna J - Importe presupuestado
    hojaMes.getRange(filaEscribir, 12).setNumberFormat("#,##0.00€"); 

    Logger.log("✅ Datos guardados en '" + nombreMes + "' correctamente.");
    Browser.msgBox("Datos guardados en '" + nombreMes + "' correctamente.");

    // Aplicar formato condicional según el estado
    var estado = datos[14][4]; // Estado
    var rangoFila = hojaMes.getRange(filaEscribir, 1, 1, nuevaFila.length); // Rango de la fila recién agregada

    if (estado == "Aceptado") {
        rangoFila.setBackground("#1c873b"); // Verde
    } else if (estado == "Pendiente") {
        rangoFila.setBackground("#fc9221"); // Naranja
    } else if (estado == "No aceptado") {
        rangoFila.setBackground("#fc4c3d"); // Rojo
    }

    var columnaEstado = 9; // La columna donde está el campo "Estado"

    // Agregar validación de datos (Dropdown) en la columna "Estado" de la nueva hoja
    var reglaValidacion = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Aceptado", "Pendiente", "No aceptado"], true)
        .setAllowInvalid(false)
        .build();
    hojaMes.getRange(filaEscribir, columnaEstado).setDataValidation(reglaValidacion);

    actualizarTablaResumen(hojaMes)

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
}

// Trigger para detectar el cambio en el valor de ESTADO en la nueva hoja creada y cambiar el color de acuerdo al valor del dropdown
function onEdit(e) {
    var hoja = e.source.getActiveSheet();
    var rango = e.range;
    
    var columnaEstado = 9; // Columna donde está el campo "Estado"

    // Si la edición no ocurre en la columna "Estado", salir
    if (rango.getColumn() !== columnaEstado) {
        return;
    }

    // Obtener el valor del estado editado
    var estado = rango.getValue();
    var fila = rango.getRow();
    var rangoFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn());

    // Aplicar colores según el estado seleccionado
    if (estado == "Aceptado") {
        rangoFila.setBackground("#1c873b"); // Verde
    } else if (estado == "Pendiente") {
        rangoFila.setBackground("#fc9221"); // Naranja
    } else if (estado == "No aceptado") {
        rangoFila.setBackground("#fc4c3d"); // Rojo
    } else {
        rangoFila.setBackground(null); // Restablecer el color si se borra el estado
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

        // Aplicar colores a los encabezados
        hojaMes.getRange("B1:C1").setBackground("#D8BFD8"); // Fondo lila claro
        hojaMes.getRange("B2:C2").setBackground("#C4AFFF")
        hojaMes.getRange("C3:E3").setBackground("#4682B4").setFontColor("#FFFFFF").setFontWeight("bold"); // Azul con letras blancas
        hojaMes.getRange("C4:C5").setBackground("#D3D3D3").setFontWeight("bold"); // Gris claro para los títulos de importes
        hojaMes.getRange("D4:D5").setBackground("#E6E6FA"); // Lila claro para importes
        hojaMes.getRange("E4:E5").setBackground("#D8BFD8"); // Lila más fuerte para N° Pacientes
        hojaMes.getRange("C6:C7").setBackground("#A9A9A9").setFontColor("#FFFFFF").setFontWeight("bold"); // Gris oscuro para "TOTAL COBRADO" y "PTO MEDIO"
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