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
            rangoFila.setBackground("#1c873b"); // Verde
        } else if (estadoNuevo == "Pendiente") {
            rangoFila.setBackground("#fc9221"); // Naranja
        } else if (estadoNuevo == "No aceptado") {
            rangoFila.setBackground("#fc4c3d"); // Rojo
        } else {
            rangoFila.setBackground(null); // Restaurar color si no es un valor válido
        }

        var paciente = hoja.getRange(fila, 2).getValue(); // Columna "Paciente"
        var fechaContacto = hoja.getRange(fila, 1).getValue(); // Columna "Fecha de Contacto"
        var importeAceptado = hoja.getRange(fila, 12).getValue(); // Columna "Importe Aceptado"
        var observaciones = hoja.getRange(fila, 15).getValue(); // Columna "Observaciones"

        var nombreMes = hoja.getName();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var hojaCobros = ss.getSheetByName("Cobros " + nombreMes);

        // **Si el estado cambia a "Aceptado" desde cualquier otro estado, agregar a hojaCobros**
        if (estadoPrevio !== "Aceptado" && estadoNuevo === "Aceptado") {
            if (hojaCobros) {
                agregarAPacientesAceptados(hojaCobros, paciente, fechaContacto, importeAceptado, observaciones);
            }
        }

        // **Si el estado cambia de "Aceptado" a cualquier otro estado, eliminar de hojaCobros**
        if (estadoPrevio === "Aceptado" && estadoNuevo !== "Aceptado") {
            if (hojaCobros) {
                eliminarPacienteDeCobros(hojaCobros, paciente, fechaContacto, importeAceptado);
            }
        }
    }
}

/**
 * Función que agrega un paciente a la hoja de cobros si su estado es "Aceptado",
 * evitando duplicados (Paciente + Fecha + Monto).
 */
function agregarAPacientesAceptados(hojaCobros, paciente, fechaContacto, importeAceptado, observaciones) {
    var ultimaFilaCobros = hojaCobros.getLastRow();
    var datosCobros = hojaCobros.getRange(5, 1, ultimaFilaCobros - 4, 3).getValues(); // Obtener datos actuales

    // Verificar si el paciente ya existe con la misma fecha y monto
    for (var i = 0; i < datosCobros.length; i++) {
        if (datosCobros[i][0] === paciente && datosCobros[i][1] === fechaContacto && datosCobros[i][2] === importeAceptado) {
            Logger.log("⛔ El paciente ya está registrado en Cobros con la misma fecha y monto. No se agregará.");
            return;
        }
    }

    var filaEscribirCobro = ultimaFilaCobros < 4 ? 5 : ultimaFilaCobros + 1;
    var nuevaFilaCobro = [paciente, fechaContacto, importeAceptado, "Pendiente", observaciones];
    hojaCobros.getRange(filaEscribirCobro, 1, 1, nuevaFilaCobro.length).setValues([nuevaFilaCobro]);
}

/**
 * Función que elimina un paciente de la hoja de cobros si su estado cambia de "Aceptado" a otro estado.
 */
function eliminarPacienteDeCobros(hojaCobros, paciente, fechaContacto, importeAceptado) {
    var ultimaFilaCobros = hojaCobros.getLastRow();
    var datosCobros = hojaCobros.getRange(5, 1, ultimaFilaCobros - 4, 3).getValues();

    for (var i = 0; i < datosCobros.length; i++) {
        if (datosCobros[i][0] === paciente && datosCobros[i][1] === fechaContacto && datosCobros[i][2] === importeAceptado) {
            hojaCobros.deleteRow(i + 5); // Eliminar fila (sumar 5 porque empieza en la fila 5)
            Logger.log("❌ Paciente eliminado de Cobros: " + paciente);
            return;
        }
    }
}



TODAVIA NO ELIMINA SI HAY DUPLICADO
No agrega a hoja cobro los nuevos registros de marzo, solo febrero
No aplica formato condicional ni data validation a marzo, solo febrero
