function showDateRangeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DateRangePicker')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecciona un Rango de Fechas');
}

function formatearFecha(fechaStr) {
  var fecha = new Date(fechaStr);
  var meses = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];
  var mes = meses[fecha.getMonth()];
  var anio = fecha.getFullYear();
  return mes + " de " + anio;
}

function sonDelMismoMes(fechaStr1, fechaStr2) {
  //retorna true or false
  var fecha1 = new Date(fechaStr1);
  var fecha2 = new Date(fechaStr2);
  return fecha1.getMonth() === fecha2.getMonth() && fecha1.getFullYear() === fecha2.getFullYear();
}

function filtrarPorFecha(hojaMes, startDateStr, endDateStr) {

  // Convertir las fechas recibidas en formato string a objetos Date
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);

  var parrillaMes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaMes);
  var ultimaFila = parrillaMes.getLastRow();
  
  // Obtener los datos desde la fila 18 hasta la última fila y de la columna A a N (1-14)
  var rango = parrillaMes.getRange(11, 1, ultimaFila - 10, 14);
  var valores = rango.getValues();
  
  // Filtrar los registros dentro del rango de fechas en la columna B (columna 2 en índice base 1)
  var datosFiltrados = valores.filter(function(fila) {
    var fechaCelda = fila[1]; // La columna B (índice 1 en base 0)

    // Asegurar que la fecha es un objeto Date
    if (!(fechaCelda instanceof Date)) {
      fechaCelda = new Date(fechaCelda); // Convertir texto a Date si es necesario
    }

    // Comparación con las fechas recibidas
    return fechaCelda >= startDate && fechaCelda <= endDate;
  });

  Logger.log("Registros filtrados: " + datosFiltrados.length);
  return datosFiltrados; // Retorna los registros que cumplen la condición
}

function mesStagingPrevisiones(fecha) {
  var stagingPrevisiones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staging Previsiones");
  var ultimaFila3 = stagingPrevisiones.getLastRow();
  
  if (ultimaFila3 < 2) {
    Logger.log("No hay datos suficientes en Staging Previsiones.");
    return [];
  }
  
  var rango = stagingPrevisiones.getRange(2, 1, ultimaFila3 - 1, 12);
  var datosStagingPrev = rango.getValues();
  
  // Extraer mes y año de la fecha proporcionada
  var fechaObj = new Date(fecha);
  var mesFiltro = fechaObj.getMonth(); // 0-11 (enero es 0)
  var anoFiltro = fechaObj.getFullYear();
  
  // Ahora comprobamos la columna B (índice 1) en lugar de A (índice 0)
  var registrosFiltrados = datosStagingPrev.filter(function(fila) {
    // Asegurarse de que el valor en la celda es una fecha
    if (fila[1] instanceof Date) {
      var fechaRegistro = fila[1];
      return fechaRegistro.getMonth() === mesFiltro && 
             fechaRegistro.getFullYear() === anoFiltro;
    } else if (typeof fila[1] === 'string') {
      // Si el valor es una cadena, intentar convertirlo a fecha
      var fechaRegistro = new Date(fila[1]);
      if (!isNaN(fechaRegistro.getTime())) {
        return fechaRegistro.getMonth() === mesFiltro && 
               fechaRegistro.getFullYear() === anoFiltro;
      }
    }
    return false;
  });
  
  return registrosFiltrados;
}

function filtrarPorRangoFechasPrev(fechaInicio, fechaFin) {
  var stagingPrevisiones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staging Previsiones");
  var ultimaFila3 = stagingPrevisiones.getLastRow();
  
  if (ultimaFila3 < 2) {
    Logger.log("No hay datos suficientes en Staging Previsiones.");
    return [];
  }
  
  var rango = stagingPrevisiones.getRange(2, 1, ultimaFila3 - 1, 12);
  var datosStagingPrev = rango.getValues();
  
  // Convertir las fechas de inicio y fin a objetos Date
  var fechaInicioObj = new Date(fechaInicio);
  var fechaFinObj = new Date(fechaFin);
  
  // Resetear las horas para comparar solo las fechas
  fechaInicioObj.setHours(0, 0, 0, 0);
  fechaFinObj.setHours(23, 59, 59, 999);
  
  // Filtrar los registros que están dentro del rango de fechas
  // Las fechas están en la columna B (índice 1)
  var registrosFiltrados = datosStagingPrev.filter(function(fila) {
    var fechaRegistro;
    
    // Determinar si el valor es una fecha
    if (fila[1] instanceof Date) {
      fechaRegistro = fila[1];
    } else if (typeof fila[1] === 'string') {
      // Convertir string a fecha si es posible
      fechaRegistro = new Date(fila[1]);
      if (isNaN(fechaRegistro.getTime())) {
        return false; // No es una fecha válida
      }
    } else {
      return false; // No es una fecha
    }
    
    // Crear una copia de la fecha para no modificar el original
    var fechaComparar = new Date(fechaRegistro);
    fechaComparar.setHours(0, 0, 0, 0);
    
    // Verificar si la fecha está dentro del rango
    return fechaComparar >= fechaInicioObj && fechaComparar <= fechaFinObj;
  });
  
  return registrosFiltrados;
}

function sumarColumnaFPrev(rangoPrev) {
  var sumaTotal = 0;
  
  // Recorrer todos los registros en rangoPrev
  for (var i = 0; i < rangoPrev.length; i++) {
    // Obtener el valor de la columna F (índice 5)
    var valorF = rangoPrev[i][5];
    
    // Verificar si el valor es un número
    if (typeof valorF === 'number') {
      sumaTotal += valorF;
    } else if (typeof valorF === 'string') {
      // Intentar convertir a número si es una cadena
      var valorNumerico = parseFloat(valorF);
      if (!isNaN(valorNumerico)) {
        sumaTotal += valorNumerico;
      }
    }
  }
  
  return sumaTotal;
}

function filtrarCobrosPorRangoFechas(fechaInicio, fechaFin) {
  var stagingCobros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staging Cobros");
  var ultimaFila = stagingCobros.getLastRow();
  
  if (ultimaFila < 2) {
    Logger.log("No hay datos suficientes en Staging Cobros.");
    return [];
  }
  
  var rango = stagingCobros.getRange(2, 1, ultimaFila - 1, 7); // Ajustado a 7 columnas (A-G)
  var datosStagingCobros = rango.getValues();
  
  // Convertir las fechas de inicio y fin a objetos Date
  var fechaInicioObj = new Date(fechaInicio);
  var fechaFinObj = new Date(fechaFin);
  
  // Resetear las horas para comparar solo las fechas
  fechaInicioObj.setHours(0, 0, 0, 0);
  fechaFinObj.setHours(23, 59, 59, 999);
  
  // Filtrar los registros que están dentro del rango de fechas
  var registrosFiltrados = datosStagingCobros.filter(function(fila) {
    var fechaRegistro;
    
    // Determinar si el valor es una fecha
    if (fila[1] instanceof Date) {
      fechaRegistro = fila[1];
    } else if (typeof fila[1] === 'string') {
      // Convertir string a fecha
      fechaRegistro = new Date(fila[1]);
      if (isNaN(fechaRegistro.getTime())) {
        return false; // No es una fecha válida
      }
    } else {
      return false; // No es una fecha
    }
    
    // Crear una copia de la fecha para no modificar el original
    var fechaComparar = new Date(fechaRegistro);
    fechaComparar.setHours(0, 0, 0, 0);
    
    // Verificar si la fecha está dentro del rango
    return fechaComparar >= fechaInicioObj && fechaComparar <= fechaFinObj;
  });
  
  return registrosFiltrados;
}

function saveDateRange(startDate, endDate) {
  if (!sonDelMismoMes(startDate, endDate)) {
    Browser.msgBox("Error: El rango de fechas deben ser del mismo mes.");
    return;
  }
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Análisis de KPIs");
  sheet.getRange('D2').setValue(startDate);  // Guarda la fecha de inicio en D2
  sheet.getRange('D3').setValue(endDate);


      // Reset all calculated values to zero to avoid showing old data
  sheet.getRange("B5:C9").setValue(0);
  sheet.getRange("C15:C17").setValue(0);
  sheet.getRange("C19").setValue(0);
  sheet.getRange("B23:C26").setValue(0);
  sheet.getRange("B30:C36").setValue(0);
  sheet.getRange("C40:C43").setValue(0);
  
  var y = sheet.getRange('D2').getValue();
  var z = sheet.getRange('D3').getValue();
  var hojaMes = formatearFecha(y);

  var fechaObj = new Date(y);
  fechaObj.setMonth(fechaObj.getMonth() + 1);

  var ultimaFila2 = sheet.getLastRow();
  if (ultimaFila2 >= 49) {
    sheet.getRange(49, 1, ultimaFila2 - 48, sheet.getLastColumn()).clearContent();
  }

  var mesPrevisiones = mesStagingPrevisiones(y);
  var proxMesPrev = mesStagingPrevisiones(fechaObj);
  var rangoPrev = filtrarPorRangoFechasPrev(y, z);
  var totalPrevMes = sumarColumnaFPrev(mesPrevisiones);
  var totalPrevRango = sumarColumnaFPrev(rangoPrev);
  var totalProxMes = sumarColumnaFPrev(proxMesPrev);
  var datosCobros = filtrarCobrosPorRangoFechas(y, z);
  var totalCobros = sumarColumnaFPrev(datosCobros);
  sheet.getRange("C15").setValue(totalPrevRango);
  sheet.getRange("C16").setValue(totalPrevMes);
  sheet.getRange("C19").setValue(totalProxMes);
  sheet.getRange("C17").setValue(totalCobros);

  var tipoPago = ["70/30 o 50/50", "FINANC", "Pronto pago", "Según TTO"];
  var partes = 0;
  var finan = 0;
  var pontoPago = 0;
  var segunTTO = 0;
  tipoPago.forEach(function(tpago) { 
    rangoPrev.forEach(function(prev) { 
        if (tpago === prev[8] && tpago === "70/30 o 50/50") {
          partes += prev[5];
        } else if (tpago === prev[8] && tpago === "FINANC") {
          finan += prev[5];
        } else if (tpago === prev[8] && tpago === "Pronto pago") {
          pontoPago += prev[5];
        } else if (tpago === prev[8] && tpago === "Según TTO") {
          segunTTO += prev[5];
        }
   });
  });
  sheet.getRange("C40").setValue(partes);
  sheet.getRange("C41").setValue(finan);
  sheet.getRange("C42").setValue(pontoPago);
  sheet.getRange("C43").setValue(segunTTO);

  
  
  var valores = filtrarPorFecha(hojaMes, y, z);

  
  
  var count_0_1000 = 0;
  var suma_0_1000 = 0;
  var count_1001_3000 = 0;
  var suma_1001_3000 = 0;
  var count_3001_6000 = 0;
  var suma_3001_6000 = 0;
  var count_6001_10000 = 0;
  var suma_6001_10000 = 0;
  var count_mayor_10000 = 0;
  var suma_mayor_10000 = 0;

  var count_aceptados = 0;
  var sum_aceptados = 0;
  var count_no_aceptados = 0;
  var sum_no_aceptados = 0;
  var count_pen_con_cita = 0;
  var sum_pen_con_cita = 0;
  var count_pen_sin_cita = 0;
  var sum_pen_sin_cita = 0;

  var count_1vta = 0;
  var sum_1vta = 0;
  var count_amp = 0;
  var sum_amp = 0;
  var count_oc_llamado = 0;
  var sum_oc_llamado = 0;
  var count_oc_vino = 0;
  var sum_oc_vino = 0;
  var count_criba = 0;
  var sum_criba = 0;
  var count_rep_anio_ant = 0;
  var sum_rep_anio_ant = 0;
  var count_rep_anio_act = 0;
  var sum_rep_anio_act = 0;

  var setDrs = new Set();

  valores.forEach(function(fila) {
    var valor = fila[10]; // Accedemos al valor de la celda en la columna K
    var estado = fila[9];
    var tipologia = fila[6];
    var dr = fila[4];

    setDrs.add(dr);

    if (valor >= 0 && valor <= 1000) {
      count_0_1000++;
      suma_0_1000 += valor;
    } else if (valor >= 1001 && valor <= 3000) {
      count_1001_3000++;
      suma_1001_3000 += valor;
    } else if (valor >= 3001 && valor <= 6000) {
      count_3001_6000++;
      suma_3001_6000 += valor;
    } else if (valor > 10000) {
      count_mayor_10000++;
      suma_mayor_10000 += valor;
    } else if (valor >= 6001 && valor <= 10000) {
      count_6001_10000++;
      suma_6001_10000 += valor;
    }

    if (estado === "Aceptado") {
      count_aceptados++;
      sum_aceptados += valor;
    } else if (estado === "No aceptado") {
      count_no_aceptados++;
      sum_no_aceptados += valor;
    } else if (estado === "Pendiente con cita") {
      count_pen_con_cita++;
      sum_pen_con_cita += valor;
    } else if (estado === "Pendiente sin cita") {
      count_pen_sin_cita++;
      sum_pen_sin_cita += valor;
    }

    if (tipologia === "1VTA") {
      count_1vta++;
      sum_1vta += valor;
    } else if (tipologia === "AMP") {
      count_amp++;
      sum_amp += valor;
    } else if (tipologia === "OC") {
      if (fila[7] === "OC Llamado") {
        count_oc_llamado++;
        sum_oc_llamado  += valor;
      } else if (fila[7] === "OC Vino él") {
        count_oc_vino++;
        sum_oc_vino += valor;
      }
    } else if (tipologia === "Criba") {
      count_criba++;
      sum_criba += valor;
    } else if (tipologia === "Repesca año anterior") {
      count_rep_anio_ant++;
      sum_rep_anio_ant += valor;
    } else if (tipologia === "Repesca año actual") {
      count_rep_anio_act++;
      sum_rep_anio_act += valor;
    }
  });

  setDrs.forEach(function(drs) {
    var vta1 = 0;
    var amp = 0;
    var criba = 0;
    var llamado = 0;
    var vinoEl = 0;
    var aceptado = 0;
    var pendienteCita = 0;
    var pendienteSinCita = 0;
    var noAceptado = 0;

    var totalDrs = 0;

    var implante = 0;
    var ortodoncia = 0;
    var protesisRemovible = 0;
    var protesisFija = 0;
    var estetica = 0;
    var tarjetaSalud = 0;
    var resenia = 0;
    var formaPrepo = 0;

    var ultimaFila = sheet.getLastRow();
    var filaDestino = ultimaFila >= 49 ? ultimaFila + 1 : 49;
    valores.forEach(function(row) {
      if (drs === row[4]) {
        if (row[6] === "1VTA") {
          vta1++;
        } else if (row[6] === "AMP") {
          amp++;
        } else if (row[6] === "Criba") {
          criba++;
          if (row[7] === "Implante") {
            implante++;
          } else if (row[7] === "Ortodoncia") {
            ortodoncia++;
          } else if (row[7] === "Prótesis removible") {
            protesisRemovible++;
          } else if (row[7] === "Prótesis fija") {
            protesisFija++;
          } else if (row[7] === "Estética") {
            estetica++;
          } else if (row[7] === "Tarjeta salud") {
            tarjetaSalud++;
          } else if (row[7] === "Forma de pago") {
            formaPrepo++;
          } else if (row[7] === "Conservadora") {
            resenia++;
          }
        } else if (row[6] === "OC") {
          if (row[7] === "OC Llamado") {
            llamado++;
          } else if (row[7] === "OC Vino él") {
            vinoEl++;
          }
        }
      if (row[9] === "Aceptado") {
        aceptado++;
      } else if (row[9] === "Pendiente con cita") {
        pendienteCita++;
      } else if (row[9] === "Pendiente sin cita") {
        pendienteSinCita++;
      } else if (row[9] === "No aceptado") {
        noAceptado++;
      }
      }
      
    });
    totalDrs = vta1 + amp + criba + llamado + vinoEl;
    var datos = [drs, vta1, amp, criba, llamado, vinoEl, totalDrs, aceptado, pendienteCita, pendienteSinCita, noAceptado];
    var datos2 = [drs, "", ortodoncia, protesisRemovible, implante, estetica, protesisFija, tarjetaSalud, resenia, formaPrepo]
    sheet.getRange(filaDestino, 1, 1, datos.length).setValues([datos]);
    var columnaInicioDatos2 = datos.length + 3;
    sheet.getRange(filaDestino, columnaInicioDatos2, 1, datos2.length).setValues([datos2]);
  });
  
  sheet.getRange("B5").setValue(count_0_1000);
  sheet.getRange("C5").setValue(suma_0_1000);
  sheet.getRange("B6").setValue(count_1001_3000);
  sheet.getRange("C6").setValue(suma_1001_3000);
  sheet.getRange("B7").setValue(count_3001_6000);
  sheet.getRange("C7").setValue(suma_3001_6000);
  sheet.getRange("B8").setValue(count_6001_10000);
  sheet.getRange("C8").setValue(suma_6001_10000);
  sheet.getRange("B9").setValue(count_mayor_10000);
  sheet.getRange("C9").setValue(suma_mayor_10000);

  sheet.getRange("B23").setValue(count_aceptados);
  sheet.getRange("C23").setValue(sum_aceptados);
  sheet.getRange("B24").setValue(count_no_aceptados);
  sheet.getRange("C24").setValue(sum_no_aceptados);
  sheet.getRange("B25").setValue(count_pen_con_cita);
  sheet.getRange("C25").setValue(sum_pen_con_cita);
  sheet.getRange("B26").setValue(count_pen_sin_cita);
  sheet.getRange("C26").setValue(sum_pen_sin_cita);

  sheet.getRange("B30").setValue(count_1vta);
  sheet.getRange("C30").setValue(sum_1vta);
  sheet.getRange("B31").setValue(count_amp);
  sheet.getRange("C31").setValue(sum_amp);
  sheet.getRange("B32").setValue(count_criba);
  sheet.getRange("C32").setValue(sum_criba);
  sheet.getRange("B33").setValue(count_oc_llamado);
  sheet.getRange("C33").setValue(sum_oc_llamado);
  sheet.getRange("B34").setValue(count_oc_vino);
  sheet.getRange("C34").setValue(sum_oc_vino);
  sheet.getRange("B35").setValue(count_rep_anio_ant);
  sheet.getRange("C35").setValue(sum_rep_anio_ant);
  sheet.getRange("B36").setValue(count_rep_anio_act);
  sheet.getRange("C36").setValue(sum_rep_anio_act);
}

