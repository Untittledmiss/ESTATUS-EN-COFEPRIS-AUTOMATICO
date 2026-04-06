function obtenerDatosPorRangoSeleccionado() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tramites_sometidos");
  const rango = hoja.getActiveRange();
  const filaInicio = rango.getRow();
  const filaFin = filaInicio + rango.getNumRows() - 1;

  for (let fila = filaInicio; fila <= filaFin; fila++) {
    const idTramite = hoja.getRange(fila, 3).getValue().toString().trim(); // Columna C

    if (!idTramite) continue;

    const url = `https://tramiteselectronicos02.cofepris.gob.mx/EstadoTramite/Consulta.aspx?id=${encodeURIComponent(idTramite)}`;

    try {
      const respuesta = UrlFetchApp.fetch(url);
      const html = respuesta.getContentText();

      const tablaMatch = html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
      if (!tablaMatch) throw new Error("No se encontró tabla");

      const tablaHtml = tablaMatch[1];
      const filas = [...tablaHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)];
      if (filas.length < 2) throw new Error("No se encontró la segunda fila");

      const fila2Html = filas[1][1];
      const celdas = [...fila2Html.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)];
      if (celdas.length === 0) throw new Error("No se encontraron datos");

      for (let j = 0; j < celdas.length; j++) {
        let textoLimpio = celdas[j][1].replace(/<[^>]*>/g, '').trim();
        textoLimpio = decodeHTMLEntities(textoLimpio);
        hoja.getRange(fila, 3 + j).setValue(textoLimpio);
      }
    } catch (e) {
      hoja.getRange(fila, 3).setNote("❌ Error: " + e.message);
      Logger.log(`Error en fila ${fila}: ${e.message}`);
    }
  }

  SpreadsheetApp.getUi().alert("Actualización completada.");
}

// Decodifica entidades HTML nombradas y numéricas
function decodeHTMLEntities(texto) {
  const doc = XmlService.parse(`<r>${texto}</r>`);
  const contenidoPlano = doc.getRootElement().getText();
  return contenidoPlano;
}

function verificarCambiosEstatus() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tramites_sometidos");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  const datos = hoja.getRange(2, 3, ultimaFila - 1).getValues(); // IDs en columna C
  const nombres = hoja.getRange(2, 1, ultimaFila - 1).getValues(); // Columna A
  const estatusAnteriores = hoja.getRange(2, 8, ultimaFila - 1).getValues(); // columna H (estatus previo)
  const estatusActuales = hoja.getRange(2, 9, ultimaFila - 1).getValues(); // columna I (estatus actual)

  const correoDestino = "@gmail.com"; // Cambia aquí por tu correo

  for (let i = 0; i < datos.length; i++) {
    const id = datos[i][0];
    if (!id) continue;

    const url = "https://tramiteselectronicos02.cofepris.gob.mx/EstadoTramite/Consulta.aspx?id=" + id;
    const nombreTramite = (nombres[i][0] || "").toString().trim();

    try {
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      const html = response.getContentText();
      const estatus = extraerEstatus(html);

      if (!estatus) continue;

      const estatusAnterior = estatusActuales[i][0]; // el valor actual en I es el anterior para comparar
      // Si columna I está vacía (primera ejecución), la inicializamos
      if (!estatusAnterior) {
        hoja.getRange(i + 2, 9).setValue(estatus); // columna I inicial
        // columna H queda vacía en primera ronda
        continue;
      }

      // Si el estatus cambió respecto al anterior guardado en I:
      if (estatus !== estatusAnterior) {
        // Guardar estatus anterior (valor viejo que estaba en I) en H
        hoja.getRange(i + 2, 8).setValue(estatusAnterior);
        // Guardar nuevo estatus en I
        hoja.getRange(i + 2, 9).setValue(estatus);

        // Enviar correo
        MailApp.sendEmail({
          to: correoDestino,
          subject: `Cambio en estatus del trámite ${nombreTramite} ID ${id}`,
          body: `El estatus del trámite ${nombreTramite} con ID ${id} ha cambiado a: ${estatus}`
        });
      }
    } catch (error) {
      Logger.log("Error con ID " + id + ": " + error);
    }
  }
}


// Extrae el estatus desde el HTML de la página
function extraerEstatus(html) {
  const match = html.match(/<span id="MainContent_estatusTramiteLabel".*?>(.*?)<\/span>/i);
  if (match && match[1]) {
    // Decodificar caracteres especiales
    const texto = Utilities.newBlob(match[1]).getDataAsString("utf-8");
    return texto.replace("Estado del trámite:", "").trim();
  }
  return null;
}
