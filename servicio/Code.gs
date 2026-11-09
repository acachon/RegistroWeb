/**
 * GEA - Servicio de sincronización con el fichero central.
 *
 * Recibe desde la web «GEA - Registro» las actividades, incidencias y puntos GPS de cada móvil y los guarda en la
 * Hoja de Google «GEA - Registro central» (pestañas Actividades, Incidencias y Puntos GPS), una fila por elemento.
 * Si un elemento ya está (mismo ID) se actualiza su fila: así no hay duplicados aunque se sincronice varias veces.
 * Los adjuntos de las incidencias (fotos, vídeos y audios) se guardan en la carpeta Incidencias de Drive.
 *
 * Instalación: en la Hoja, Extensiones → Apps Script, pegar este código, poner la CLAVE y guardar.
 * Después: Implementar → Nueva implementación → Aplicación web → Ejecutar como «Yo» y acceso «Cualquier usuario».
 */
const CLAVE = 'PON_AQUI_LA_CLAVE';                                // la misma que se escribe en la web (Configuración)
const CARPETA_INCIDENCIAS = '1D1MRL6GS4r3xzdicljNRFktkiV5ySvFM'; // carpeta «Incidencias» de GEA en Drive
const HOJAS = {actividades: 'Actividades', incidencias: 'Incidencias', puntos: 'Puntos GPS'};

function doGet() {
  return salida({ok: true, servicio: 'GEA - Registro central'});
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return salida({ok: false, error: 'Datos no válidos'}); }
  if (!d || d.clave !== CLAVE) return salida({ok: false, error: 'Clave incorrecta'});
  if (d.tipo === 'adjunto') return salida(guardarAdjunto(d));
  // Una sincronización a la vez: si dos móviles envían juntos, el segundo espera
  const cerrojo = LockService.getScriptLock();
  if (!cerrojo.tryLock(30000)) return salida({ok: false, error: 'Fichero ocupado, inténtalo de nuevo'});
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet(), r = {ok: true};
    Object.keys(HOJAS).forEach(k => { r[k] = volcar(libro.getSheetByName(HOJAS[k]), d[k] || []); });
    return salida(r);
  } finally {
    cerrojo.releaseLock();
  }
}

// Escribe las filas (objetos {cabecera: valor}) en la hoja: actualiza las que ya están por ID y añade las nuevas al final.
// Las columnas se buscan por su nombre en la fila 1, así que se pueden reordenar o añadir columnas propias.
function volcar(hoja, filas) {
  if (!filas.length) return {nuevas: 0, actualizadas: 0, ids: []};
  const ncol = hoja.getLastColumn(), cab = hoja.getRange(1, 1, 1, ncol).getValues()[0].map(String);
  const ultima = hoja.getLastRow(), filaDe = {};
  if (ultima > 1) hoja.getRange(2, 1, ultima - 1, 1).getValues().forEach((v, i) => { filaDe[String(v[0])] = i + 2; });
  const ahora = new Date(), nuevas = [];
  let actualizadas = 0;
  filas.forEach(f => {
    const fila = cab.map(c => c === 'Sincronizado' ? ahora : convertir(c, f[c]));
    if (filaDe[f.ID] === -1) return; // repetida en el mismo envío
    if (filaDe[f.ID]) { hoja.getRange(filaDe[f.ID], 1, 1, ncol).setValues([fila]); actualizadas++; }
    else { nuevas.push(fila); filaDe[f.ID] = -1; }
  });
  if (nuevas.length) {
    const desde = hoja.getLastRow() + 1, faltan = desde + nuevas.length - 1 - hoja.getMaxRows();
    if (faltan > 0) hoja.insertRowsAfter(hoja.getMaxRows(), faltan + 100);
    hoja.getRange(desde, 1, nuevas.length, ncol).setValues(nuevas);
    cab.forEach((c, i) => { if (FORMATOS[c]) hoja.getRange(desde, i + 1, nuevas.length, 1).setNumberFormat(FORMATOS[c]); });
  }
  return {nuevas: nuevas.length, actualizadas: actualizadas, ids: filas.map(f => f.ID)};
}

// Fechas y horas como números de fecha/hora de la hoja (días desde 30/12/1899), con su formato: se pueden filtrar,
// agrupar y sumar en tablas dinámicas y no dependen de la zona horaria del script
const FORMATOS = {'Fecha': 'dd/mm/yyyy', 'Hora': 'hh:mm', 'Duración': '[h]:mm', 'Sincronizado': 'dd/mm/yyyy hh:mm'};
function convertir(cabecera, v) {
  if (v === undefined || v === null) return '';
  if (cabecera === 'Fecha' && /^\d{4}-\d{2}-\d{2}$/.test(v)) { const p = v.split('-').map(Number); return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000; }
  if (cabecera === 'Hora' && /^\d{2}:\d{2}$/.test(v)) { const p = v.split(':').map(Number); return (p[0] * 60 + p[1]) / 1440; }
  if (cabecera === 'Duración' && typeof v === 'number') return v / 1440; // minutos → días (formato [h]:mm)
  // Un texto que empiece por = + - @ se guarda como texto, nunca como fórmula
  if (typeof v === 'string' && /^[=+\-@]/.test(v) && !/^-?\d/.test(v)) return "'" + v;
  return v;
}

// Guarda un adjunto en la carpeta Incidencias con su nombre (fecha_hora_empleado_tipo_adjunto). Si ya existe, no lo duplica.
function guardarAdjunto(d) {
  try {
    const carpeta = DriveApp.getFolderById(CARPETA_INCIDENCIAS), ya = carpeta.getFilesByName(d.nombre);
    if (ya.hasNext()) return {ok: true, url: ya.next().getUrl(), existia: true};
    const archivo = carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(d.datos), d.mime || 'application/octet-stream', d.nombre));
    return {ok: true, url: archivo.getUrl()};
  } catch (err) {
    return {ok: false, error: 'No se pudo guardar el adjunto: ' + err.message};
  }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
