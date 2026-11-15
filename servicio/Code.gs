/**
 * GEA - Servicio de sincronización con el fichero central.
 *
 * Recibe desde la web «GEA - Registro» las actividades, incidencias y puntos GPS de cada móvil y los guarda en la
 * Hoja de Google «GEA - Registro central» (pestañas Actividades, Incidencias y Puntos GPS), una fila por elemento.
 * Si un elemento ya está (mismo ID) se actualiza su fila: así no hay duplicados aunque se sincronice varias veces.
 * Los adjuntos de las incidencias (fotos, vídeos y audios) se guardan en la carpeta Incidencias de Drive.
 *
 * LECTURA (icono «Actualizar» de la web y la App, y calendario del Supervisor): con la CLAVE_LECTURA el servicio devuelve las
 * actividades, incidencias y puntos GPS de los días pedidos (máximo 3), de todos los empleados o de los indicados.
 * Si CLAVE_LECTURA está vacía, la lectura queda desactivada. Debe ser DISTINTA de la CLAVE y solo la conoce el supervisor.
 *
 * Instalación: en la Hoja, Extensiones → Apps Script, pegar este código, poner las claves y guardar.
 * Después: Implementar → Nueva implementación → Aplicación web → Ejecutar como «Yo» y acceso «Cualquier usuario».
 * Si ya estaba implementada: Implementar → Gestionar implementaciones → lápiz → Versión: «Nueva versión» → Implementar
 * (la dirección del servicio no cambia).
 */
const CLAVE = 'PON-AQUI-LA-CLAVE-DE-ESCRITURA';                                // la misma que se escribe en la web (Configuración)
const CLAVE_LECTURA = '';                                           // ← ELIGE una clave de lectura (distinta de la anterior); se escribe en Configuración → Clave de lectura
const CARPETA_INCIDENCIAS = '1D1MRL6GS4r3xzdicljNRFktkiV5ySvFM'; // carpeta «Incidencias» de GEA en Drive
const HOJAS = {actividades: 'Actividades', incidencias: 'Incidencias', puntos: 'Puntos GPS'};
const MAX_DIAS_LECTURA = 3;                                         // máximo de días por consulta

function doGet() {
  return salida({ok: true, servicio: 'GEA - Registro central'});
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return salida({ok: false, error: 'Datos no válidos'}); }
  // Lectura (supervisor): clave propia, distinta de la de escritura
  if (d && d.tipo === 'leer') {
    if (!CLAVE_LECTURA || d.clave !== CLAVE_LECTURA) return salida({ok: false, error: 'Clave de lectura incorrecta'});
    try { return salida(leer(d)); }
    catch (err) { return salida({ok: false, error: 'No se pudo leer: ' + err.message}); }
  }
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

// ---------- Lectura ----------
const texto = x => x instanceof Date ? x.toISOString() : String(x === null || x === undefined ? '' : x);
const numero = x => (x === '' || x === null || x === undefined || isNaN(Number(x))) ? null : Number(x);

// Filas de una hoja cuya «Fecha y hora ISO» cae en alguno de los rangos [desde, hasta). Las filas se añaden por orden de envío, así que se
// lee desde el final por lotes y se para cuando un lote entero es anterior (con 7 días de margen por envíos tardíos) al primer rango:
// la hoja puede crecer meses sin que la consulta se vuelva lenta.
function filasEnRangos(nombre, rangos, empleados) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre), n = hoja.getLastRow();
  if (n < 2) return {cab: [], filas: []};
  const ncol = hoja.getLastColumn(), cab = hoja.getRange(1, 1, 1, ncol).getValues()[0].map(String);
  const ci = cab.indexOf('Fecha y hora ISO'), ce = cab.indexOf('Empleado');
  const minimo = rangos.reduce((m, r) => r[0] < m ? r[0] : m, '9999');
  const margen = new Date(new Date(minimo).getTime() - 7 * 86400000).toISOString();
  const filas = [], LOTE = 20000;
  let fin = n;
  while (fin >= 2) {
    const ini = Math.max(2, fin - LOTE + 1), v = hoja.getRange(ini, 1, fin - ini + 1, ncol).getValues();
    let recientes = false;
    v.forEach(f => {
      const t = texto(f[ci]);
      if (t >= margen) recientes = true;
      if (rangos.some(r => t >= r[0] && t < r[1]) && (!empleados || empleados[texto(f[ce])])) filas.push(f);
    });
    if (!recientes) break;
    fin = ini - 1;
  }
  return {cab: cab, filas: filas};
}

// d.rangos: [[desde, hasta], …] en texto ISO (máximo MAX_DIAS_LECTURA); d.empleados: lista de nombres (opcional, vacío = todos).
// Devuelve actividades e incidencias como objetos y los puntos GPS en formato compacto:
// [id de la actividad, segundos desde 1970, latitud, longitud, precisión]
function leer(d) {
  const rangos = (d.rangos || []).slice(0, MAX_DIAS_LECTURA);
  if (!rangos.length) return {ok: false, error: 'Sin días que leer'};
  let empleados = null;
  if (d.empleados && d.empleados.length) { empleados = {}; d.empleados.forEach(x => { empleados[String(x)] = true; }); }
  const A = filasEnRangos(HOJAS.actividades, rangos, empleados), I = filasEnRangos(HOJAS.incidencias, rangos, empleados);
  const P = filasEnRangos(HOJAS.puntos, rangos, empleados);
  const a = n => A.cab.indexOf(n), i = n => I.cab.indexOf(n), p = n => P.cab.indexOf(n);
  const actividades = A.filas.map(f => ({
    id: texto(f[a('ID')]), nombre: texto(f[a('Empleado')]), actividad: texto(f[a('Actividad')]), estado: texto(f[a('Estado')]), iso: texto(f[a('Fecha y hora ISO')]),
    lat: numero(f[a('Latitud')]), lon: numero(f[a('Longitud')]), precision: numero(f[a('Precisión (m)')]),
    km: numero(f[a('Distancia (km)')]), ha: numero(f[a('Área (ha)')]), vel: numero(f[a('Velocidad (km/h)')]),
    pts: numero(f[a('Puntos GPS')]), pausa: numero(f[a('Min pausada')]), prueba: texto(f[a('Datos de prueba')]) === 'Sí'
  }));
  const incidencias = I.filas.map(f => {
    const nombres = texto(f[i('Archivos adjuntos')]).split('\n').filter(Boolean), enlaces = texto(f[i('Enlaces adjuntos')]).split('\n').filter(Boolean);
    return {
      id: texto(f[i('ID')]), nombre: texto(f[i('Empleado')]), tipo: texto(f[i('Tipo')]), comentario: texto(f[i('Comentario')]), iso: texto(f[i('Fecha y hora ISO')]),
      lat: numero(f[i('Latitud')]), lon: numero(f[i('Longitud')]), precision: numero(f[i('Precisión (m)')]), prueba: texto(f[i('Datos de prueba')]) === 'Sí',
      adjuntos: nombres.map((nombre, k) => {
        const m = nombre.match(/_(foto|video|audio)\d*\.[A-Za-z0-9]+$/);
        return {tipo: m ? m[1] : 'foto', nombre: nombre, url: enlaces.length === nombres.length ? enlaces[k] : ''};
      })
    };
  });
  const puntos = P.filas.map(f => [texto(f[p('ID actividad')]), Math.round(new Date(texto(f[p('Fecha y hora ISO')])).getTime() / 1000),
    numero(f[p('Latitud')]), numero(f[p('Longitud')]), numero(f[p('Precisión (m)')])]);
  return {ok: true, actividades: actividades, incidencias: incidencias, puntos: puntos, generado: new Date().toISOString()};
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
