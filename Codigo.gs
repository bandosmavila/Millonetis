/**
 * LOTERÍA DE AMIGOS — Backend (Google Apps Script)
 * ==================================================
 * Patrón: un único doGet(e) que despacha por e.parameter.accion,
 * lee/escribe en el Spreadsheet activo y devuelve siempre JSON.
 *
 * IMPORTANTE ANTES DE DESPLEGAR:
 * 1) Cambia CLAVE por una contraseña larga y aleatoria SOLO TUYA.
 *    Esa misma CLAVE debe copiarse luego en index.html y tabla.html.
 * 2) Este proyecto de Apps Script debe estar vinculado a una hoja de
 *    cálculo NUEVA, sin ninguna relación con la de BandosMávila.
 */

var CLAVE = 'mLn7-Kx92QwPzR4vT8sB1cYh6fJd0Nea';
var JORNADA_POR_DEFECTO = 'Navidad2026';
var HOJA_APORTACIONES = 'Aportaciones';
var PRECIO_DECIMO = 20;

// API NO OFICIAL (no es de Loterías y Apuestas del Estado) usada para
// comprobar automáticamente los premios de la Lotería de Navidad. Lleva
// años funcionando de forma extraoficial, pero podría cambiar o dejar de
// funcionar sin aviso. Solo sirve para el sorteo de Navidad (números de
// 5 cifras), no para otros sorteos.
var API_COMPROBADOR_NAVIDAD = 'http://api.elpais.com/ws/LoteriaNavidadPremiados';

function doGet(e) {
  var accion = e.parameter.accion;
  var resultado;

  try {
    switch (accion) {
      case 'aportar':
        resultado = aportarDecimo(e.parameter);
        break;
      case 'listar':
        resultado = listarAportaciones(e.parameter);
        break;
      case 'premio':
        resultado = registrarPremio(e.parameter);
        break;
      case 'comprobar':
        resultado = comprobarPremios(e.parameter);
        break;
      default:
        resultado = { ok: false, error: 'Acción no reconocida' };
    }
  } catch (err) {
    resultado = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function comprobarClave(clave) {
  return clave === CLAVE;
}

function getHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_APORTACIONES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_APORTACIONES);
    hoja.appendRow(['Jornada', 'Nombre', 'Numero', 'Precio', 'Recargo', 'Premio', 'Timestamp']);
  }
  return hoja;
}

/**
 * Registra el décimo de una persona para una jornada.
 * La comprobación de "¿ya ha aportado esta persona en esta jornada?"
 * se hace aquí, releyendo la hoja en el momento de guardar (nunca solo
 * en el navegador), para evitar duplicados aunque dos personas envíen
 * el formulario casi a la vez.
 */
function aportarDecimo(p) {
  if (!comprobarClave(p.clave)) return { ok: false, error: 'Clave incorrecta' };

  var nombre = (p.nombre || '').trim();
  var numero = (p.numero || '').trim();
  var recargo = parseFloat(p.recargo || '0') || 0;
  var jornada = p.jornada || JORNADA_POR_DEFECTO;

  if (!nombre || !numero) {
    return { ok: false, error: 'Faltan datos (nombre o número)' };
  }

  var hoja = getHoja();
  var datos = hoja.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === jornada && datos[i][1] === nombre) {
      return { ok: false, error: 'Ya has registrado tu décimo para esta jornada' };
    }
  }

  hoja.appendRow([jornada, nombre, numero, PRECIO_DECIMO, recargo, '', new Date()]);
  return { ok: true };
}

/**
 * Devuelve todas las aportaciones de una jornada, más el total de premio.
 */
function listarAportaciones(p) {
  var jornada = (p.jornada || JORNADA_POR_DEFECTO);
  var hoja = getHoja();
  var datos = hoja.getDataRange().getValues();
  var filas = [];
  var totalPremio = 0;

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === jornada) {
      var premioCelda = datos[i][5];
      var premio = parseFloat(premioCelda) || 0;
      if (premioCelda !== '') totalPremio += premio;
      filas.push({
        nombre: datos[i][1],
        numero: datos[i][2],
        precio: datos[i][3],
        recargo: datos[i][4],
        premio: premioCelda === '' ? null : premio
      });
    }
  }

  return { ok: true, jornada: jornada, aportaciones: filas, totalPremio: totalPremio };
}

/**
 * (Solo gestor, protegido por CLAVE) Anota el premio que ha tocado a
 * un número concreto de una jornada.
 */
function registrarPremio(p) {
  if (!comprobarClave(p.clave)) return { ok: false, error: 'Clave incorrecta' };

  var jornada = p.jornada || JORNADA_POR_DEFECTO;
  var numero = (p.numero || '').trim();
  var premio = parseFloat(p.premio || '0') || 0;

  var hoja = getHoja();
  var datos = hoja.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === jornada && String(datos[i][2]) === numero) {
      hoja.getRange(i + 1, 6).setValue(premio); // columna F = Premio
      return { ok: true };
    }
  }

  return { ok: false, error: 'No se encontró ese número en esta jornada' };
}

/**
 * (Solo gestor, protegido por CLAVE) Consulta automáticamente el premio de
 * cada número registrado en una jornada, usando una API NO OFICIAL (ver
 * API_COMPROBADOR_NAVIDAD arriba), y lo guarda directamente en la hoja.
 *
 * El premio que se guarda es el que corresponde a un décimo de 20€, tal
 * cual lo da la API — no tiene en cuenta el recargo. El reparto (devolver
 * antes el recargo a quien pagó de más, y repartir el resto entre los 9)
 * es un cálculo aparte que se hace a mano.
 */
function comprobarPremios(p) {
  if (!comprobarClave(p.clave)) return { ok: false, error: 'Clave incorrecta' };

  var jornada = p.jornada || JORNADA_POR_DEFECTO;
  var hoja = getHoja();
  var datos = hoja.getDataRange().getValues();
  var cachePremios = {};
  var estadoSorteo = null;
  var resultados = [];

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] !== jornada) continue;

    var nombre = datos[i][1];
    var numero = String(datos[i][2]);

    if (!(numero in cachePremios)) {
      cachePremios[numero] = consultarPremioApi(numero);
      if (cachePremios[numero].status !== null) {
        estadoSorteo = cachePremios[numero].status;
      }
    }

    var info = cachePremios[numero];
    if (info.premio !== null) {
      hoja.getRange(i + 1, 6).setValue(info.premio); // columna F = Premio
    }

    resultados.push({
      numero: numero,
      nombre: nombre,
      premio: info.premio,
      error: info.error
    });
  }

  return { ok: true, jornada: jornada, estadoSorteo: estadoSorteo, resultados: resultados };
}

function consultarPremioApi(numero) {
  try {
    var url = API_COMPROBADOR_NAVIDAD + '?s=1&n=' + encodeURIComponent(numero);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    return {
      premio: parseFloat(json.premio) || 0,
      status: (typeof json.status === 'number') ? json.status : null,
      error: null
    };
  } catch (e) {
    return { premio: null, status: null, error: e.message };
  }
}
