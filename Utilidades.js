/**
 * ==================== UTILIDADES (TASA DE CAMBIO MULTIMONEDA) ====================
 */

function obtenerTasasDeCambio() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('TASAS_USD_GLOBAL');
  
  if (cached) {
    return JSON.parse(cached); // Retorna objeto { PEN: 3.75, COP: 4200, ... }
  }

  try {
    // API gratuita confiable. Base USD.
    const response = UrlFetchApp.fetch('https://api.exchangerate-api.com/v4/latest/USD');
    const data = JSON.parse(response.getContentText());
    
    if (data && data.rates) {
      // Guardamos todas las tasas por 6 horas
      cache.put('TASAS_USD_GLOBAL', JSON.stringify(data.rates), 21600); 
      return data.rates;
    }
  } catch (e) {
    Logger.log('Error API TC: ' + e.message);
  }
  
  // Fallback de emergencia si falla la API
  return { PEN: 3.75, COP: 4100, MXN: 17.5, USD: 1 }; 
}

// Mantenemos esta función para compatibilidad con MetaAds (usa solo PEN o USD por defecto)
function obtenerTasaActual() {
  const tasas = obtenerTasasDeCambio();
  return tasas['PEN'] || 3.75;
}

function obtenerHistorialTasasConRelleno() {
  // Para ventas históricas, simplificaremos usando la tasa actual
  // Si necesitas precisión histórica exacta por día para COP y PEN,
  // la lógica se complica mucho. Por ahora usamos la tasa del día para todo el cálculo.
  // Esto es estándar en reportes rápidos.
  return obtenerTasasDeCambio();
}

/**
 * Obtiene la tasa de cambio histórica para una fecha específica
 * Usa frankfurter.app (API gratuita con históricos)
 * @param {Date|string} fecha - Fecha para consultar
 * @param {string} moneda - Código de moneda (PEN, COP, MXN, ARS)
 * @returns {number} - Tasa de cambio (cuántas unidades de moneda por 1 USD)
 */
function obtenerTasaHistorica(fecha, moneda) {
  if (moneda === 'USD') return 1;

  const cache = CacheService.getScriptCache();
  const fechaStr = fecha instanceof Date
    ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : fecha;

  const cacheKey = `TASA_${moneda}_${fechaStr}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    return parseFloat(cached);
  }

  try {
    // frankfurter.app es gratuita y soporta históricos
    const url = `https://api.frankfurter.app/${fechaStr}?from=USD&to=${moneda}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.rates && data.rates[moneda]) {
        const tasa = data.rates[moneda];
        cache.put(cacheKey, tasa.toString(), 86400); // Cache por 24 horas
        return tasa;
      }
    }
  } catch (e) {
    Logger.log(`Error obteniendo tasa histórica ${moneda} para ${fechaStr}: ${e.message}`);
  }

  // Fallback: usar tasa actual
  const tasasActuales = obtenerTasasDeCambio();
  return tasasActuales[moneda] || 1;
}

/**
 * Obtiene tasas de cambio para un rango de fechas
 * @param {Date} fechaInicio - Fecha de inicio
 * @param {Date} fechaFin - Fecha de fin
 * @param {string} moneda - Código de moneda
 * @returns {Object} - Mapa de fecha -> tasa { '2024-12-01': 3.75, ... }
 */
function obtenerTasasRango(fechaInicio, fechaFin, moneda) {
  if (moneda === 'USD') return {};

  const tasas = {};
  const fechaActual = new Date(fechaInicio);

  while (fechaActual <= fechaFin) {
    const fechaStr = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    tasas[fechaStr] = obtenerTasaHistorica(fechaStr, moneda);
    fechaActual.setDate(fechaActual.getDate() + 1);
  }

  return tasas;
}

/**
 * Calcula las fechas de inicio y fin para un rango dado
 * @param {string} rango - Nombre del rango (HOY, AYER, LAST_3D, etc.)
 * @param {string} timezone - Zona horaria
 * @returns {Object} - { inicio: Date, fin: Date }
 */
function calcularFechasDeRango(rango, timezone) {
  const hoy = new Date();
  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);

  let inicio, fin;

  switch (rango) {
    case 'HOY':
      inicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
      fin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);
      break;
    case 'AYER':
      inicio = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate());
      fin = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 23, 59, 59);
      break;
    case 'LAST_3D':
      fin = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 23, 59, 59);
      inicio = new Date(fin);
      inicio.setDate(fin.getDate() - 2);
      inicio.setHours(0, 0, 0, 0);
      break;
    case 'LAST_5D':
      fin = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 23, 59, 59);
      inicio = new Date(fin);
      inicio.setDate(fin.getDate() - 4);
      inicio.setHours(0, 0, 0, 0);
      break;
    case 'LAST_7D':
      fin = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 23, 59, 59);
      inicio = new Date(fin);
      inicio.setDate(fin.getDate() - 6);
      inicio.setHours(0, 0, 0, 0);
      break;
    case 'LAST_30D':
      fin = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 23, 59, 59);
      inicio = new Date(fin);
      inicio.setDate(fin.getDate() - 29);
      inicio.setHours(0, 0, 0, 0);
      break;
    case 'MAXIMUM':
    default:
      // MAXIMUM va desde 2020 hasta AHORA (incluye hoy)
      fin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);
      inicio = new Date(2020, 0, 1); // Desde 2020
      break;
  }

  return { inicio, fin };
}