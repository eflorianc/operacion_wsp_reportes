/**
 * MÓDULO DE EXTRACCIÓN DE CREATIVOS
 * Extrae datos de anuncios desde Meta Ads Manager
 */

/**
 * Construye el parámetro de tiempo para la API de Meta
 * @param {string} rango - El rango seleccionado por el usuario
 * @param {string} timezone - Zona horaria del spreadsheet
 * @returns {string|null} - Parámetro formateado o null si es inválido
 */
function construirParametroTiempo(rango, timezone) {
  // Presets válidos de Meta API
  const presetsValidos = ['today', 'yesterday', 'last_3d', 'last_7d', 'last_14d', 'last_28d', 'last_30d', 'last_90d', 'maximum'];

  // Mapeo de aliases
  const aliases = {
    'lifetime': 'maximum',
    'historico': 'maximum',
    'all': 'maximum'
  };

  // Si es un alias, convertirlo
  if (aliases[rango]) {
    return `date_preset=${aliases[rango]}`;
  }

  // Si es un preset válido, usarlo directamente
  if (presetsValidos.includes(rango)) {
    return `date_preset=${rango}`;
  }

  // Rangos personalizados (last_Xd donde X no es estándar)
  const matchCustom = rango.match(/^last_(\d+)d$/);
  if (matchCustom) {
    const dias = parseInt(matchCustom[1]);
    return construirTimeRange(dias, timezone);
  }

  return null;
}

/**
 * Construye un time_range personalizado para X días atrás
 * @param {number} dias - Número de días hacia atrás
 * @param {string} timezone - Zona horaria
 * @returns {string} - Parámetro time_range formateado
 */
function construirTimeRange(dias, timezone) {
  const hoy = new Date();
  const hasta = new Date(hoy);
  hasta.setDate(hoy.getDate() - 1); // Ayer (datos más recientes completos)

  const desde = new Date(hasta);
  desde.setDate(hasta.getDate() - dias + 1); // X días atrás desde ayer

  const formatearFecha = (fecha) => Utilities.formatDate(fecha, timezone, 'yyyy-MM-dd');

  // Meta API requiere JSON con comillas dobles, URL-encoded
  const timeRange = {
    since: formatearFecha(desde),
    until: formatearFecha(hasta)
  };

  return 'time_range=' + encodeURIComponent(JSON.stringify(timeRange));
}

/**
 * Extrae datos de una cuenta publicitaria específica
 * @param {string} accountId - ID de la cuenta publicitaria
 * @param {string} token - Token de acceso
 * @param {string} timeParam - Parámetro de tiempo formateado
 * @param {string} filtroProducto - Filtro opcional por nombre de producto
 * @returns {Array} - Array de objetos con datos de anuncios
 */
function extraerDatosDeCuenta(accountId, token, timeParam, filtroProducto) {
  const campos = 'campaign_name,adset_name,ad_name,ad_id,spend,impressions,reach,clicks';
  const url = `https://graph.facebook.com/v22.0/${accountId}/insights?level=ad&fields=${campos}&${timeParam}&limit=500&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const errorData = JSON.parse(responseText);
    const errorMsg = errorData.error ? errorData.error.message : 'Error desconocido';
    throw new Error(errorMsg);
  }

  const json = JSON.parse(responseText);
  let datos = json.data || [];

  // Manejar paginación si hay más resultados
  let nextUrl = json.paging && json.paging.next;
  while (nextUrl) {
    const nextResponse = UrlFetchApp.fetch(nextUrl, { muteHttpExceptions: true });
    if (nextResponse.getResponseCode() === 200) {
      const nextJson = JSON.parse(nextResponse.getContentText());
      datos = datos.concat(nextJson.data || []);
      nextUrl = nextJson.paging && nextJson.paging.next;
    } else {
      break;
    }
  }

  // Filtrar por producto si se especificó
  if (filtroProducto && filtroProducto.length > 0) {
    const filtroLower = filtroProducto.toLowerCase();
    datos = datos.filter(item =>
      item.campaign_name && item.campaign_name.toLowerCase().includes(filtroLower)
    );
  }

  return datos;
}

/**
 * Extrae datos completos de cuenta incluyendo estados y presupuestos
 * OPTIMIZADO: Usa Batch API de Facebook, país desde configuración
 * @param {string} accountId - ID de cuenta publicitaria
 * @param {string} token - Token de acceso
 * @param {string} timeParam - Parámetro de tiempo formateado
 * @param {string} filtroProducto - Filtro opcional por nombre de producto
 * @returns {Array} - Array de objetos con datos completos de anuncios
 */
function extraerDatosCompletoDeCuenta(accountId, token, timeParam, filtroProducto) {
  // Obtener configuración para mapear países
  const config = obtenerConfiguracion();
  const monedasPais = config.MONEDAS_PAIS;
  const paisesConocidos = Object.keys(monedasPais);

  // Obtener insights básicos - ACTUALIZADO A v22.0
  // unique_inline_link_clicks = clics únicos en el enlace (campo correcto)
  const campos = 'campaign_name,adset_name,ad_name,ad_id,campaign_id,adset_id,spend,impressions,reach,unique_inline_link_clicks';
  const url = `https://graph.facebook.com/v22.0/${accountId}/insights?level=ad&fields=${campos}&${timeParam}&limit=500&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const errorData = JSON.parse(responseText);
    const errorMsg = errorData.error ? errorData.error.message : 'Error desconocido';
    throw new Error(errorMsg);
  }

  const json = JSON.parse(responseText);
  let datos = json.data || [];

  // LOG: Ver qué campos vienen en el primer resultado
  if (datos.length > 0) {
    Logger.log('Campos disponibles en insights: ' + JSON.stringify(Object.keys(datos[0])));
    Logger.log('Primer item completo: ' + JSON.stringify(datos[0]));
  }

  // Manejar paginación
  let nextUrl = json.paging && json.paging.next;
  while (nextUrl) {
    const nextResponse = UrlFetchApp.fetch(nextUrl, { muteHttpExceptions: true });
    if (nextResponse.getResponseCode() === 200) {
      const nextJson = JSON.parse(nextResponse.getContentText());
      datos = datos.concat(nextJson.data || []);
      nextUrl = nextJson.paging && nextJson.paging.next;
    } else {
      break;
    }
  }

  // Filtrar por producto si se especificó
  if (filtroProducto && filtroProducto.length > 0) {
    const filtroLower = filtroProducto.toLowerCase();
    datos = datos.filter(item =>
      item.campaign_name && item.campaign_name.toLowerCase().includes(filtroLower)
    );
  }

  if (datos.length === 0) {
    return [];
  }

  // Identificar IDs únicos para evitar consultas duplicadas
  const adsUnicos = [...new Set(datos.map(d => d.ad_id))];
  const adsetsUnicos = [...new Set(datos.map(d => d.adset_id))];
  const campaignsUnicos = [...new Set(datos.map(d => d.campaign_id))];

  // Caches para almacenar los datos
  const adCache = {};
  const adsetCache = {};
  const campaignCache = {};

  // Función para hacer Batch API requests - ACTUALIZADO A v22.0
  function ejecutarBatch(requests, token) {
    const batchUrl = 'https://graph.facebook.com/v22.0/';
    const payload = {
      access_token: token,
      batch: JSON.stringify(requests)
    };

    try {
      const response = UrlFetchApp.fetch(batchUrl, {
        method: 'post',
        payload: payload,
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      } else {
        Logger.log('Error en batch: ' + response.getContentText());
      }
    } catch (e) {
      Logger.log('Error ejecutando batch: ' + e.message);
    }
    return [];
  }

  // Obtener estados de ads en batches de 50
  for (let i = 0; i < adsUnicos.length; i += 50) {
    const batch = adsUnicos.slice(i, i + 50);
    const requests = batch.map(adId => ({
      method: 'GET',
      relative_url: `${adId}?fields=effective_status`
    }));

    const results = ejecutarBatch(requests, token);

    results.forEach((result, idx) => {
      if (result && result.code === 200) {
        try {
          const data = JSON.parse(result.body);
          adCache[batch[idx]] = data.effective_status || 'UNKNOWN';
        } catch (e) {
          adCache[batch[idx]] = 'UNKNOWN';
        }
      } else {
        adCache[batch[idx]] = 'UNKNOWN';
      }
    });
  }

  // Obtener estados y presupuestos de adsets en batches de 50
  for (let i = 0; i < adsetsUnicos.length; i += 50) {
    const batch = adsetsUnicos.slice(i, i + 50);
    const requests = batch.map(adsetId => ({
      method: 'GET',
      relative_url: `${adsetId}?fields=effective_status,daily_budget,lifetime_budget`
    }));

    const results = ejecutarBatch(requests, token);

    results.forEach((result, idx) => {
      if (result && result.code === 200) {
        try {
          const data = JSON.parse(result.body);

          let presupuesto = 0;
          if (data.daily_budget) {
            presupuesto = parseFloat(data.daily_budget) / 100;
          } else if (data.lifetime_budget) {
            presupuesto = parseFloat(data.lifetime_budget) / 100;
          }

          adsetCache[batch[idx]] = {
            status: data.effective_status || 'UNKNOWN',
            presupuesto: presupuesto
          };
        } catch (e) {
          adsetCache[batch[idx]] = { status: 'UNKNOWN', presupuesto: 0 };
        }
      } else {
        adsetCache[batch[idx]] = { status: 'UNKNOWN', presupuesto: 0 };
      }
    });
  }

  // Obtener estados de campaigns en batches de 50
  for (let i = 0; i < campaignsUnicos.length; i += 50) {
    const batch = campaignsUnicos.slice(i, i + 50);
    const requests = batch.map(campaignId => ({
      method: 'GET',
      relative_url: `${campaignId}?fields=effective_status`
    }));

    const results = ejecutarBatch(requests, token);

    results.forEach((result, idx) => {
      if (result && result.code === 200) {
        try {
          const data = JSON.parse(result.body);
          campaignCache[batch[idx]] = data.effective_status || 'UNKNOWN';
        } catch (e) {
          campaignCache[batch[idx]] = 'UNKNOWN';
        }
      } else {
        campaignCache[batch[idx]] = 'UNKNOWN';
      }
    });
  }

  // Función para extraer país del nombre de campaña
  function extraerPaisDeCampana(nombreCampana) {
    if (!nombreCampana) return 'N/A';
    const nombreUpper = nombreCampana.toUpperCase();

    for (const pais of paisesConocidos) {
      if (nombreUpper.includes(pais)) {
        return pais;
      }
    }
    return 'N/A';
  }

  // Enriquecer datos usando los caches
  const datosEnriquecidos = datos.map(item => {
    const adStatus = adCache[item.ad_id] || 'UNKNOWN';
    const adsetData = adsetCache[item.adset_id] || { status: 'UNKNOWN', presupuesto: 0 };
    const campaignStatus = campaignCache[item.campaign_id] || 'UNKNOWN';

    // Extraer país del nombre de la campaña
    const pais = extraerPaisDeCampana(item.campaign_name);

    // Determinar estado general
    let estadoGeneral = 'EN PAUSA';
    if (adStatus === 'ACTIVE' && adsetData.status === 'ACTIVE' && campaignStatus === 'ACTIVE') {
      estadoGeneral = 'EN CIRCULACIÓN';
    }

    return {
      ...item,
      ad_status: adStatus,
      adset_status: adsetData.status,
      campaign_status: campaignStatus,
      estado_general: estadoGeneral,
      presupuesto: adsetData.presupuesto,
      pais: pais
    };
  });

  return datosEnriquecidos;
}

/**
 * Extrae datos de TODOS los rangos de tiempo en una sola tabla
 * Rangos: maximum, today, yesterday, last_3d, last_5d
 */
function extraerTodosLosRangos() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const token = props.getProperty('META_TOKEN');
  const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

  if (!token) {
    ui.alert('❌ Error', 'No hay token de Meta configurado.', ui.ButtonSet.OK);
    return;
  }

  if (cuentas.length === 0) {
    ui.alert('❌ Error', 'No hay cuentas publicitarias configuradas.', ui.ButtonSet.OK);
    return;
  }

  // 1. Seleccionar país con listado numérico
  const paises = [
    { num: 0, nombre: 'TODOS', filtro: '' },
    { num: 1, nombre: 'PERU', filtro: 'peru' },
    { num: 2, nombre: 'MEXICO', filtro: 'mexico' },
    { num: 3, nombre: 'COLOMBIA', filtro: 'colombia' },
    { num: 4, nombre: 'ARGENTINA', filtro: 'argentina' }
  ];

  const menuPaises = paises.map(p => `${p.num}. ${p.nombre}`).join('\n');
  const respPais = ui.prompt(
    '🌎 Seleccionar País',
    `Ingresa el número del país:\n\n${menuPaises}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (respPais.getSelectedButton() !== ui.Button.OK) return;

  const numPais = parseInt(respPais.getResponseText().trim());
  const paisSeleccionado = paises.find(p => p.num === numPais);

  if (!paisSeleccionado) {
    ui.alert('❌ Error', 'Número de país no válido.', ui.ButtonSet.OK);
    return;
  }

  const filtroPais = paisSeleccionado.filtro;

  // 2. Preguntar filtro de producto (opcional)
  const respProd = ui.prompt('🔍 Filtro de Producto', 'Introduce el nombre del producto (o déjalo vacío para todos):', ui.ButtonSet.OK_CANCEL);
  if (respProd.getSelectedButton() !== ui.Button.OK) return;
  const filtroProducto = respProd.getResponseText().trim();

  // Preparar hoja
  let hoja = ss.getSheetByName('📊 Todos los Rangos');
  if (!hoja) hoja = ss.insertSheet('📊 Todos los Rangos');
  hoja.clear();

  // Quitar filtros existentes si los hay
  if (hoja.getFilter()) {
    hoja.getFilter().remove();
  }

  const headers = [
    'TIPO', 'RANGO', 'CAMPAÑA', 'CONJUNTO', 'ANUNCIO', 'AD ID', 'PAÍS', 'ESTADO', 'PRESUPUESTO',
    'GASTO', 'IGV', 'GASTO TOTAL', 'FACT USD', 'ROAS', 'UTILIDAD', 'ROI',
    'IMPRESIONES', 'ALCANCE', 'CLICS ÚNICOS', 'COSTO POR CLIC', 'MENSAJES', '% MENSAJES', 'COSTO POR MENSAJE',
    '# VENTAS', 'CVR', 'COSTO POR COMPRA', 'CPM', 'T.C.'
  ];
  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#9c27b0')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

  // Fijar la primera fila para que siempre sea visible
  hoja.setFrozenRows(1);

  // Definir los rangos a extraer
  const rangos = [
    { nombre: 'HOY', param: 'date_preset=today' },
    { nombre: 'AYER', param: 'date_preset=yesterday' },
    { nombre: 'LAST_3D', param: 'date_preset=last_3d' },
    { nombre: 'LAST_5D', param: construirTimeRange(5, ss.getSpreadsheetTimeZone()) },
    { nombre: 'LAST_7D', param: 'date_preset=last_7d' },
    { nombre: 'LAST_30D', param: 'date_preset=last_30d' },
    { nombre: 'MAXIMUM', param: 'date_preset=maximum' }
  ];

  let todosLosResultados = [];
  let errores = [];
  let totalPorRango = {};
  let datosPorRango = {}; // Almacena datos agrupados por rango

  // Mostrar progreso
  ss.toast('Extrayendo datos de todos los rangos...', '⏳ Procesando', -1);

  rangos.forEach((rangoConfig, index) => {
    ss.toast(`Extrayendo: ${rangoConfig.nombre} (${index + 1}/${rangos.length})`, '⏳ Procesando', -1);
    totalPorRango[rangoConfig.nombre] = 0;
    datosPorRango[rangoConfig.nombre] = {
      filas: [],
      totales: { gasto: 0, igv: 0, gastoTotal: 0, alcance: 0, clics: 0, mensajes: 0, factUSD: 0, numVentas: 0, utilidad: 0, impresiones: 0, presupuesto: 0 }
    };

    // Calcular fechas del rango y obtener compras y mensajes filtrados por ese rango
    const fechasRango = calcularFechasDeRango(rangoConfig.nombre, ss.getSpreadsheetTimeZone());
    const comprasPorAdId = obtenerComprasPorAdId(filtroPais, fechasRango.inicio, fechasRango.fin);
    const mensajesPorAdId = obtenerMensajesPorAdId(filtroPais, fechasRango.inicio, fechasRango.fin);
    Logger.log(`${rangoConfig.nombre}: Compras encontradas para ${Object.keys(comprasPorAdId).length} AD IDs`);
    Logger.log(`${rangoConfig.nombre}: Mensajes encontrados para ${Object.keys(mensajesPorAdId).length} AD IDs`);

    cuentas.forEach(cuenta => {
      try {
        // Usar la función completa que obtiene estados y presupuestos
        let datos = extraerDatosCompletoDeCuenta(cuenta[0], token, rangoConfig.param, filtroProducto);

        // Filtrar por país si se seleccionó uno específico
        if (filtroPais) {
          datos = datos.filter(item =>
            item.campaign_name && item.campaign_name.toLowerCase().includes(filtroPais)
          );
        }

        datos.forEach((item, itemIdx) => {
          const gasto = parseFloat(item.spend) || 0;
          const igv = gasto * 0.18;
          const gastoTotal = gasto + igv;
          const alcance = parseInt(item.reach) || 0;
          const clicsUnicos = parseInt(item.unique_inline_link_clicks) || 0;
          const impresiones = parseInt(item.impressions) || 0;
          const cpm = impresiones > 0 ? (gasto / impresiones) * 1000 : 0;

          // LOG: Ver primer item para debug de clics únicos en enlace
          if (itemIdx === 0) {
            Logger.log(`Primer anuncio - unique_inline_link_clicks: ${item.unique_inline_link_clicks}, clicsUnicos: ${clicsUnicos}`);
          }

          // Obtener datos adicionales del item
          const estadoGeneral = item.estado_general || 'EN PAUSA';
          const presupuesto = item.presupuesto || 0;
          const pais = item.pais || 'N/A';

          // Buscar facturación y mensajes por AD ID para este rango
          const adId = item.ad_id || '';
          let factUSD = 0;
          let numVentas = 0;
          let tasaCambio = '';
          let numMensajes = 0;

          // Vincular compras del rango actual
          if (adId && comprasPorAdId[adId]) {
            const compra = comprasPorAdId[adId];
            factUSD = compra.totalUSD;
            numVentas = compra.cantidadVentas;
            tasaCambio = compra.tasa;
          }

          // Vincular mensajes del rango actual
          if (adId && mensajesPorAdId[adId]) {
            numMensajes = mensajesPorAdId[adId].cantidadMensajes;
          }

          // LOG: Ver primer anuncio con mensajes
          if (itemIdx === 0 && numMensajes > 0) {
            Logger.log(`Primer anuncio con mensajes - AD ID: ${adId}, mensajes: ${numMensajes}`);
          }

          const roas = gastoTotal > 0 ? factUSD / gastoTotal : 0;
          const utilidad = factUSD - gastoTotal;
          const roi = gastoTotal > 0 ? utilidad / gastoTotal : 0;

          // Nuevos cálculos
          const costoPorClic = clicsUnicos > 0 ? gastoTotal / clicsUnicos : 0;
          const porcentajeMensajes = clicsUnicos > 0 ? numMensajes / clicsUnicos : 0;
          const costoPorMensaje = numMensajes > 0 ? gastoTotal / numMensajes : 0;
          const cvr = numMensajes > 0 ? numVentas / numMensajes : 0;
          const costoPorCompra = numVentas > 0 ? gastoTotal / numVentas : 0;

          datosPorRango[rangoConfig.nombre].filas.push([
            'DATO',
            rangoConfig.nombre,
            item.campaign_name || '',
            item.adset_name || '',
            item.ad_name || '',
            adId,
            pais,
            estadoGeneral,
            presupuesto,
            gasto,
            igv,
            gastoTotal,
            factUSD,
            roas,
            utilidad,
            roi,
            impresiones,
            alcance,
            clicsUnicos,
            costoPorClic,
            numMensajes,
            porcentajeMensajes,
            costoPorMensaje,
            numVentas,
            cvr,
            costoPorCompra,
            cpm,
            tasaCambio
          ]);

          // Acumular totales
          datosPorRango[rangoConfig.nombre].totales.gasto += gasto;
          datosPorRango[rangoConfig.nombre].totales.igv += igv;
          datosPorRango[rangoConfig.nombre].totales.gastoTotal += gastoTotal;
          datosPorRango[rangoConfig.nombre].totales.alcance += alcance;
          datosPorRango[rangoConfig.nombre].totales.clics += clicsUnicos;
          datosPorRango[rangoConfig.nombre].totales.mensajes += numMensajes;
          datosPorRango[rangoConfig.nombre].totales.factUSD += factUSD;
          datosPorRango[rangoConfig.nombre].totales.numVentas += numVentas;
          datosPorRango[rangoConfig.nombre].totales.utilidad += utilidad;
          datosPorRango[rangoConfig.nombre].totales.impresiones += impresiones;
          datosPorRango[rangoConfig.nombre].totales.presupuesto += presupuesto;

          totalPorRango[rangoConfig.nombre]++;
        });
      } catch (e) {
        errores.push(`${rangoConfig.nombre} - Cuenta ${cuenta[0]}: ${e.message}`);
        Logger.log(`Error en ${rangoConfig.nombre}, cuenta ${cuenta[0]}: ${e.message}`);
      }
    });
  });

  // Construir resultados finales con filas de totales
  rangos.forEach(rangoConfig => {
    const datos = datosPorRango[rangoConfig.nombre];
    if (datos.filas.length > 0) {
      // Agregar filas de datos
      todosLosResultados = todosLosResultados.concat(datos.filas);

      // Calcular métricas de los totales
      const totales = datos.totales;
      const cpmTotal = totales.impresiones > 0 ? (totales.gasto / totales.impresiones) * 1000 : 0;
      const roasTotal = totales.gastoTotal > 0 ? totales.factUSD / totales.gastoTotal : 0;
      const roiTotal = totales.gastoTotal > 0 ? totales.utilidad / totales.gastoTotal : 0;

      // Nuevos cálculos de totales
      const costoPorClicTotal = totales.clics > 0 ? totales.gastoTotal / totales.clics : 0;
      const porcentajeMensajesTotal = totales.clics > 0 ? totales.mensajes / totales.clics : 0;
      const costoPorMensajeTotal = totales.mensajes > 0 ? totales.gastoTotal / totales.mensajes : 0;
      const cvrTotal = totales.mensajes > 0 ? totales.numVentas / totales.mensajes : 0;
      const costoPorCompraTotal = totales.numVentas > 0 ? totales.gastoTotal / totales.numVentas : 0;

      // Agregar fila de totales (con nuevas columnas)
      todosLosResultados.push([
        'TOTAL',
        rangoConfig.nombre,
        `(${datos.filas.length} anuncios)`,
        '', '', '', '', '',  // CONJUNTO, ANUNCIO, AD ID, PAÍS, ESTADO vacíos
        totales.presupuesto,  // PRESUPUESTO total
        totales.gasto,
        totales.igv,
        totales.gastoTotal,
        totales.factUSD,
        roasTotal,
        totales.utilidad,
        roiTotal,
        totales.impresiones,
        totales.alcance,
        totales.clics,        // CLICS ÚNICOS
        costoPorClicTotal,
        totales.mensajes,     // MENSAJES
        porcentajeMensajesTotal,
        costoPorMensajeTotal,
        totales.numVentas,
        cvrTotal,
        costoPorCompraTotal,
        cpmTotal,
        ''  // T.C. vacío en totales
      ]);
    }
  });

  // Escribir resultados
  if (todosLosResultados.length > 0) {
    hoja.getRange(2, 1, todosLosResultados.length, headers.length).setValues(todosLosResultados);

    // Formatear columnas (ajustado para nuevas columnas reorganizadas)
    hoja.getRange(2, 9, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');    // PRESUPUESTO
    hoja.getRange(2, 10, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // GASTO
    hoja.getRange(2, 11, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // IGV
    hoja.getRange(2, 12, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // GASTO TOTAL
    hoja.getRange(2, 13, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // FACT USD
    hoja.getRange(2, 14, todosLosResultados.length, 1).setNumberFormat('0.00');        // ROAS
    hoja.getRange(2, 15, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // UTILIDAD
    hoja.getRange(2, 16, todosLosResultados.length, 1).setNumberFormat('0.00%');       // ROI
    hoja.getRange(2, 17, todosLosResultados.length, 1).setNumberFormat('#,##0');       // IMPRESIONES
    hoja.getRange(2, 18, todosLosResultados.length, 1).setNumberFormat('#,##0');       // ALCANCE
    hoja.getRange(2, 19, todosLosResultados.length, 1).setNumberFormat('#,##0');       // CLICS ÚNICOS
    hoja.getRange(2, 20, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // COSTO POR CLIC
    hoja.getRange(2, 21, todosLosResultados.length, 1).setNumberFormat('#,##0');       // MENSAJES
    hoja.getRange(2, 22, todosLosResultados.length, 1).setNumberFormat('0.00%');       // % MENSAJES
    hoja.getRange(2, 23, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // COSTO POR MENSAJE
    hoja.getRange(2, 24, todosLosResultados.length, 1).setNumberFormat('#,##0');       // # VENTAS
    hoja.getRange(2, 25, todosLosResultados.length, 1).setNumberFormat('0.00%');       // CVR
    hoja.getRange(2, 26, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // COSTO POR COMPRA
    hoja.getRange(2, 27, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // CPM
    hoja.getRange(2, 28, todosLosResultados.length, 1).setNumberFormat('0.00');        // T.C.

    // Colorear filas por rango
    const colores = {
      'HOY': '#e8f5e9',
      'AYER': '#fff3e0',
      'LAST_3D': '#e3f2fd',
      'LAST_5D': '#fce4ec',
      'LAST_7D': '#e0f7fa',
      'LAST_30D': '#fff8e1',
      'MAXIMUM': '#f3e5f5'
    };

    // Colores más oscuros para filas de totales
    const coloresTotales = {
      'HOY': '#a5d6a7',
      'AYER': '#ffcc80',
      'LAST_3D': '#90caf9',
      'LAST_5D': '#f48fb1',
      'LAST_7D': '#80deea',
      'LAST_30D': '#ffe082',
      'MAXIMUM': '#ce93d8'
    };

    todosLosResultados.forEach((fila, i) => {
      const tipo = fila[0];  // DATO o TOTAL
      const rango = fila[1]; // HOY, AYER, etc.
      const rowRange = hoja.getRange(i + 2, 1, 1, headers.length);

      if (tipo === 'TOTAL') {
        // Fila de totales: color oscuro y negrita
        const colorTotal = coloresTotales[rango] || '#bdbdbd';
        rowRange.setBackground(colorTotal)
                .setFontWeight('bold')
                .setBorder(true, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID);
      } else {
        // Fila normal
        const color = colores[rango] || '#ffffff';
        rowRange.setBackground(color);
      }
    });

    // Agregar filtro automático en las cabeceras
    const rangoFiltro = hoja.getRange(1, 1, todosLosResultados.length + 1, headers.length);
    rangoFiltro.createFilter();

    // Agregar fila de TOTAL GENERAL que se actualiza automáticamente con filtros
    // Esta fila usa SUBTOTAL() que solo suma filas visibles
    const filaTotal = todosLosResultados.length + 2;
    const primeraFilaDatos = 2;
    const ultimaFilaDatos = todosLosResultados.length + 1;

    // Fórmulas SUBTOTAL (109 = SUM excluyendo filas ocultas y filtradas)
    const formulaPresupuesto = `=SUBTOTAL(109,I${primeraFilaDatos}:I${ultimaFilaDatos})`;
    const formulaGasto = `=SUBTOTAL(109,J${primeraFilaDatos}:J${ultimaFilaDatos})`;
    const formulaIGV = `=SUBTOTAL(109,K${primeraFilaDatos}:K${ultimaFilaDatos})`;
    const formulaGastoTotal = `=SUBTOTAL(109,L${primeraFilaDatos}:L${ultimaFilaDatos})`;
    const formulaFactUSD = `=SUBTOTAL(109,M${primeraFilaDatos}:M${ultimaFilaDatos})`;
    const formulaROAS = `=IF(L${filaTotal}>0,M${filaTotal}/L${filaTotal},0)`;
    const formulaUtilidad = `=SUBTOTAL(109,O${primeraFilaDatos}:O${ultimaFilaDatos})`;
    const formulaROI = `=IF(L${filaTotal}>0,O${filaTotal}/L${filaTotal},0)`;
    const formulaImpresiones = `=SUBTOTAL(109,Q${primeraFilaDatos}:Q${ultimaFilaDatos})`;
    const formulaAlcance = `=SUBTOTAL(109,R${primeraFilaDatos}:R${ultimaFilaDatos})`;
    const formulaClicsUnicos = `=SUBTOTAL(109,S${primeraFilaDatos}:S${ultimaFilaDatos})`;
    const formulaCostoPorClic = `=IF(S${filaTotal}>0,L${filaTotal}/S${filaTotal},0)`;
    const formulaMensajes = `=SUBTOTAL(109,U${primeraFilaDatos}:U${ultimaFilaDatos})`;
    const formulaPorcentajeMensajes = `=IF(S${filaTotal}>0,U${filaTotal}/S${filaTotal},0)`;
    const formulaCostoPorMensaje = `=IF(U${filaTotal}>0,L${filaTotal}/U${filaTotal},0)`;
    const formulaVentas = `=SUBTOTAL(109,X${primeraFilaDatos}:X${ultimaFilaDatos})`;
    const formulaCVR = `=IF(U${filaTotal}>0,X${filaTotal}/U${filaTotal},0)`;
    const formulaCostoPorCompra = `=IF(X${filaTotal}>0,L${filaTotal}/X${filaTotal},0)`;
    const formulaCPM = `=IF(Q${filaTotal}>0,(J${filaTotal}/Q${filaTotal})*1000,0)`;

    hoja.getRange(filaTotal, 1, 1, headers.length).setValues([[
      '',
      'TOTAL GENERAL',
      '(se actualiza automáticamente al filtrar)',
      '', '', '', '', '',
      formulaPresupuesto,
      formulaGasto,
      formulaIGV,
      formulaGastoTotal,
      formulaFactUSD,
      formulaROAS,
      formulaUtilidad,
      formulaROI,
      formulaImpresiones,
      formulaAlcance,
      formulaClicsUnicos,
      formulaCostoPorClic,
      formulaMensajes,
      formulaPorcentajeMensajes,
      formulaCostoPorMensaje,
      formulaVentas,
      formulaCVR,
      formulaCostoPorCompra,
      formulaCPM,
      ''
    ]]);

    // Formatear fila de total
    const rangoTotal = hoja.getRange(filaTotal, 1, 1, headers.length);
    rangoTotal.setBackground('#00695c')
              .setFontColor('white')
              .setFontWeight('bold')
              .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);

    hoja.getRange(filaTotal, 9).setNumberFormat('$#,##0.00');   // PRESUPUESTO
    hoja.getRange(filaTotal, 10).setNumberFormat('$#,##0.00');  // GASTO
    hoja.getRange(filaTotal, 11).setNumberFormat('$#,##0.00');  // IGV
    hoja.getRange(filaTotal, 12).setNumberFormat('$#,##0.00');  // GASTO TOTAL
    hoja.getRange(filaTotal, 13).setNumberFormat('$#,##0.00');  // FACT USD
    hoja.getRange(filaTotal, 14).setNumberFormat('0.00');       // ROAS
    hoja.getRange(filaTotal, 15).setNumberFormat('$#,##0.00');  // UTILIDAD
    hoja.getRange(filaTotal, 16).setNumberFormat('0.00%');      // ROI
    hoja.getRange(filaTotal, 17).setNumberFormat('#,##0');      // IMPRESIONES
    hoja.getRange(filaTotal, 18).setNumberFormat('#,##0');      // ALCANCE
    hoja.getRange(filaTotal, 19).setNumberFormat('#,##0');      // CLICS ÚNICOS
    hoja.getRange(filaTotal, 20).setNumberFormat('$#,##0.00');  // COSTO POR CLIC
    hoja.getRange(filaTotal, 21).setNumberFormat('#,##0');      // MENSAJES
    hoja.getRange(filaTotal, 22).setNumberFormat('0.00%');      // % MENSAJES
    hoja.getRange(filaTotal, 23).setNumberFormat('$#,##0.00');  // COSTO POR MENSAJE
    hoja.getRange(filaTotal, 24).setNumberFormat('#,##0');      // # VENTAS
    hoja.getRange(filaTotal, 25).setNumberFormat('0.00%');      // CVR
    hoja.getRange(filaTotal, 26).setNumberFormat('$#,##0.00');  // COSTO POR COMPRA
    hoja.getRange(filaTotal, 27).setNumberFormat('$#,##0.00');  // CPM

    // Ocultar la columna TIPO (opcional, el usuario puede filtrar por ella)
    // hoja.hideColumns(1);

    hoja.autoResizeColumns(1, headers.length);
    hoja.setColumnWidth(1, 60); // Columna TIPO más angosta

    // Resumen
    let resumen = `✅ Extracción completada\n\n🌎 País: ${paisSeleccionado.nombre}\n`;
    if (filtroProducto) resumen += `🔍 Producto: ${filtroProducto}\n`;
    resumen += '\n';
    for (const rango in totalPorRango) {
      resumen += `• ${rango}: ${totalPorRango[rango]} anuncios\n`;
    }
    resumen += `\n📊 Total: ${todosLosResultados.length - Object.keys(totalPorRango).filter(r => totalPorRango[r] > 0).length} anuncios`;

    if (errores.length > 0) {
      resumen += `\n\n⚠️ ${errores.length} error(es) encontrados.`;
    }

    ss.toast('', '', 1); // Cerrar toast
    ui.alert('Extracción Múltiple Completada', resumen, ui.ButtonSet.OK);
  } else {
    ss.toast('', '', 1);
    let mensaje = '⚠️ No se encontraron datos.';
    if (errores.length > 0) {
      mensaje += '\n\nErrores:\n' + errores.slice(0, 5).join('\n');
    }
    ui.alert('Sin Resultados', mensaje, ui.ButtonSet.OK);
  }
}

/**
 * Extrae reporte consolidado POR PRODUCTOS en lugar de por anuncios individuales
 * Muestra métricas agrupadas por producto y rango de fechas
 */
function extraerReportePorProductos() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const token = props.getProperty('META_TOKEN');
  const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

  if (!token) {
    ui.alert('❌ Error', 'No hay token de Meta configurado.', ui.ButtonSet.OK);
    return;
  }

  if (cuentas.length === 0) {
    ui.alert('❌ Error', 'No hay cuentas publicitarias configuradas.', ui.ButtonSet.OK);
    return;
  }

  // 1. Seleccionar país
  const paises = [
    { num: 0, nombre: 'TODOS', filtro: '' },
    { num: 1, nombre: 'PERU', filtro: 'peru' },
    { num: 2, nombre: 'MEXICO', filtro: 'mexico' },
    { num: 3, nombre: 'COLOMBIA', filtro: 'colombia' },
    { num: 4, nombre: 'ARGENTINA', filtro: 'argentina' }
  ];

  const menuPaises = paises.map(p => `${p.num}. ${p.nombre}`).join('\n');
  const respPais = ui.prompt(
    '🌎 Seleccionar País',
    `Ingresa el número del país:\n\n${menuPaises}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (respPais.getSelectedButton() !== ui.Button.OK) return;

  const numPais = parseInt(respPais.getResponseText().trim());
  const paisSeleccionado = paises.find(p => p.num === numPais);

  if (!paisSeleccionado) {
    ui.alert('❌ Error', 'Número de país no válido.', ui.ButtonSet.OK);
    return;
  }

  const filtroPais = paisSeleccionado.filtro;

  // 2. Obtener productos configurados
  const config = obtenerConfiguracion();
  const productosConfig = config.PRODUCTOS || {};

  if (Object.keys(productosConfig).length === 0) {
    ui.alert('❌ Error', 'No hay productos configurados. Ve a Configuración > Configurar Palabras Clave.', ui.ButtonSet.OK);
    return;
  }

  // Preparar hoja
  let hoja = ss.getSheetByName('📦 Reporte por Productos');
  if (!hoja) hoja = ss.insertSheet('📦 Reporte por Productos');
  hoja.clear();

  // Quitar filtros existentes si los hay
  if (hoja.getFilter()) {
    hoja.getFilter().remove();
  }

  const headers = [
    'TIPO', 'PRODUCTO', 'RANGO',
    'GASTO', 'IGV', 'GASTO TOTAL', 'FACT USD', 'ROAS', 'UTILIDAD', 'ROI',
    'IMPRESIONES', 'ALCANCE', 'CLICS ÚNICOS', 'COSTO POR CLIC',
    'MENSAJES', '% MENSAJES', 'COSTO POR MENSAJE',
    '# VENTAS', 'CVR', 'COSTO POR COMPRA', 'CPM'
  ];

  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#00796b')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

  // Fijar la primera fila
  hoja.setFrozenRows(1);

  // Definir los rangos a extraer
  const rangos = [
    { nombre: 'HOY', param: 'date_preset=today' },
    { nombre: 'AYER', param: 'date_preset=yesterday' },
    { nombre: 'LAST_3D', param: 'date_preset=last_3d' },
    { nombre: 'LAST_5D', param: construirTimeRange(5, ss.getSpreadsheetTimeZone()) },
    { nombre: 'LAST_7D', param: 'date_preset=last_7d' },
    { nombre: 'LAST_30D', param: 'date_preset=last_30d' },
    { nombre: 'MAXIMUM', param: 'date_preset=maximum' }
  ];

  let todosLosResultados = [];
  let errores = [];

  // OPTIMIZACIÓN: Extraer TODOS los datos UNA SOLA VEZ por rango, luego filtrar localmente
  ss.toast('Extrayendo datos de la API...', '⏳ Procesando', -1);

  // Objeto para guardar todos los datos por rango
  const datosPorRango = {};

  // Paso 1: Extraer TODOS los datos de TODOS los rangos (una sola vez)
  rangos.forEach((rangoConfig, index) => {
    ss.toast(`Extrayendo rango ${index + 1}/${rangos.length}: ${rangoConfig.nombre}...`, '⏳ API', -1);

    // Calcular fechas del rango
    const fechasRango = calcularFechasDeRango(rangoConfig.nombre, ss.getSpreadsheetTimeZone());
    const comprasPorAdId = obtenerComprasPorAdId(filtroPais, fechasRango.inicio, fechasRango.fin);
    const mensajesPorAdId = obtenerMensajesPorAdId(filtroPais, fechasRango.inicio, fechasRango.fin);

    datosPorRango[rangoConfig.nombre] = {
      datos: [],
      comprasPorAdId: comprasPorAdId,
      mensajesPorAdId: mensajesPorAdId
    };

    // Extraer datos de todas las cuentas para este rango
    cuentas.forEach(cuenta => {
      try {
        let datos = extraerDatosCompletoDeCuenta(cuenta[0], token, rangoConfig.param, '');

        // Filtrar por país si es necesario
        if (filtroPais) {
          datos = datos.filter(item =>
            item.campaign_name && item.campaign_name.toLowerCase().includes(filtroPais)
          );
        }

        // Guardar todos los datos sin filtrar por producto
        datosPorRango[rangoConfig.nombre].datos = datosPorRango[rangoConfig.nombre].datos.concat(datos);

      } catch (e) {
        errores.push(`${rangoConfig.nombre} - Cuenta ${cuenta[0]}: ${e.message}`);
        Logger.log(`Error en rango ${rangoConfig.nombre}, cuenta ${cuenta[0]}: ${e.message}`);
      }
    });

    Logger.log(`${rangoConfig.nombre}: ${datosPorRango[rangoConfig.nombre].datos.length} anuncios extraídos`);
  });

  // Paso 2: Procesar cada producto filtrando los datos ya obtenidos
  ss.toast('Procesando productos...', '⏳ Calculando', -1);

  let productoIndex = 0;
  const numProductos = Object.keys(productosConfig).length;

  for (const nombreProducto in productosConfig) {
    productoIndex++;
    const producto = productosConfig[nombreProducto];
    const palabrasClave = producto.palabrasClave || [];
    const paisProducto = producto.pais ? producto.pais.toLowerCase() : '';

    // Filtrar por país si es necesario
    if (filtroPais && paisProducto && paisProducto !== filtroPais) {
      Logger.log(`Saltando producto "${nombreProducto}" porque es de "${paisProducto}" y se filtró por "${filtroPais}"`);
      continue;
    }

    ss.toast(`Procesando producto ${productoIndex}/${numProductos}: ${nombreProducto}...`, '⏳ Calculando', -1);

    // Para cada rango, filtrar datos localmente
    rangos.forEach((rangoConfig) => {
      const rangoData = datosPorRango[rangoConfig.nombre];
      const comprasPorAdId = rangoData.comprasPorAdId;
      const mensajesPorAdId = rangoData.mensajesPorAdId;

      // Acumuladores para este producto+rango
      let totales = {
        gasto: 0,
        igv: 0,
        gastoTotal: 0,
        alcance: 0,
        clics: 0,
        mensajes: 0,
        factUSD: 0,
        numVentas: 0,
        utilidad: 0,
        impresiones: 0
      };

      // Filtrar datos por palabras clave del producto
      const datosFiltrados = rangoData.datos.filter(item => {
        if (!item.campaign_name) return false;
        const campaignNameUpper = item.campaign_name.toUpperCase();
        // El anuncio coincide si contiene ALGUNA de las palabras clave
        return palabrasClave.some(keyword =>
          campaignNameUpper.includes(keyword.toUpperCase())
        );
      });

      // Agregar métricas de los datos filtrados
      datosFiltrados.forEach(item => {
        const gasto = parseFloat(item.spend) || 0;
        const igv = gasto * 0.18;
        const gastoTotal = gasto + igv;
        const alcance = parseInt(item.reach) || 0;
        const clicsUnicos = parseInt(item.unique_inline_link_clicks) || 0;
        const impresiones = parseInt(item.impressions) || 0;

        const adId = item.ad_id || '';
        let factUSD = 0;
        let numVentas = 0;
        let numMensajes = 0;

        // Vincular compras y mensajes
        if (adId && comprasPorAdId[adId]) {
          const compra = comprasPorAdId[adId];
          factUSD = compra.totalUSD;
          numVentas = compra.cantidadVentas;
        }

        if (adId && mensajesPorAdId[adId]) {
          numMensajes = mensajesPorAdId[adId].cantidadMensajes;
        }

        const utilidad = factUSD - gastoTotal;

        // Acumular totales
        totales.gasto += gasto;
        totales.igv += igv;
        totales.gastoTotal += gastoTotal;
        totales.alcance += alcance;
        totales.clics += clicsUnicos;
        totales.mensajes += numMensajes;
        totales.factUSD += factUSD;
        totales.numVentas += numVentas;
        totales.utilidad += utilidad;
        totales.impresiones += impresiones;
      });

      // Solo agregar fila si hay datos
      if (totales.gastoTotal > 0 || totales.alcance > 0) {
        // Calcular métricas finales
        const roas = totales.gastoTotal > 0 ? totales.factUSD / totales.gastoTotal : 0;
        const roi = totales.gastoTotal > 0 ? totales.utilidad / totales.gastoTotal : 0;
        const cpm = totales.impresiones > 0 ? (totales.gasto / totales.impresiones) * 1000 : 0;
        const costoPorClic = totales.clics > 0 ? totales.gastoTotal / totales.clics : 0;
        const porcentajeMensajes = totales.clics > 0 ? totales.mensajes / totales.clics : 0;
        const costoPorMensaje = totales.mensajes > 0 ? totales.gastoTotal / totales.mensajes : 0;
        const cvr = totales.mensajes > 0 ? totales.numVentas / totales.mensajes : 0;
        const costoPorCompra = totales.numVentas > 0 ? totales.gastoTotal / totales.numVentas : 0;

        todosLosResultados.push([
          'DATO',
          nombreProducto,
          rangoConfig.nombre,
          totales.gasto,
          totales.igv,
          totales.gastoTotal,
          totales.factUSD,
          roas,
          totales.utilidad,
          roi,
          totales.impresiones,
          totales.alcance,
          totales.clics,
          costoPorClic,
          totales.mensajes,
          porcentajeMensajes,
          costoPorMensaje,
          totales.numVentas,
          cvr,
          costoPorCompra,
          cpm
        ]);
      }
    });
  }

  // Agrupar por rango para calcular totales
  const totalesPorRango = {};
  rangos.forEach(r => {
    totalesPorRango[r.nombre] = {
      gasto: 0, igv: 0, gastoTotal: 0, alcance: 0, clics: 0, mensajes: 0,
      factUSD: 0, numVentas: 0, utilidad: 0, impresiones: 0
    };
  });

  todosLosResultados.forEach(fila => {
    const rango = fila[2];
    if (totalesPorRango[rango]) {
      totalesPorRango[rango].gasto += fila[3];
      totalesPorRango[rango].igv += fila[4];
      totalesPorRango[rango].gastoTotal += fila[5];
      totalesPorRango[rango].factUSD += fila[6];
      totalesPorRango[rango].utilidad += fila[8];
      totalesPorRango[rango].impresiones += fila[10];
      totalesPorRango[rango].alcance += fila[11];
      totalesPorRango[rango].clics += fila[12];
      totalesPorRango[rango].mensajes += fila[14];
      totalesPorRango[rango].numVentas += fila[17];
    }
  });

  // Insertar filas de totales por rango
  const resultadosConTotales = [];
  rangos.forEach(rangoConfig => {
    // Agregar filas de datos de este rango
    const filasDatos = todosLosResultados.filter(f => f[2] === rangoConfig.nombre);
    resultadosConTotales.push(...filasDatos);

    // Agregar fila de total si hay datos
    if (filasDatos.length > 0) {
      const totales = totalesPorRango[rangoConfig.nombre];
      const roas = totales.gastoTotal > 0 ? totales.factUSD / totales.gastoTotal : 0;
      const roi = totales.gastoTotal > 0 ? totales.utilidad / totales.gastoTotal : 0;
      const cpm = totales.impresiones > 0 ? (totales.gasto / totales.impresiones) * 1000 : 0;
      const costoPorClic = totales.clics > 0 ? totales.gastoTotal / totales.clics : 0;
      const porcentajeMensajes = totales.clics > 0 ? totales.mensajes / totales.clics : 0;
      const costoPorMensaje = totales.mensajes > 0 ? totales.gastoTotal / totales.mensajes : 0;
      const cvr = totales.mensajes > 0 ? totales.numVentas / totales.mensajes : 0;
      const costoPorCompra = totales.numVentas > 0 ? totales.gastoTotal / totales.numVentas : 0;

      resultadosConTotales.push([
        'TOTAL',
        `TOTAL ${rangoConfig.nombre}`,
        rangoConfig.nombre,
        totales.gasto,
        totales.igv,
        totales.gastoTotal,
        totales.factUSD,
        roas,
        totales.utilidad,
        roi,
        totales.impresiones,
        totales.alcance,
        totales.clics,
        costoPorClic,
        totales.mensajes,
        porcentajeMensajes,
        costoPorMensaje,
        totales.numVentas,
        cvr,
        costoPorCompra,
        cpm
      ]);
    }
  });

  // Escribir resultados
  if (resultadosConTotales.length > 0) {
    hoja.getRange(2, 1, resultadosConTotales.length, headers.length).setValues(resultadosConTotales);

    // Formatear columnas
    hoja.getRange(2, 4, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');   // GASTO
    hoja.getRange(2, 5, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');   // IGV
    hoja.getRange(2, 6, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');   // GASTO TOTAL
    hoja.getRange(2, 7, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');   // FACT USD
    hoja.getRange(2, 8, resultadosConTotales.length, 1).setNumberFormat('0.00');        // ROAS
    hoja.getRange(2, 9, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');   // UTILIDAD
    hoja.getRange(2, 10, resultadosConTotales.length, 1).setNumberFormat('0.00%');      // ROI
    hoja.getRange(2, 11, resultadosConTotales.length, 1).setNumberFormat('#,##0');      // IMPRESIONES
    hoja.getRange(2, 12, resultadosConTotales.length, 1).setNumberFormat('#,##0');      // ALCANCE
    hoja.getRange(2, 13, resultadosConTotales.length, 1).setNumberFormat('#,##0');      // CLICS ÚNICOS
    hoja.getRange(2, 14, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');  // COSTO POR CLIC
    hoja.getRange(2, 15, resultadosConTotales.length, 1).setNumberFormat('#,##0');      // MENSAJES
    hoja.getRange(2, 16, resultadosConTotales.length, 1).setNumberFormat('0.00%');      // % MENSAJES
    hoja.getRange(2, 17, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');  // COSTO POR MENSAJE
    hoja.getRange(2, 18, resultadosConTotales.length, 1).setNumberFormat('#,##0');      // # VENTAS
    hoja.getRange(2, 19, resultadosConTotales.length, 1).setNumberFormat('0.00%');      // CVR
    hoja.getRange(2, 20, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');  // COSTO POR COMPRA
    hoja.getRange(2, 21, resultadosConTotales.length, 1).setNumberFormat('$#,##0.00');  // CPM

    // Colorear filas por rango
    const colores = {
      'HOY': '#e8f5e9',
      'AYER': '#fff3e0',
      'LAST_3D': '#e3f2fd',
      'LAST_5D': '#fce4ec',
      'LAST_7D': '#e0f7fa',
      'LAST_30D': '#fff8e1',
      'MAXIMUM': '#f3e5f5'
    };

    const coloresTotales = {
      'HOY': '#a5d6a7',
      'AYER': '#ffcc80',
      'LAST_3D': '#90caf9',
      'LAST_5D': '#f48fb1',
      'LAST_7D': '#80deea',
      'LAST_30D': '#ffe082',
      'MAXIMUM': '#ce93d8'
    };

    resultadosConTotales.forEach((fila, i) => {
      const tipo = fila[0];
      const rango = fila[2];
      const rowRange = hoja.getRange(i + 2, 1, 1, headers.length);

      if (tipo === 'TOTAL') {
        const colorTotal = coloresTotales[rango] || '#bdbdbd';
        rowRange.setBackground(colorTotal)
                .setFontWeight('bold')
                .setBorder(true, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID);
      } else {
        const color = colores[rango] || '#ffffff';
        rowRange.setBackground(color);
      }
    });

    // Agregar filtro automático
    const rangoFiltro = hoja.getRange(1, 1, resultadosConTotales.length + 1, headers.length);
    rangoFiltro.createFilter();

    // Agregar TOTAL GENERAL con fórmulas SUBTOTAL
    const filaTotal = resultadosConTotales.length + 2;
    const primeraFilaDatos = 2;
    const ultimaFilaDatos = resultadosConTotales.length + 1;

    const formulaGasto = `=SUBTOTAL(109,D${primeraFilaDatos}:D${ultimaFilaDatos})`;
    const formulaIGV = `=SUBTOTAL(109,E${primeraFilaDatos}:E${ultimaFilaDatos})`;
    const formulaGastoTotal = `=SUBTOTAL(109,F${primeraFilaDatos}:F${ultimaFilaDatos})`;
    const formulaFactUSD = `=SUBTOTAL(109,G${primeraFilaDatos}:G${ultimaFilaDatos})`;
    const formulaROAS = `=IF(F${filaTotal}>0,G${filaTotal}/F${filaTotal},0)`;
    const formulaUtilidad = `=SUBTOTAL(109,I${primeraFilaDatos}:I${ultimaFilaDatos})`;
    const formulaROI = `=IF(F${filaTotal}>0,I${filaTotal}/F${filaTotal},0)`;
    const formulaImpresiones = `=SUBTOTAL(109,K${primeraFilaDatos}:K${ultimaFilaDatos})`;
    const formulaAlcance = `=SUBTOTAL(109,L${primeraFilaDatos}:L${ultimaFilaDatos})`;
    const formulaClicsUnicos = `=SUBTOTAL(109,M${primeraFilaDatos}:M${ultimaFilaDatos})`;
    const formulaCostoPorClic = `=IF(M${filaTotal}>0,F${filaTotal}/M${filaTotal},0)`;
    const formulaMensajes = `=SUBTOTAL(109,O${primeraFilaDatos}:O${ultimaFilaDatos})`;
    const formulaPorcentajeMensajes = `=IF(M${filaTotal}>0,O${filaTotal}/M${filaTotal},0)`;
    const formulaCostoPorMensaje = `=IF(O${filaTotal}>0,F${filaTotal}/O${filaTotal},0)`;
    const formulaVentas = `=SUBTOTAL(109,R${primeraFilaDatos}:R${ultimaFilaDatos})`;
    const formulaCVR = `=IF(O${filaTotal}>0,R${filaTotal}/O${filaTotal},0)`;
    const formulaCostoPorCompra = `=IF(R${filaTotal}>0,F${filaTotal}/R${filaTotal},0)`;
    const formulaCPM = `=IF(K${filaTotal}>0,(D${filaTotal}/K${filaTotal})*1000,0)`;

    hoja.getRange(filaTotal, 1, 1, headers.length).setValues([[
      '',
      'TOTAL GENERAL',
      '(se actualiza con filtros)',
      formulaGasto,
      formulaIGV,
      formulaGastoTotal,
      formulaFactUSD,
      formulaROAS,
      formulaUtilidad,
      formulaROI,
      formulaImpresiones,
      formulaAlcance,
      formulaClicsUnicos,
      formulaCostoPorClic,
      formulaMensajes,
      formulaPorcentajeMensajes,
      formulaCostoPorMensaje,
      formulaVentas,
      formulaCVR,
      formulaCostoPorCompra,
      formulaCPM
    ]]);

    // Formatear fila de total general
    const rangoTotal = hoja.getRange(filaTotal, 1, 1, headers.length);
    rangoTotal.setBackground('#004d40')
              .setFontColor('white')
              .setFontWeight('bold')
              .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK);

    hoja.getRange(filaTotal, 4).setNumberFormat('$#,##0.00');   // GASTO
    hoja.getRange(filaTotal, 5).setNumberFormat('$#,##0.00');   // IGV
    hoja.getRange(filaTotal, 6).setNumberFormat('$#,##0.00');   // GASTO TOTAL
    hoja.getRange(filaTotal, 7).setNumberFormat('$#,##0.00');   // FACT USD
    hoja.getRange(filaTotal, 8).setNumberFormat('0.00');        // ROAS
    hoja.getRange(filaTotal, 9).setNumberFormat('$#,##0.00');   // UTILIDAD
    hoja.getRange(filaTotal, 10).setNumberFormat('0.00%');      // ROI
    hoja.getRange(filaTotal, 11).setNumberFormat('#,##0');      // IMPRESIONES
    hoja.getRange(filaTotal, 12).setNumberFormat('#,##0');      // ALCANCE
    hoja.getRange(filaTotal, 13).setNumberFormat('#,##0');      // CLICS ÚNICOS
    hoja.getRange(filaTotal, 14).setNumberFormat('$#,##0.00');  // COSTO POR CLIC
    hoja.getRange(filaTotal, 15).setNumberFormat('#,##0');      // MENSAJES
    hoja.getRange(filaTotal, 16).setNumberFormat('0.00%');      // % MENSAJES
    hoja.getRange(filaTotal, 17).setNumberFormat('$#,##0.00');  // COSTO POR MENSAJE
    hoja.getRange(filaTotal, 18).setNumberFormat('#,##0');      // # VENTAS
    hoja.getRange(filaTotal, 19).setNumberFormat('0.00%');      // CVR
    hoja.getRange(filaTotal, 20).setNumberFormat('$#,##0.00');  // COSTO POR COMPRA
    hoja.getRange(filaTotal, 21).setNumberFormat('$#,##0.00');  // CPM

    hoja.autoResizeColumns(1, headers.length);
    hoja.setColumnWidth(1, 60); // Columna TIPO más angosta

    // Resumen
    const numProductos = Object.keys(productosConfig).filter(p => {
      const prod = productosConfig[p];
      const paisProd = prod.pais ? prod.pais.toLowerCase() : '';
      return !filtroPais || !paisProd || paisProd === filtroPais;
    }).length;

    let resumen = `✅ Reporte por Productos completado\n\n`;
    resumen += `🌎 País: ${paisSeleccionado.nombre}\n`;
    resumen += `📦 Productos analizados: ${numProductos}\n`;
    resumen += `📊 Rangos: ${rangos.length}\n`;
    resumen += `📈 Total de filas: ${resultadosConTotales.length}`;

    if (errores.length > 0) {
      resumen += `\n\n⚠️ ${errores.length} error(es) encontrados.`;
    }

    ss.toast('', '', 1); // Cerrar toast
    ui.alert('Reporte Completado', resumen, ui.ButtonSet.OK);
  } else {
    ss.toast('', '', 1);
    let mensaje = '⚠️ No se encontraron datos.';
    if (errores.length > 0) {
      mensaje += '\n\nErrores:\n' + errores.slice(0, 5).join('\n');
    }
    ui.alert('Sin Resultados', mensaje, ui.ButtonSet.OK);
  }
}

/**
 * Función rápida para extraer datos de los últimos 5 días sin prompts
 * Se puede llamar directamente o programar con trigger
 */
function extraerUltimos5Dias() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const token = props.getProperty('META_TOKEN');
  const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

  if (!token || cuentas.length === 0) {
    ui.alert('❌ Error', 'Faltan credenciales de Meta. Configúralas primero.', ui.ButtonSet.OK);
    return;
  }

  let hoja = ss.getSheetByName('📈 Últimos 5 Días');
  if (!hoja) hoja = ss.insertSheet('📈 Últimos 5 Días');
  hoja.clear();

  const headers = ['FECHA EXTRACCIÓN', 'CAMPAÑA', 'CONJUNTO', 'ANUNCIO', 'AD ID', 'GASTO', 'IGV', 'GASTO TOTAL', 'ALCANCE', 'CLICS', 'CPM', 'IMPRESIONES'];
  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#34a853')
      .setFontColor('white')
      .setFontWeight('bold');

  const timeParam = construirTimeRange(5, ss.getSpreadsheetTimeZone());
  const fechaExtraccion = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');

  let resultados = [];

  cuentas.forEach(cuenta => {
    try {
      const datos = extraerDatosDeCuenta(cuenta[0], token, timeParam, '');
      datos.forEach(item => {
        const gasto = parseFloat(item.spend) || 0;
        const igv = gasto * 0.18;
        const gastoTotal = gasto + igv;
        const alcance = parseInt(item.reach) || 0;
        const clics = parseInt(item.clicks) || 0;
        const impresiones = parseInt(item.impressions) || 0;
        const cpm = impresiones > 0 ? (gasto / impresiones) * 1000 : 0;

        resultados.push([
          fechaExtraccion,
          item.campaign_name || '',
          item.adset_name || '',
          item.ad_name || '',
          item.ad_id || '',
          gasto,
          igv,
          gastoTotal,
          alcance,
          clics,
          cpm,
          impresiones
        ]);
      });
    } catch (e) {
      Logger.log(`Error en cuenta ${cuenta[0]}: ${e.message}`);
    }
  });

  if (resultados.length > 0) {
    hoja.getRange(2, 1, resultados.length, headers.length).setValues(resultados);
    hoja.getRange(2, 6, resultados.length, 1).setNumberFormat('$#,##0.00');  // GASTO
    hoja.getRange(2, 7, resultados.length, 1).setNumberFormat('$#,##0.00');  // IGV
    hoja.getRange(2, 8, resultados.length, 1).setNumberFormat('$#,##0.00');  // GASTO TOTAL
    hoja.getRange(2, 9, resultados.length, 2).setNumberFormat('#,##0');      // ALCANCE, CLICS
    hoja.getRange(2, 11, resultados.length, 1).setNumberFormat('$#,##0.00'); // CPM
    hoja.getRange(2, 12, resultados.length, 1).setNumberFormat('#,##0');     // IMPRESIONES
    hoja.autoResizeColumns(1, headers.length);

    ui.alert('✅ Extracción Completada', `Se encontraron ${resultados.length} anuncios en los últimos 5 días.`, ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Sin Datos', 'No se encontraron anuncios en los últimos 5 días.', ui.ButtonSet.OK);
  }
}

/**
 * ==================== FUNCIONES DE FACTURACIÓN ====================
 */

/**
 * Obtiene todas las compras de los spreadsheets vinculados
 * Agrupa por AD ID (POST ID sin "Ads-") y convierte a USD con tasa actual
 * @param {string} filtroPais - Filtro opcional de país (lowercase)
 * @param {Date} fechaInicio - Fecha de inicio para filtrar compras (opcional)
 * @param {Date} fechaFin - Fecha de fin para filtrar compras (opcional)
 * @returns {Object} - Mapa de adId -> { totalUSD, totalLocal, moneda, tasa, cantidadVentas }
 */
function obtenerComprasPorAdId(filtroPais, fechaInicio, fechaFin) {
  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;
  const mapaMonedas = config.MONEDAS_PAIS;
  const tasasActuales = obtenerTasasDeCambio();

  // Columnas de la hoja Compras
  const COL_FECHA = 3;      // Columna C - Fecha
  const COL_VALOR = 4;      // Columna D - Valor
  const COL_POST_ID = 5;    // Columna E - POST ID
  const COL_PAIS = 11;      // Columna K - País

  if (hojas.length === 0) {
    Logger.log('No hay hojas de ventas configuradas');
    return {};
  }

  const limpiar = (t) => t ? t.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() : "";

  // Mapa de adId -> datos de facturación
  const facturacionPorAd = {};
  let totalRegistrosLeidos = 0;
  let registrosSinPostId = 0;
  let registrosFiltradosPorPais = 0;

  hojas.forEach(hojaObj => {
    const id = hojaObj.id || hojaObj;
    const nombreHoja = hojaObj.nombre || '';

    // FILTRAR POR PAÍS: Solo leer spreadsheets que correspondan al país del producto
    if (filtroPais && filtroPais.length > 0) {
      const filtroUpper = filtroPais.toUpperCase();
      const nombreUpper = nombreHoja.toUpperCase();

      // Si el nombre de la hoja no contiene el país del filtro, saltarla
      if (!nombreUpper.includes(filtroUpper)) {
        Logger.log(`Saltando spreadsheet "${nombreHoja}" porque no coincide con país "${filtroPais}"`);
        return;
      }
    }

    try {
      Logger.log(`Leyendo compras de spreadsheet: "${nombreHoja}" (${id})`);
      const ss = SpreadsheetApp.openById(id);
      const hoja = ss.getSheetByName('Compras');
      if (!hoja) {
        Logger.log(`Hoja "Compras" no encontrada en ${id}`);
        return;
      }

      const lastRow = hoja.getLastRow();
      if (lastRow < 2) return;

      // Leer todas las columnas necesarias
      const datos = hoja.getRange(2, 1, lastRow - 1, COL_PAIS).getValues();
      Logger.log(`Leyendo ${datos.length} filas de spreadsheet ${id}`);

      datos.forEach((fila, idx) => {
        totalRegistrosLeidos++;

        const fechaCompra = fila[COL_FECHA - 1];
        const valorLocal = parseFloat(fila[COL_VALOR - 1]) || 0;
        const postIdRaw = fila[COL_POST_ID - 1];
        const paisRaw = fila[COL_PAIS - 1];

        // Solo saltar si no hay POST ID (el valor puede ser 0)
        if (!postIdRaw) {
          registrosSinPostId++;
          return;
        }

        // Filtrar por rango de fechas si se especificó
        if (fechaInicio && fechaFin && fechaCompra) {
          const fechaCompraDate = new Date(fechaCompra);
          if (fechaCompraDate < fechaInicio || fechaCompraDate > fechaFin) {
            return; // Saltar esta compra si está fuera del rango
          }
        }

        // Extraer AD ID del POST ID (puede venir como "Ads-123456" o solo "123456")
        const postIdStr = postIdRaw.toString().trim();

        // Quitar cualquier prefijo "Ads-", "ads-", "ADS-", "Ads_", etc.
        // El regex elimina: "Ads" (case insensitive) + opcional guión/guión bajo/espacio
        let adId = postIdStr.replace(/^ads[\s\-_]*/i, '').trim();

        // Si después de quitar el prefijo no queda nada, usar el original
        if (!adId) {
          adId = postIdStr;
        }

        // Validar que tengamos un AD ID
        if (!adId || adId === '') {
          registrosSinPostId++;
          return;
        }

        const paisLimpio = limpiar(paisRaw);

        // Filtrar por país si se especificó
        if (filtroPais && paisLimpio && !paisLimpio.toLowerCase().includes(filtroPais)) {
          registrosFiltradosPorPais++;
          return;
        }

        // Obtener moneda y tasa del país
        const codigoMoneda = mapaMonedas[paisLimpio] || 'USD';
        const tasa = tasasActuales[codigoMoneda] || 1;
        const montoUSD = valorLocal / tasa;

        // Acumular por AD ID
        if (!facturacionPorAd[adId]) {
          facturacionPorAd[adId] = {
            totalUSD: 0,
            totalLocal: 0,
            moneda: codigoMoneda,
            tasa: tasa,
            cantidadVentas: 0,
            pais: paisLimpio
          };
        }

        facturacionPorAd[adId].totalUSD += montoUSD;
        facturacionPorAd[adId].totalLocal += valorLocal;
        facturacionPorAd[adId].cantidadVentas++;
      });

    } catch (error) {
      Logger.log(`Error leyendo spreadsheet ${id}: ${error.message}`);
    }
  });

  Logger.log(`=== RESUMEN LECTURA COMPRAS ===`);
  Logger.log(`Total registros leídos: ${totalRegistrosLeidos}`);
  Logger.log(`Registros sin POST ID: ${registrosSinPostId}`);
  Logger.log(`Registros filtrados por país: ${registrosFiltradosPorPais}`);
  Logger.log(`AD IDs únicos encontrados: ${Object.keys(facturacionPorAd).length}`);

  // Log de totales por AD ID para debug
  let totalVentas = 0;
  for (const adId in facturacionPorAd) {
    totalVentas += facturacionPorAd[adId].cantidadVentas;
  }
  Logger.log(`Total ventas contabilizadas: ${totalVentas}`);

  return facturacionPorAd;
}

/**
 * Obtiene mensajes agrupados por AD ID desde los spreadsheets de mensajes
 * Similar a obtenerComprasPorAdId pero para la hoja "Mensajes"
 * @param {string} filtroPais - Filtro opcional de país (lowercase)
 * @param {Date} fechaInicio - Fecha de inicio para filtrar mensajes
 * @param {Date} fechaFin - Fecha de fin para filtrar mensajes
 * @returns {Object} - Mapa de adId -> { cantidadMensajes, pais }
 */
function obtenerMensajesPorAdId(filtroPais, fechaInicio, fechaFin) {
  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;

  if (hojas.length === 0) {
    Logger.log('No hay hojas de mensajes configuradas');
    return {};
  }

  // Mapa de adId -> datos de mensajes
  const mensajesPorAd = {};

  // Normalizar fechas del rango (solo año, mes, día)
  const inicioNormalizado = fechaInicio ? new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), fechaInicio.getDate()) : null;
  const finNormalizado = fechaFin ? new Date(fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate(), 23, 59, 59) : null;

  hojas.forEach(hojaObj => {
    const id = hojaObj.id || hojaObj;
    const nombreHoja = hojaObj.nombre || '';

    // FILTRAR POR PAÍS: Solo leer spreadsheets que correspondan al país del producto
    if (filtroPais && filtroPais.length > 0) {
      const filtroUpper = filtroPais.toUpperCase();
      const nombreUpper = nombreHoja.toUpperCase();

      if (!nombreUpper.includes(filtroUpper)) {
        return; // Saltar este spreadsheet
      }
    }

    try {
      const ss = SpreadsheetApp.openById(id);
      const hoja = ss.getSheetByName('Mensajes');

      if (!hoja) return;

      const lastRow = hoja.getLastRow();
      if (lastRow < 2) return;

      // Leer columnas A (fecha) y E (POST ID)
      const COL_FECHA = 1;
      const COL_POST_ID = 5;
      const datos = hoja.getRange(2, 1, lastRow - 1, 5).getValues();

      datos.forEach(fila => {
        const fechaRaw = fila[COL_FECHA - 1];
        const postIdRaw = fila[COL_POST_ID - 1];

        // Validar POST ID
        if (!postIdRaw) return;

        // Parsear fecha
        let fechaObj;
        if (fechaRaw instanceof Date) {
          fechaObj = fechaRaw;
        } else {
          fechaObj = new Date(fechaRaw);
        }

        if (isNaN(fechaObj.getTime())) return;

        // Normalizar fecha (solo año, mes, día)
        const fechaNormalizada = new Date(fechaObj.getFullYear(), fechaObj.getMonth(), fechaObj.getDate());

        // Filtrar por rango de fechas
        if (inicioNormalizado && finNormalizado) {
          if (fechaNormalizada < inicioNormalizado || fechaNormalizada > finNormalizado) {
            return; // Fuera del rango
          }
        }

        // Extraer AD ID del POST ID (igual que las funciones de diagnóstico)
        const postIdStr = postIdRaw.toString().trim();
        const adId = postIdStr.replace(/^ads[\s\-_]*/i, '');

        if (!adId) return;

        // Agregar al mapa
        if (!mensajesPorAd[adId]) {
          mensajesPorAd[adId] = {
            cantidadMensajes: 0
          };
        }

        mensajesPorAd[adId].cantidadMensajes++;
      });

    } catch (error) {
      Logger.log(`Error leyendo spreadsheet ${id} (Mensajes): ${error.message}`);
    }
  });

  return mensajesPorAd;
}

/**
 * DIAGNÓSTICO: Contar mensajes de AYER en PERÚ y mostrar en X1
 */
function contarMensajesAyerPeru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;

  // Calcular fecha de AYER (26 de diciembre de 2025)
  const hoy = new Date();
  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);

  // Normalizar fechas (solo año, mes, día)
  const ayerNormalizado = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate());

  Logger.log(`Buscando mensajes de AYER: ${ayerNormalizado.toISOString()}`);

  // Buscar spreadsheet de PERÚ
  let spreadsheetPeru = null;
  for (const hojaObj of hojas) {
    const nombreHoja = hojaObj.nombre || '';
    if (nombreHoja.toUpperCase().includes('PERU')) {
      spreadsheetPeru = hojaObj;
      Logger.log(`Encontrado spreadsheet de PERÚ: "${nombreHoja}" (${hojaObj.id})`);
      break;
    }
  }

  if (!spreadsheetPeru) {
    Logger.log('❌ No se encontró spreadsheet de PERÚ');
    ss.getRange('X1').setValue('ERROR: No hay spreadsheet PERÚ');
    return;
  }

  try {
    const ssPeru = SpreadsheetApp.openById(spreadsheetPeru.id);
    const hojaMensajes = ssPeru.getSheetByName('Mensajes');

    if (!hojaMensajes) {
      Logger.log('❌ No se encontró hoja "Mensajes" en spreadsheet de PERÚ');
      ss.getRange('X1').setValue('ERROR: No hay hoja Mensajes');
      return;
    }

    const lastRow = hojaMensajes.getLastRow();
    Logger.log(`Total de filas en Mensajes: ${lastRow}`);

    if (lastRow < 2) {
      ss.getRange('X1').setValue(0);
      return;
    }

    // Leer columna A (fecha) de todas las filas
    const datos = hojaMensajes.getRange(2, 1, lastRow - 1, 1).getValues();

    let contadorAyer = 0;
    let totalRegistros = 0;

    Logger.log('Primeras 10 fechas encontradas:');

    datos.forEach((fila, idx) => {
      totalRegistros++;
      const fecha = fila[0];

      if (!fecha) return;

      let fechaObj;
      if (fecha instanceof Date) {
        fechaObj = fecha;
      } else {
        fechaObj = new Date(fecha);
      }

      if (isNaN(fechaObj.getTime())) {
        if (idx < 10) Logger.log(`  Fila ${idx + 2}: Fecha inválida`);
        return;
      }

      // Normalizar (solo año, mes, día)
      const fechaNormalizada = new Date(fechaObj.getFullYear(), fechaObj.getMonth(), fechaObj.getDate());

      if (idx < 10) {
        Logger.log(`  Fila ${idx + 2}: ${fechaObj.toISOString()} → ${fechaNormalizada.toISOString()}`);
      }

      // Comparar si es AYER
      if (fechaNormalizada.getTime() === ayerNormalizado.getTime()) {
        contadorAyer++;
      }
    });

    Logger.log(`Total registros procesados: ${totalRegistros}`);
    Logger.log(`Mensajes de AYER (${ayerNormalizado.toDateString()}): ${contadorAyer}`);

    // Escribir resultado en X1
    ss.getRange('X1').setValue(contadorAyer);
    ss.toast(`Mensajes de AYER en PERÚ: ${contadorAyer}`, 'Diagnóstico', 5);

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ss.getRange('X1').setValue('ERROR: ' + error.message);
  }
}

/**
 * DIAGNÓSTICO: Contar mensajes de un AD ID específico en AYER y mostrar en X2
 */
function contarMensajesPorAdId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;

  const adIdBuscado = '120239870046160380'; // AD ID específico

  // Calcular fecha de AYER
  const hoy = new Date();
  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);
  const ayerNormalizado = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate());

  Logger.log(`Buscando mensajes del AD ID ${adIdBuscado} en fecha AYER: ${ayerNormalizado.toISOString()}`);

  // Buscar spreadsheet de PERÚ
  let spreadsheetPeru = null;
  for (const hojaObj of hojas) {
    const nombreHoja = hojaObj.nombre || '';
    if (nombreHoja.toUpperCase().includes('PERU')) {
      spreadsheetPeru = hojaObj;
      Logger.log(`Encontrado spreadsheet de PERÚ: "${nombreHoja}" (${hojaObj.id})`);
      break;
    }
  }

  if (!spreadsheetPeru) {
    Logger.log('❌ No se encontró spreadsheet de PERÚ');
    ss.getRange('X2').setValue('ERROR: No hay spreadsheet PERÚ');
    return;
  }

  try {
    const ssPeru = SpreadsheetApp.openById(spreadsheetPeru.id);
    const hojaMensajes = ssPeru.getSheetByName('Mensajes');

    if (!hojaMensajes) {
      Logger.log('❌ No se encontró hoja "Mensajes"');
      ss.getRange('X2').setValue('ERROR: No hay hoja Mensajes');
      return;
    }

    const lastRow = hojaMensajes.getLastRow();
    Logger.log(`Total de filas en Mensajes: ${lastRow}`);

    if (lastRow < 2) {
      ss.getRange('X2').setValue(0);
      return;
    }

    // Leer columnas A (fecha) y E (POST ID)
    const COL_FECHA = 1;
    const COL_POST_ID = 5;
    const datos = hojaMensajes.getRange(2, 1, lastRow - 1, 5).getValues();

    let contadorMensajes = 0;
    let totalProcesados = 0;

    Logger.log('Primeros 10 registros procesados:');

    datos.forEach((fila, idx) => {
      totalProcesados++;

      const fecha = fila[COL_FECHA - 1];
      const postIdRaw = fila[COL_POST_ID - 1];

      // Validar fecha
      if (!fecha) return;

      let fechaObj;
      if (fecha instanceof Date) {
        fechaObj = fecha;
      } else {
        fechaObj = new Date(fecha);
      }

      if (isNaN(fechaObj.getTime())) return;

      // Normalizar fecha
      const fechaNormalizada = new Date(fechaObj.getFullYear(), fechaObj.getMonth(), fechaObj.getDate());

      // Verificar si es de AYER
      const esDeAyer = fechaNormalizada.getTime() === ayerNormalizado.getTime();

      // Extraer AD ID del POST ID
      if (!postIdRaw) return;

      const postIdStr = postIdRaw.toString().trim();
      const adIdLimpio = postIdStr.replace(/^ads[\s\-_]*/i, '');

      if (idx < 10) {
        Logger.log(`  Fila ${idx + 2}: POST ID="${postIdStr}" → AD ID="${adIdLimpio}", Fecha=${fechaNormalizada.toDateString()}, Es AYER=${esDeAyer}`);
      }

      // Contar si coincide el AD ID y es de AYER
      if (adIdLimpio === adIdBuscado && esDeAyer) {
        contadorMensajes++;
        if (idx < 10) Logger.log(`    ✅ MATCH! Mensaje del AD ID ${adIdBuscado} de AYER`);
      }
    });

    Logger.log(`Total registros procesados: ${totalProcesados}`);
    Logger.log(`Mensajes del AD ID ${adIdBuscado} en AYER: ${contadorMensajes}`);

    // Escribir resultado en X2
    ss.getRange('X2').setValue(contadorMensajes);
    ss.toast(`Mensajes del AD ID ${adIdBuscado} en AYER: ${contadorMensajes}`, 'Diagnóstico', 5);

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ss.getRange('X2').setValue('ERROR: ' + error.message);
  }
}

/**
 * DIAGNÓSTICO: Contar TODOS los mensajes de un AD ID específico (sin filtro de fecha) y mostrar en X3
 */
function contarMensajesTotalesPorAdId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;

  const adIdBuscado = '120239870046160380'; // AD ID específico

  Logger.log(`Buscando TODOS los mensajes del AD ID ${adIdBuscado} (sin filtro de fecha)`);

  // Buscar spreadsheet de PERÚ
  let spreadsheetPeru = null;
  for (const hojaObj of hojas) {
    const nombreHoja = hojaObj.nombre || '';
    if (nombreHoja.toUpperCase().includes('PERU')) {
      spreadsheetPeru = hojaObj;
      Logger.log(`Encontrado spreadsheet de PERÚ: "${nombreHoja}" (${hojaObj.id})`);
      break;
    }
  }

  if (!spreadsheetPeru) {
    Logger.log('❌ No se encontró spreadsheet de PERÚ');
    ss.getRange('X3').setValue('ERROR: No hay spreadsheet PERÚ');
    return;
  }

  try {
    const ssPeru = SpreadsheetApp.openById(spreadsheetPeru.id);
    const hojaMensajes = ssPeru.getSheetByName('Mensajes');

    if (!hojaMensajes) {
      Logger.log('❌ No se encontró hoja "Mensajes"');
      ss.getRange('X3').setValue('ERROR: No hay hoja Mensajes');
      return;
    }

    const lastRow = hojaMensajes.getLastRow();
    Logger.log(`Total de filas en Mensajes: ${lastRow}`);

    if (lastRow < 2) {
      ss.getRange('X3').setValue(0);
      return;
    }

    // Leer solo columna E (POST ID) - no necesitamos fecha
    const COL_POST_ID = 5;
    const datos = hojaMensajes.getRange(2, COL_POST_ID, lastRow - 1, 1).getValues();

    let contadorMensajes = 0;
    let totalProcesados = 0;

    Logger.log('Primeros 10 POST IDs procesados:');

    datos.forEach((fila, idx) => {
      totalProcesados++;

      const postIdRaw = fila[0];

      // Extraer AD ID del POST ID
      if (!postIdRaw) return;

      const postIdStr = postIdRaw.toString().trim();
      const adIdLimpio = postIdStr.replace(/^ads[\s\-_]*/i, '');

      if (idx < 10) {
        Logger.log(`  Fila ${idx + 2}: POST ID="${postIdStr}" → AD ID="${adIdLimpio}"`);
      }

      // Contar si coincide el AD ID
      if (adIdLimpio === adIdBuscado) {
        contadorMensajes++;
        if (idx < 10) Logger.log(`    ✅ MATCH! Mensaje del AD ID ${adIdBuscado}`);
      }
    });

    Logger.log(`Total registros procesados: ${totalProcesados}`);
    Logger.log(`TOTAL de mensajes del AD ID ${adIdBuscado}: ${contadorMensajes}`);

    // Escribir resultado en X3
    ss.getRange('X3').setValue(contadorMensajes);
    ss.toast(`Total de mensajes del AD ID ${adIdBuscado}: ${contadorMensajes}`, 'Diagnóstico', 5);

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ss.getRange('X3').setValue('ERROR: ' + error.message);
  }
}

/**
 * Diagnóstico completo de extracción de rangos
 * Muestra qué compras se están matcheando para MAXIMUM
 */
function diagnosticarExtraccionRangos() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Calcular fechas del rango MAXIMUM
    const fechasRango = calcularFechasDeRango('MAXIMUM', ss.getSpreadsheetTimeZone());

    Logger.log('=== DIAGNÓSTICO EXTRACCIÓN RANGOS ===');
    Logger.log(`Rango MAXIMUM: ${fechasRango.inicio} a ${fechasRango.fin}`);

    // Obtener compras sin filtro de país
    const comprasPorAdId = obtenerComprasPorAdId('', fechasRango.inicio, fechasRango.fin);

    let mensaje = '🔍 DIAGNÓSTICO EXTRACCIÓN RANGOS (MAXIMUM)\n\n';
    mensaje += `📅 Rango de fechas:\n`;
    mensaje += `  Inicio: ${fechasRango.inicio.toLocaleDateString()}\n`;
    mensaje += `  Fin: ${fechasRango.fin.toLocaleDateString()}\n\n`;
    mensaje += `📊 Total AD IDs encontrados: ${Object.keys(comprasPorAdId).length}\n\n`;

    // Mostrar totales por país
    const totalesPorPais = {};
    let totalUSDGlobal = 0;

    for (const adId in comprasPorAdId) {
      const compra = comprasPorAdId[adId];
      const pais = compra.pais || 'SIN PAÍS';

      if (!totalesPorPais[pais]) {
        totalesPorPais[pais] = {
          totalUSD: 0,
          totalLocal: 0,
          moneda: compra.moneda,
          ventas: 0
        };
      }

      totalesPorPais[pais].totalUSD += compra.totalUSD;
      totalesPorPais[pais].totalLocal += compra.totalLocal;
      totalesPorPais[pais].ventas += compra.cantidadVentas;
      totalUSDGlobal += compra.totalUSD;
    }

    mensaje += '💰 TOTALES POR PAÍS:\n';
    mensaje += '━'.repeat(50) + '\n';

    for (const pais in totalesPorPais) {
      const datos = totalesPorPais[pais];
      mensaje += `\n${pais}:\n`;
      mensaje += `  Ventas: ${datos.ventas}\n`;
      mensaje += `  Total Local: ${datos.totalLocal.toFixed(2)} ${datos.moneda}\n`;
      mensaje += `  Total USD: $${datos.totalUSD.toFixed(2)}\n`;
    }

    mensaje += '\n' + '━'.repeat(50) + '\n';
    mensaje += `\n💵 TOTAL USD GLOBAL: $${totalUSDGlobal.toFixed(2)}\n\n`;

    // Mostrar primeros 10 AD IDs
    mensaje += '📋 PRIMEROS 10 AD IDs CON VENTAS:\n';
    mensaje += '━'.repeat(50) + '\n';

    const adIds = Object.keys(comprasPorAdId).slice(0, 10);
    adIds.forEach((adId, i) => {
      const compra = comprasPorAdId[adId];
      mensaje += `\n${i + 1}. AD ID: "${adId}"\n`;
      mensaje += `   País: ${compra.pais}\n`;
      mensaje += `   Ventas: ${compra.cantidadVentas}\n`;
      mensaje += `   Total: ${compra.totalLocal.toFixed(2)} ${compra.moneda}\n`;
      mensaje += `   Tasa: ${compra.tasa.toFixed(2)}\n`;
      mensaje += `   USD: $${compra.totalUSD.toFixed(2)}\n`;
    });

    // Información adicional sobre POST IDs
    mensaje += '\n\n📝 NOTA IMPORTANTE:\n';
    mensaje += '━'.repeat(50) + '\n';
    mensaje += 'Los POST IDs del spreadsheet se limpian así:\n';
    mensaje += '• "Ads-123456789" → "123456789"\n';
    mensaje += '• "ads-123456789" → "123456789"\n';
    mensaje += '• "123456789" → "123456789"\n';
    mensaje += '\nEstos AD IDs se matchean con los de Meta Ads.';

    Logger.log(mensaje);
    ui.alert(
      '✅ Diagnóstico Completo',
      `Revisa el Log (Extensiones > Apps Script > Ver registros)\n\n` +
      `Resumen:\n` +
      `• Total AD IDs: ${Object.keys(comprasPorAdId).length}\n` +
      `• Total USD: $${totalUSDGlobal.toFixed(2)}\n\n` +
      `Ver log para detalle completo.`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error', e.message + '\n' + e.stack, ui.ButtonSet.OK);
  }
}

/**
 * Diagnóstico completo de por qué un AD ID no aparece en el reporte
 */
function diagnosticarAdIdCompleto() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resp = ui.prompt(
    '🔍 Diagnóstico Completo AD ID',
    'Ingresa el AD ID que NO aparece en "Todos los Rangos":',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const adIdBuscado = resp.getResponseText().trim();
  if (!adIdBuscado) {
    ui.alert('Error', 'Debes ingresar un AD ID', ui.ButtonSet.OK);
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('META_TOKEN');
    const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

    let mensaje = `🔍 DIAGNÓSTICO COMPLETO AD ID: ${adIdBuscado}\n\n`;
    mensaje += '━'.repeat(60) + '\n\n';

    // PASO 1: Verificar en spreadsheet de ventas
    mensaje += '📋 PASO 1: BUSCAR EN SPREADSHEET DE VENTAS\n\n';

    const config = obtenerConfiguracion();
    const hojas = config.VENTAS_BOT.HOJAS;
    const COL_VALOR = 4;
    const COL_POST_ID = 5;
    const COL_PAIS = 11;

    let ventasEncontradas = [];
    let totalVentasSpreadsheet = 0;
    let totalUSDSpreadsheet = 0;

    hojas.forEach(hojaObj => {
      const id = hojaObj.id || hojaObj;
      try {
        const ssVentas = SpreadsheetApp.openById(id);
        const hoja = ssVentas.getSheetByName('Compras');
        if (!hoja) return;

        const lastRow = hoja.getLastRow();
        if (lastRow < 2) return;

        const datos = hoja.getRange(2, 1, lastRow - 1, COL_PAIS).getValues();

        datos.forEach((fila, idx) => {
          const postIdRaw = fila[COL_POST_ID - 1];
          if (!postIdRaw) return;

          const postIdStr = postIdRaw.toString().trim();
          let adId = postIdStr.replace(/^ads[\s\-_]*/i, '').trim();
          if (!adId) adId = postIdStr;

          if (adId === adIdBuscado || postIdStr.includes(adIdBuscado)) {
            const nombre = fila[1];
            const valor = parseFloat(fila[COL_VALOR - 1]) || 0;
            const pais = fila[COL_PAIS - 1];

            ventasEncontradas.push({
              fila: idx + 2,
              nombre: nombre,
              postId: postIdStr,
              adId: adId,
              valor: valor,
              pais: pais
            });

            totalVentasSpreadsheet++;
            totalUSDSpreadsheet += valor;
          }
        });
      } catch (error) {
        Logger.log(`Error leyendo spreadsheet ${id}: ${error.message}`);
      }
    });

    if (ventasEncontradas.length > 0) {
      mensaje += `✅ ENCONTRADO en spreadsheet de ventas\n\n`;
      mensaje += `Total de ventas: ${ventasEncontradas.length}\n`;
      mensaje += `Total valor: ${totalUSDSpreadsheet.toFixed(2)}\n\n`;
      mensaje += 'Detalle de ventas:\n';
      ventasEncontradas.forEach((v, i) => {
        mensaje += `${i + 1}. Fila ${v.fila}: "${v.nombre}"\n`;
        mensaje += `   POST_ID: "${v.postId}" → AD_ID: "${v.adId}"\n`;
        mensaje += `   Valor: ${v.valor.toFixed(2)} | País: ${v.pais}\n\n`;
      });
    } else {
      mensaje += `❌ NO ENCONTRADO en spreadsheet de ventas\n\n`;
    }

    mensaje += '━'.repeat(60) + '\n\n';

    // PASO 2: Verificar en Meta Ads API
    mensaje += '📡 PASO 2: BUSCAR EN META ADS API (RANGO MAXIMUM)\n\n';

    let encontradoEnMeta = false;
    let datosMeta = null;

    if (token && cuentas.length > 0) {
      for (const cuenta of cuentas) {
        try {
          const url = `https://graph.facebook.com/v22.0/${cuenta[0]}/ads?fields=id,name,adset_id,campaign_id,effective_status&date_preset=maximum&access_token=${token}&limit=1000`;
          const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          const result = JSON.parse(response.getContentText());

          if (result.data) {
            for (const ad of result.data) {
              if (ad.id === adIdBuscado) {
                encontradoEnMeta = true;
                datosMeta = ad;
                break;
              }
            }
          }

          if (encontradoEnMeta) break;
        } catch (e) {
          mensaje += `Error consultando cuenta ${cuenta[0]}: ${e.message}\n`;
        }
      }

      if (encontradoEnMeta) {
        mensaje += `✅ ENCONTRADO en Meta Ads API\n\n`;
        mensaje += `AD ID: ${datosMeta.id}\n`;
        mensaje += `Nombre: ${datosMeta.name}\n`;
        mensaje += `Estado: ${datosMeta.effective_status}\n`;
        mensaje += `Campaign ID: ${datosMeta.campaign_id}\n`;
        mensaje += `AdSet ID: ${datosMeta.adset_id}\n\n`;
      } else {
        mensaje += `❌ NO ENCONTRADO en Meta Ads API (rango MAXIMUM)\n\n`;
        mensaje += `⚠️ POSIBLES CAUSAS:\n`;
        mensaje += `• El anuncio fue eliminado de Meta Ads\n`;
        mensaje += `• El anuncio pertenece a otra cuenta publicitaria\n`;
        mensaje += `• El anuncio está fuera del rango MAXIMUM de Meta\n`;
        mensaje += `• El AD ID es incorrecto\n\n`;
      }
    } else {
      mensaje += `⚠️ No se pudo verificar en Meta Ads (falta token o cuentas)\n\n`;
    }

    mensaje += '━'.repeat(60) + '\n\n';

    // CONCLUSIÓN
    mensaje += '🎯 CONCLUSIÓN:\n\n';

    if (ventasEncontradas.length > 0 && encontradoEnMeta) {
      mensaje += `✅ El AD ID existe en ambos lados (spreadsheet y Meta Ads)\n`;
      mensaje += `   Debería aparecer en "Todos los Rangos"\n\n`;
      mensaje += `⚠️ Si NO aparece, el problema puede ser:\n`;
      mensaje += `   • Filtro de país aplicado\n`;
      mensaje += `   • Filtro de producto aplicado\n`;
      mensaje += `   • Error en la consulta de insights de Meta Ads\n`;
    } else if (ventasEncontradas.length > 0 && !encontradoEnMeta) {
      mensaje += `⚠️ El AD ID existe en spreadsheet pero NO en Meta Ads\n`;
      mensaje += `   El anuncio fue eliminado o no está disponible\n`;
      mensaje += `   Por eso NO aparece en "Todos los Rangos"\n`;
    } else if (ventasEncontradas.length === 0 && encontradoEnMeta) {
      mensaje += `⚠️ El AD ID existe en Meta Ads pero NO tiene ventas\n`;
      mensaje += `   Verificar que el POST ID en el spreadsheet sea correcto\n`;
    } else {
      mensaje += `❌ El AD ID NO existe en ningún lado\n`;
      mensaje += `   Verificar que el AD ID sea correcto\n`;
    }

    Logger.log(mensaje);

    // Mensaje corto para alert
    let alertMsg = `AD ID: ${adIdBuscado}\n\n`;
    alertMsg += `Spreadsheet: ${ventasEncontradas.length > 0 ? '✅ Encontrado' : '❌ No encontrado'}\n`;
    alertMsg += `Meta Ads: ${encontradoEnMeta ? '✅ Encontrado' : '❌ No encontrado'}\n\n`;

    if (ventasEncontradas.length > 0) {
      alertMsg += `Ventas: ${ventasEncontradas.length}\n`;
      alertMsg += `Total: ${totalUSDSpreadsheet.toFixed(2)}\n\n`;
    }

    alertMsg += 'Ver log completo: Extensiones > Apps Script > Ver registros';

    ui.alert('🔍 Diagnóstico Completo', alertMsg, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', e.message + '\n' + e.stack, ui.ButtonSet.OK);
  }
}

/**
 * Función de diagnóstico para verificar compras de un AD ID específico
 * Ejecutar desde el editor de Apps Script
 */
function diagnosticarAdId() {
  const ui = SpreadsheetApp.getUi();

  const resp = ui.prompt('🔍 Diagnóstico AD ID', 'Ingresa el AD ID a verificar:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const adIdBuscado = resp.getResponseText().trim();
  if (!adIdBuscado) {
    ui.alert('Error', 'Debes ingresar un AD ID', ui.ButtonSet.OK);
    return;
  }

  const config = obtenerConfiguracion();
  const hojas = config.VENTAS_BOT.HOJAS;

  const COL_VALOR = 4;
  const COL_POST_ID = 5;
  const COL_PAIS = 11;

  let resultados = [];
  let totalEncontradas = 0;

  hojas.forEach(hojaObj => {
    const id = hojaObj.id || hojaObj;

    try {
      const ss = SpreadsheetApp.openById(id);
      const hoja = ss.getSheetByName('Compras');
      if (!hoja) return;

      const lastRow = hoja.getLastRow();
      if (lastRow < 2) return;

      const datos = hoja.getRange(2, 1, lastRow - 1, COL_PAIS).getValues();

      datos.forEach((fila, idx) => {
        const postIdRaw = fila[COL_POST_ID - 1];
        if (!postIdRaw) return;

        const postIdStr = postIdRaw.toString().trim();
        // Usar el mismo regex mejorado que en obtenerComprasPorAdId
        let adId = postIdStr.replace(/^ads[\s\-_]*/i, '').trim();
        if (!adId) adId = postIdStr;

        // Buscar coincidencia exacta o parcial
        if (adId === adIdBuscado || postIdStr.includes(adIdBuscado) || adId.includes(adIdBuscado)) {
          totalEncontradas++;
          const nombre = fila[1];  // Columna B - Nombre
          const valor = fila[COL_VALOR - 1];
          const pais = fila[COL_PAIS - 1];
          resultados.push(`Fila ${idx + 2}: "${nombre}" | POST_ID="${postIdStr}" → AD_ID="${adId}" | Valor=${valor} | País=${pais}`);
        }
      });

    } catch (error) {
      resultados.push(`Error en spreadsheet ${id}: ${error.message}`);
    }
  });

  let mensaje = `=== DIAGNÓSTICO AD ID: ${adIdBuscado} ===\n\n`;
  mensaje += `Total compras encontradas: ${totalEncontradas}\n\n`;

  if (resultados.length > 0) {
    mensaje += `Detalle (primeras 20):\n`;
    mensaje += resultados.slice(0, 20).join('\n');
    if (resultados.length > 20) {
      mensaje += `\n... y ${resultados.length - 20} más`;
    }
  } else {
    mensaje += `No se encontraron compras con este AD ID.\n\n`;
    mensaje += `Verifica:\n`;
    mensaje += `- El formato del POST ID en la hoja (ej: "Ads-${adIdBuscado}")\n`;
    mensaje += `- Que la columna E contenga el POST ID`;
  }

  // Mostrar en un alert (limitado) y en el log completo
  Logger.log(mensaje);
  Logger.log('\n=== TODOS LOS RESULTADOS ===');
  resultados.forEach(r => Logger.log(r));

  ui.alert('Diagnóstico', mensaje.substring(0, 1500), ui.ButtonSet.OK);
}
