/**
 * MÓDULO DE EXTRACCIÓN DE CREATIVOS
 * Extrae datos de anuncios desde Meta Ads Manager
 */

/**
 * Función principal para extraer gasto por anuncio
 */
function extraerGastoPorAnuncio() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Pedir Producto
  const respProd = ui.prompt('🔍 Producto', 'Introduce el nombre (o déjalo vacío para todos):', ui.ButtonSet.OK_CANCEL);
  if (respProd.getSelectedButton() !== ui.Button.OK) return;
  const nombreProducto = respProd.getResponseText().trim();

  // 2. Pedir Rango
  const respFecha = ui.prompt(
    '📅 Rango de Fechas',
    'Opciones válidas:\n• today\n• yesterday\n• last_3d\n• last_7d\n• last_14d\n• last_30d\n• last_5d (personalizado)\n• maximum (histórico completo)',
    ui.ButtonSet.OK_CANCEL
  );
  if (respFecha.getSelectedButton() !== ui.Button.OK) return;
  const rango = respFecha.getResponseText().trim().toLowerCase();

  // 3. Validar credenciales
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_TOKEN');
  const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

  if (!token) {
    ui.alert('❌ Error', 'No hay token de Meta configurado. Ve a Configuración > Configurar Token de Meta.', ui.ButtonSet.OK);
    return;
  }

  if (cuentas.length === 0) {
    ui.alert('❌ Error', 'No hay cuentas publicitarias configuradas. Ve a Configuración > Configurar Cuentas Publicitarias.', ui.ButtonSet.OK);
    return;
  }

  // 4. Preparar hoja de resultados
  let hoja = ss.getSheetByName('📈 Gasto por Anuncio');
  if (!hoja) hoja = ss.insertSheet('📈 Gasto por Anuncio');
  hoja.clear();

  const headers = ['FILTRO', 'RANGO', 'CAMPAÑA', 'CONJUNTO', 'ANUNCIO', 'AD ID', 'GASTO', 'IGV', 'GASTO TOTAL', 'ALCANCE', 'CLICS', 'CPM', 'IMPRESIONES'];
  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');

  // 5. Construir parámetro de tiempo
  const timeParam = construirParametroTiempo(rango, ss.getSpreadsheetTimeZone());

  if (!timeParam) {
    ui.alert('❌ Error', 'Rango de fechas no válido: ' + rango, ui.ButtonSet.OK);
    return;
  }

  // 6. Extracción de datos
  let resultados = [];
  let errores = [];

  cuentas.forEach(cuenta => {
    const accId = cuenta[0];
    try {
      const datos = extraerDatosDeCuenta(accId, token, timeParam, nombreProducto);
      resultados = resultados.concat(datos.map(item => {
        const gasto = parseFloat(item.spend) || 0;
        const igv = gasto * 0.18;
        const gastoTotal = gasto + igv;
        const alcance = parseInt(item.reach) || 0;
        const clics = parseInt(item.clicks) || 0;
        const impresiones = parseInt(item.impressions) || 0;
        const cpm = impresiones > 0 ? (gasto / impresiones) * 1000 : 0;

        return [
          nombreProducto || 'TODOS',
          rango.toUpperCase(),
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
        ];
      }));
    } catch (e) {
      errores.push(`Cuenta ${accId}: ${e.message}`);
      Logger.log(`Error en cuenta ${accId}: ${e.message}`);
    }
  });

  // 7. Escribir resultados
  if (resultados.length > 0) {
    hoja.getRange(2, 1, resultados.length, headers.length).setValues(resultados);

    // Formatear columnas numéricas
    hoja.getRange(2, 7, resultados.length, 1).setNumberFormat('$#,##0.00');  // GASTO
    hoja.getRange(2, 8, resultados.length, 1).setNumberFormat('$#,##0.00');  // IGV
    hoja.getRange(2, 9, resultados.length, 1).setNumberFormat('$#,##0.00');  // GASTO TOTAL
    hoja.getRange(2, 10, resultados.length, 2).setNumberFormat('#,##0');     // ALCANCE, CLICS
    hoja.getRange(2, 12, resultados.length, 1).setNumberFormat('$#,##0.00'); // CPM
    hoja.getRange(2, 13, resultados.length, 1).setNumberFormat('#,##0');     // IMPRESIONES

    // Ajustar anchos de columna
    hoja.autoResizeColumns(1, headers.length);

    let mensaje = `✅ Éxito: ${resultados.length} anuncios encontrados.`;
    if (errores.length > 0) {
      mensaje += `\n\n⚠️ Errores en ${errores.length} cuenta(s).`;
    }
    ui.alert('Extracción Completada', mensaje, ui.ButtonSet.OK);
  } else {
    let mensaje = '⚠️ No se encontraron datos para los criterios especificados.';
    if (errores.length > 0) {
      mensaje += '\n\nErrores:\n' + errores.join('\n');
    }
    ui.alert('Sin Resultados', mensaje, ui.ButtonSet.OK);
  }
}

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
  const url = `https://graph.facebook.com/v21.0/${accountId}/insights?level=ad&fields=${campos}&${timeParam}&limit=500&access_token=${token}`;

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
    'TIPO', 'RANGO', 'CAMPAÑA', 'CONJUNTO', 'ANUNCIO', 'AD ID',
    'GASTO', 'IGV', 'GASTO TOTAL', 'FACT USD', 'ROAS', 'UTILIDAD', 'ROI',
    'ALCANCE', 'CLICS', 'CPM', '# VENTAS', 'T.C.', 'IMPRESIONES'
  ];
  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#9c27b0')
      .setFontColor('white')
      .setFontWeight('bold');

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
      totales: { gasto: 0, igv: 0, gastoTotal: 0, alcance: 0, clics: 0, factUSD: 0, numVentas: 0, utilidad: 0, impresiones: 0 }
    };

    // Calcular fechas del rango y obtener compras filtradas por ese rango
    const fechasRango = calcularFechasDeRango(rangoConfig.nombre, ss.getSpreadsheetTimeZone());
    const comprasPorAdId = obtenerComprasPorAdId(filtroPais, fechasRango.inicio, fechasRango.fin);
    Logger.log(`${rangoConfig.nombre}: Compras encontradas para ${Object.keys(comprasPorAdId).length} AD IDs`);

    cuentas.forEach(cuenta => {
      try {
        let datos = extraerDatosDeCuenta(cuenta[0], token, rangoConfig.param, filtroProducto);

        // Filtrar por país si se seleccionó uno específico
        if (filtroPais) {
          datos = datos.filter(item =>
            item.campaign_name && item.campaign_name.toLowerCase().includes(filtroPais)
          );
        }

        datos.forEach(item => {
          const gasto = parseFloat(item.spend) || 0;
          const igv = gasto * 0.18;
          const gastoTotal = gasto + igv;
          const alcance = parseInt(item.reach) || 0;
          const clics = parseInt(item.clicks) || 0;
          const impresiones = parseInt(item.impressions) || 0;
          const cpm = impresiones > 0 ? (gasto / impresiones) * 1000 : 0;

          // Buscar facturación por AD ID para este rango
          const adId = item.ad_id || '';
          let factUSD = 0;
          let numVentas = 0;
          let tasaCambio = '';

          // Vincular compras del rango actual
          if (adId && comprasPorAdId[adId]) {
            const compra = comprasPorAdId[adId];
            factUSD = compra.totalUSD;
            numVentas = compra.cantidadVentas;
            tasaCambio = compra.tasa;
          }

          const roas = gastoTotal > 0 ? factUSD / gastoTotal : 0;
          const utilidad = factUSD - gastoTotal;
          const roi = gastoTotal > 0 ? utilidad / gastoTotal : 0;

          datosPorRango[rangoConfig.nombre].filas.push([
            'DATO',
            rangoConfig.nombre,
            item.campaign_name || '',
            item.adset_name || '',
            item.ad_name || '',
            adId,
            gasto,
            igv,
            gastoTotal,
            factUSD,
            roas,
            utilidad,
            roi,
            alcance,
            clics,
            cpm,
            numVentas,
            tasaCambio,
            impresiones
          ]);

          // Acumular totales
          datosPorRango[rangoConfig.nombre].totales.gasto += gasto;
          datosPorRango[rangoConfig.nombre].totales.igv += igv;
          datosPorRango[rangoConfig.nombre].totales.gastoTotal += gastoTotal;
          datosPorRango[rangoConfig.nombre].totales.alcance += alcance;
          datosPorRango[rangoConfig.nombre].totales.clics += clics;
          datosPorRango[rangoConfig.nombre].totales.factUSD += factUSD;
          datosPorRango[rangoConfig.nombre].totales.numVentas += numVentas;
          datosPorRango[rangoConfig.nombre].totales.utilidad += utilidad;
          datosPorRango[rangoConfig.nombre].totales.impresiones += impresiones;

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

      // Calcular CPM, ROAS y ROI de los totales
      const totales = datos.totales;
      const cpmTotal = totales.impresiones > 0 ? (totales.gasto / totales.impresiones) * 1000 : 0;
      const roasTotal = totales.gastoTotal > 0 ? totales.factUSD / totales.gastoTotal : 0;
      const roiTotal = totales.gastoTotal > 0 ? totales.utilidad / totales.gastoTotal : 0;

      // Agregar fila de totales
      todosLosResultados.push([
        'TOTAL',
        rangoConfig.nombre,
        `(${datos.filas.length} anuncios)`,
        '', '', '',
        totales.gasto,
        totales.igv,
        totales.gastoTotal,
        totales.factUSD,
        roasTotal,
        totales.utilidad,
        roiTotal,
        totales.alcance,
        totales.clics,
        cpmTotal,
        totales.numVentas,
        '',  // T.C. vacío en totales
        totales.impresiones
      ]);
    }
  });

  // Escribir resultados
  if (todosLosResultados.length > 0) {
    hoja.getRange(2, 1, todosLosResultados.length, headers.length).setValues(todosLosResultados);

    // Formatear columnas
    hoja.getRange(2, 7, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // GASTO
    hoja.getRange(2, 8, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // IGV
    hoja.getRange(2, 9, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');   // GASTO TOTAL
    hoja.getRange(2, 10, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');  // FACT USD
    hoja.getRange(2, 11, todosLosResultados.length, 1).setNumberFormat('0.00');       // ROAS
    hoja.getRange(2, 12, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');  // UTILIDAD
    hoja.getRange(2, 13, todosLosResultados.length, 1).setNumberFormat('0.00%');      // ROI
    hoja.getRange(2, 14, todosLosResultados.length, 2).setNumberFormat('#,##0');      // ALCANCE, CLICS
    hoja.getRange(2, 16, todosLosResultados.length, 1).setNumberFormat('$#,##0.00');  // CPM
    hoja.getRange(2, 17, todosLosResultados.length, 1).setNumberFormat('#,##0');      // # VENTAS
    hoja.getRange(2, 18, todosLosResultados.length, 1).setNumberFormat('0.00');       // T.C.
    hoja.getRange(2, 19, todosLosResultados.length, 1).setNumberFormat('#,##0');      // IMPRESIONES

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

    try {
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

        // Extraer AD ID: quitar "Ads-" del POST ID
        const postIdStr = postIdRaw.toString().trim();
        // Quitar cualquier prefijo "Ads-" o "ads-" o similar
        const adId = postIdStr.replace(/^Ads[-_]?/i, '').trim();

        if (!adId) {
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
        const adId = postIdStr.replace(/^Ads[-_]?/i, '').trim();

        // Buscar coincidencia exacta o parcial
        if (adId === adIdBuscado || postIdStr.includes(adIdBuscado)) {
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
