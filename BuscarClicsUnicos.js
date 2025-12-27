/**
 * Extrae los clics únicos en el enlace (unique_inline_link_clicks)
 * desde Meta Ads API para un AD ID específico
 * y lo coloca en la celda X1
 */
function extraerClicsUnicosEnlaceDesdeMetaParaAdId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();

  const token = props.getProperty('META_TOKEN');
  const cuentas = JSON.parse(props.getProperty('META_CUENTAS') || '[]');

  if (!token) {
    ss.getRange('X1').setValue('ERROR: No hay token');
    Logger.log('❌ No hay token de Meta configurado');
    return;
  }

  if (cuentas.length === 0) {
    ss.getRange('X1').setValue('ERROR: No hay cuentas');
    Logger.log('❌ No hay cuentas configuradas');
    return;
  }

  const adIdBuscado = '120240489569360635';

  Logger.log(`🔍 Buscando clics únicos en el enlace para AD ID: ${adIdBuscado}`);
  Logger.log(`📅 Rango: Últimos 30 días (desde hace 30 días hasta ayer)`);

  // Calcular fechas: desde hace 30 días hasta ayer
  const hoy = new Date();
  const ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);

  const hace30Dias = new Date(ayer);
  hace30Dias.setDate(ayer.getDate() - 29); // 30 días en total incluyendo ayer

  const timezone = ss.getSpreadsheetTimeZone();
  const formatearFecha = (fecha) => Utilities.formatDate(fecha, timezone, 'yyyy-MM-dd');

  const fechaDesde = formatearFecha(hace30Dias);
  const fechaHasta = formatearFecha(ayer);

  Logger.log(`   Desde: ${fechaDesde}`);
  Logger.log(`   Hasta: ${fechaHasta}`);

  // El campo correcto es: unique_inline_link_clicks
  // Este campo representa los clics únicos en enlaces del anuncio
  const campos = 'unique_inline_link_clicks,inline_link_clicks';

  let encontrado = false;
  let clicsUnicosEnlace = 0;

  // Buscar en todas las cuentas publicitarias
  for (const cuenta of cuentas) {
    const accId = cuenta[0];

    try {
      Logger.log(`📡 Consultando cuenta ${accId}...`);

      // Construir time_range para últimos 30 días
      const timeRange = JSON.stringify({
        since: fechaDesde,
        until: fechaHasta
      });

      // Hacer consulta directa al AD específico para obtener insights
      const url = `https://graph.facebook.com/v22.0/${adIdBuscado}/insights?fields=${campos}&time_range=${encodeURIComponent(timeRange)}&access_token=${token}`;

      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        const json = JSON.parse(responseText);
        const datos = json.data || [];

        if (datos.length > 0) {
          encontrado = true;
          const insight = datos[0];

          // Extraer el campo unique_inline_link_clicks
          clicsUnicosEnlace = parseInt(insight.unique_inline_link_clicks) || 0;
          const clicsEnlace = parseInt(insight.inline_link_clicks) || 0;

          Logger.log(`✅ AD ID encontrado en cuenta ${accId}`);
          Logger.log(`   Clics únicos en enlace: ${clicsUnicosEnlace}`);
          Logger.log(`   Clics en enlace (totales): ${clicsEnlace}`);
          Logger.log(`   Datos completos: ${JSON.stringify(insight)}`);

          // Escribir en X1
          ss.getRange('X1').setValue(clicsUnicosEnlace);
          ss.toast(`Clics únicos en enlace: ${clicsUnicosEnlace}`, '✅ Completado', 5);

          break; // Ya encontramos el anuncio
        }
      } else if (responseCode === 400) {
        // Error 400 puede significar que el anuncio no pertenece a esta cuenta
        Logger.log(`⚠️ AD ID no pertenece a la cuenta ${accId}`);
      } else {
        Logger.log(`⚠️ Error ${responseCode} consultando cuenta ${accId}: ${responseText}`);
      }

    } catch (e) {
      Logger.log(`❌ Error consultando cuenta ${accId}: ${e.message}`);
    }
  }

  if (!encontrado) {
    Logger.log(`❌ No se encontró el AD ID ${adIdBuscado} en ninguna cuenta`);
    ss.getRange('X1').setValue('NO ENCONTRADO');
    ss.toast(`AD ID ${adIdBuscado} no encontrado en Meta Ads`, '⚠️ Advertencia', 5);
  }
}

/**
 * Función de diagnóstico: Muestra TODOS los campos de clics disponibles
 * para entender las diferencias
 */
function diagnosticarCamposClicsMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const token = props.getProperty('META_TOKEN');
  const adIdBuscado = '120240489569360635';

  if (!token) {
    Logger.log('❌ No hay token');
    return;
  }

  // Solicitar TODOS los campos relacionados con clics
  const campos = 'clicks,unique_clicks,inline_link_clicks,unique_inline_link_clicks,outbound_clicks,unique_outbound_clicks,inline_link_click_ctr,unique_link_clicks_ctr';

  try {
    const url = `https://graph.facebook.com/v22.0/${adIdBuscado}/insights?fields=${campos}&date_preset=maximum&access_token=${token}`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const datos = json.data || [];

    if (datos.length > 0) {
      const insight = datos[0];

      Logger.log('=== DIAGNÓSTICO CAMPOS DE CLICS ===');
      Logger.log(`AD ID: ${adIdBuscado}\n`);
      Logger.log(`clicks (todos los clics): ${insight.clicks || 0}`);
      Logger.log(`unique_clicks (clics únicos - todos): ${insight.unique_clicks || 0}`);
      Logger.log(`inline_link_clicks (clics en enlace): ${insight.inline_link_clicks || 0}`);
      Logger.log(`unique_inline_link_clicks (CLICS ÚNICOS EN ENLACE): ${insight.unique_inline_link_clicks || 0}`);
      Logger.log(`outbound_clicks (clics salientes): ${insight.outbound_clicks || 0}`);
      Logger.log(`unique_outbound_clicks (clics salientes únicos): ${insight.unique_outbound_clicks || 0}`);
      Logger.log(`\nCampo correcto a usar: unique_inline_link_clicks = ${insight.unique_inline_link_clicks || 0}`);

      // Escribir todos los valores en celdas para comparación
      ss.getRange('X1').setValue(parseInt(insight.unique_inline_link_clicks) || 0);
      ss.getRange('X2').setValue(parseInt(insight.unique_clicks) || 0);
      ss.getRange('X3').setValue(parseInt(insight.inline_link_clicks) || 0);

      ss.getRange('W1').setValue('unique_inline_link_clicks');
      ss.getRange('W2').setValue('unique_clicks');
      ss.getRange('W3').setValue('inline_link_clicks');

    } else {
      Logger.log('❌ No se encontraron datos para este AD ID');
    }

  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
}
