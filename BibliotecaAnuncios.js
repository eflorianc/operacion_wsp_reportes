/**
 * ==================== BIBLIOTECA DE ANUNCIOS DE META ====================
 * Módulo independiente para consultar la Biblioteca de Anuncios de Meta
 * y extraer información de gasto y datos públicos de anuncios
 */

/**
 * FUNCIÓN ALTERNATIVA RECOMENDADA
 * Obtiene el gasto directamente del AD ID usando Insights API
 * (No requiere permisos especiales de Biblioteca de Anuncios)
 */
function obtenerGastoPorAdId() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Solicitar el AD ID
  const respId = ui.prompt(
    '🔍 Obtener Gasto por AD ID',
    'Ingresa el AD ID del anuncio:\n(Puedes encontrarlo en Ads Manager o en tus reportes)',
    ui.ButtonSet.OK_CANCEL
  );

  if (respId.getSelectedButton() !== ui.Button.OK) return;
  const adId = respId.getResponseText().trim();

  if (!adId) {
    ui.alert('Error', 'Debes ingresar un AD ID', ui.ButtonSet.OK);
    return;
  }

  // 2. Obtener token
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_TOKEN');

  if (!token) {
    ui.alert('❌ Error', 'No hay token de Meta configurado.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Consultando datos del anuncio...', '⏳ Procesando', -1);

  try {
    // Consultar insights del anuncio
    const campos = 'spend,impressions,reach,unique_inline_link_clicks,ad_name,campaign_name,adset_name';
    const url = `https://graph.facebook.com/v22.0/${adId}/insights?fields=${campos}&date_preset=maximum&access_token=${token}`;

    Logger.log(`Consultando AD ID: ${adId}`);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    ss.toast('', '', 1);

    if (responseCode !== 200) {
      const errorData = JSON.parse(responseText);
      ui.alert('❌ Error', `No se pudo obtener datos del AD ID ${adId}\n\n${errorData.error.message}`, ui.ButtonSet.OK);
      return;
    }

    const json = JSON.parse(responseText);
    const datos = json.data || [];

    if (datos.length === 0) {
      ui.alert('Sin Datos', `No se encontraron datos para el AD ID ${adId}\n\nPosibles causas:\n• El AD ID no existe\n• El anuncio pertenece a otra cuenta\n• No hay datos de insights disponibles`, ui.ButtonSet.OK);
      return;
    }

    const insight = datos[0];

    // Mostrar resultado
    let mensaje = '💰 GASTO DEL ANUNCIO\n';
    mensaje += '━'.repeat(50) + '\n\n';
    mensaje += `🆔 AD ID: ${adId}\n\n`;
    mensaje += `📊 Campaña: ${insight.campaign_name || 'N/A'}\n`;
    mensaje += `📁 Conjunto: ${insight.adset_name || 'N/A'}\n`;
    mensaje += `📝 Anuncio: ${insight.ad_name || 'N/A'}\n\n`;
    mensaje += `💵 GASTO TOTAL: $${parseFloat(insight.spend || 0).toFixed(2)} USD\n`;
    mensaje += `👁️ Impresiones: ${parseInt(insight.impressions || 0).toLocaleString()}\n`;
    mensaje += `📍 Alcance: ${parseInt(insight.reach || 0).toLocaleString()}\n`;
    mensaje += `🔗 Clics únicos en enlace: ${parseInt(insight.unique_inline_link_clicks || 0).toLocaleString()}\n\n`;
    mensaje += `📅 Período: Histórico completo (maximum)\n`;

    Logger.log('=== DATOS DEL ANUNCIO ===');
    Logger.log(JSON.stringify(insight, null, 2));

    ui.alert('✅ Datos Obtenidos', mensaje, ui.ButtonSet.OK);

    // Escribir en hoja
    escribirGastoEnHoja(adId, insight);

  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
    Logger.log(`Error: ${e.message}`);
  }
}

/**
 * Escribe el gasto en una hoja
 */
function escribirGastoEnHoja(adId, insight) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('💰 Gastos por AD ID');

  if (!hoja) {
    hoja = ss.insertSheet('💰 Gastos por AD ID');
    hoja.setTabColor('#2e7d32');

    const headers = [
      'FECHA CONSULTA', 'AD ID', 'CAMPAÑA', 'CONJUNTO', 'ANUNCIO',
      'GASTO USD', 'IMPRESIONES', 'ALCANCE', 'CLICS ÚNICOS'
    ];

    hoja.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground('#2e7d32')
         .setFontColor('white')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }

  const nuevaFila = [
    new Date(),
    adId,
    insight.campaign_name || 'N/A',
    insight.adset_name || 'N/A',
    insight.ad_name || 'N/A',
    parseFloat(insight.spend || 0),
    parseInt(insight.impressions || 0),
    parseInt(insight.reach || 0),
    parseInt(insight.unique_inline_link_clicks || 0)
  ];

  const ultimaFila = hoja.getLastRow();
  hoja.getRange(ultimaFila + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  // Formatear
  hoja.getRange(ultimaFila + 1, 6).setNumberFormat('$#,##0.00');
  hoja.getRange(ultimaFila + 1, 7, 1, 3).setNumberFormat('#,##0');

  hoja.autoResizeColumns(1, nuevaFila.length);

  ss.toast('Guardado en "💰 Gastos por AD ID"', '✅', 3);
}

/**
 * Función principal para buscar un anuncio en la Biblioteca de Anuncios
 * y extraer su gasto
 * NOTA: Requiere permisos especiales - usa obtenerGastoPorAdId() en su lugar
 */
function buscarGastoEnBibliotecaAnuncios() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Solicitar el ID del anuncio de la biblioteca
  const respId = ui.prompt(
    '🔍 Buscar en Biblioteca de Anuncios',
    'Ingresa el ID único del anuncio de la Biblioteca de Meta:\n(Puedes encontrarlo en la URL de facebook.com/ads/library)',
    ui.ButtonSet.OK_CANCEL
  );

  if (respId.getSelectedButton() !== ui.Button.OK) return;
  const adArchiveId = respId.getResponseText().trim();

  if (!adArchiveId) {
    ui.alert('Error', 'Debes ingresar un ID de anuncio', ui.ButtonSet.OK);
    return;
  }

  // 2. Solicitar el país (requerido por la API)
  const respPais = ui.prompt(
    '🌍 Seleccionar País',
    'Ingresa el código del país (2 letras):\n\nEjemplos:\n• PE (Perú)\n• MX (México)\n• CO (Colombia)\n• AR (Argentina)\n• US (Estados Unidos)',
    ui.ButtonSet.OK_CANCEL
  );

  if (respPais.getSelectedButton() !== ui.Button.OK) return;
  const paisCode = respPais.getResponseText().trim().toUpperCase();

  if (paisCode.length !== 2) {
    ui.alert('Error', 'El código de país debe ser de 2 letras (ej: PE, MX, CO)', ui.ButtonSet.OK);
    return;
  }

  // 3. Obtener token
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_TOKEN');

  if (!token) {
    ui.alert('❌ Error', 'No hay token de Meta configurado. Ve a Configuración > Configurar Token de Meta.', ui.ButtonSet.OK);
    return;
  }

  // 4. Mostrar mensaje de carga
  ss.toast('Consultando Biblioteca de Anuncios de Meta...', '⏳ Procesando', -1);

  try {
    const resultado = consultarBibliotecaAnuncios(adArchiveId, paisCode, token);

    if (resultado.error) {
      ss.toast('', '', 1);
      ui.alert('❌ Error', resultado.error, ui.ButtonSet.OK);
      return;
    }

    if (!resultado.encontrado) {
      ss.toast('', '', 1);
      ui.alert(
        '⚠️ No Encontrado',
        `No se encontró el anuncio con ID "${adArchiveId}" en la Biblioteca de Anuncios de ${paisCode}.\n\n` +
        `Verifica que:\n` +
        `• El ID sea correcto\n` +
        `• El país sea correcto\n` +
        `• El anuncio esté activo o haya estado activo recientemente`,
        ui.ButtonSet.OK
      );
      return;
    }

    // 5. Mostrar resultado
    ss.toast('', '', 1);
    mostrarResultadoBiblioteca(resultado, ui);

  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('Error', `Error al consultar la biblioteca: ${e.message}`, ui.ButtonSet.OK);
    Logger.log(`Error: ${e.message}\n${e.stack}`);
  }
}

/**
 * Consulta la Biblioteca de Anuncios de Meta
 * @param {string} adArchiveId - ID del anuncio en la biblioteca
 * @param {string} paisCode - Código ISO del país (2 letras)
 * @param {string} token - Token de acceso de Meta
 * @returns {Object} - Resultado con datos del anuncio
 */
function consultarBibliotecaAnuncios(adArchiveId, paisCode, token) {
  // IMPORTANTE: La API de Biblioteca de Anuncios permite ver datos PÚBLICOS
  // de anuncios de cualquier anunciante (incluyendo competencia)

  // SOLO campos absolutamente básicos
  // Usar el mínimo indispensable para evitar error 500
  const campos = 'id,ad_snapshot_url,page_name';

  // NOTA IMPORTANTE: La API de Biblioteca de Anuncios tiene limitaciones
  // severas para anuncios comerciales. Muchos campos solo están disponibles
  // para anuncios políticos/sociales.

  // CAMBIO IMPORTANTE: Usar búsqueda más específica
  // En lugar de search_terms, usar search_page_ids puede ser más confiable
  // Pero primero intentemos con un formato más simple de la query

  const params = {
    access_token: token,
    ad_reached_countries: JSON.stringify([paisCode]),
    search_terms: adArchiveId,
    ad_type: 'ALL',  // Buscar en todos los tipos de anuncios
    fields: campos,
    limit: '100'
  };

  // Construir query string
  const queryString = Object.keys(params)
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');

  const url = `https://graph.facebook.com/v22.0/ads_archive?${queryString}`;

  Logger.log(`=== CONSULTA A BIBLIOTECA DE ANUNCIOS ===`);
  Logger.log(`URL: ${url.replace(token, 'TOKEN_OCULTO')}`);
  Logger.log(`ID buscado: ${adArchiveId}`);
  Logger.log(`País: ${paisCode}`);
  Logger.log(`Campos solicitados: ${campos}`);

  try {
    const options = {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Content-Type': 'application/json'
      }
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log(`Error ${responseCode}: ${responseText}`);

      let errorData;
      try {
        errorData = JSON.parse(responseText);
      } catch (e) {
        return {
          error: `Error ${responseCode}: ${responseText}`,
          encontrado: false,
          codigoError: responseCode
        };
      }

      let mensajeError = errorData.error ? errorData.error.message : 'Error desconocido';

      // Log detallado del error
      Logger.log(`Error completo: ${JSON.stringify(errorData)}`);

      // Detectar errores específicos
      if (responseCode === 500 && mensajeError.includes('unknown error')) {
        mensajeError = `⚠️ Error 500 de la API de Biblioteca de Anuncios\n\n` +
                      `Este error suele ocurrir porque:\n\n` +
                      `1. El ID "${adArchiveId}" no es válido para la Biblioteca de Anuncios\n` +
                      `   Los IDs de la Biblioteca son diferentes a los AD IDs de Ads Manager\n\n` +
                      `2. El anuncio no está disponible públicamente\n\n` +
                      `3. El anuncio es muy antiguo o fue eliminado\n\n` +
                      `SOLUCIÓN RECOMENDADA:\n` +
                      `• Ve directamente a: https://www.facebook.com/ads/library/\n` +
                      `• Busca el anuncio manualmente\n` +
                      `• Copia el ID correcto desde la URL\n\n` +
                      `Ejemplo de URL: facebook.com/ads/library/?id=XXXXXXXXX\n` +
                      `El ID correcto sería: XXXXXXXXX\n\n` +
                      `Error técnico: ${errorData.error.message}`;
      } else if (mensajeError.includes('OAuth') || responseCode === 400) {
        mensajeError = `❌ Error de Permisos (${responseCode})\n\n` +
                      `Aunque tienes permisos "ads_read", la API de Biblioteca\n` +
                      `puede requerir configuración adicional.\n\n` +
                      `Error: ${errorData.error.message}`;
      }

      return {
        error: mensajeError,
        encontrado: false,
        codigoError: responseCode
      };
    }

    const json = JSON.parse(responseText);
    const datos = json.data || [];

    Logger.log(`✅ Respuesta exitosa!`);
    Logger.log(`Resultados encontrados: ${datos.length}`);

    if (datos.length > 0) {
      Logger.log(`Datos del primer resultado:`);
      Logger.log(JSON.stringify(datos[0], null, 2));
    } else {
      Logger.log(`⚠️ No se encontraron resultados para ID: ${adArchiveId}`);
    }

    // Buscar el anuncio específico
    let anuncioEncontrado = null;

    // Primero buscar por ID exacto
    for (const anuncio of datos) {
      if (anuncio.id === adArchiveId || anuncio.id === String(adArchiveId)) {
        anuncioEncontrado = anuncio;
        Logger.log(`✓ Encontrado por ID exacto: ${anuncio.id}`);
        break;
      }
    }

    // Si no, tomar el primero (la búsqueda puede haberlo encontrado)
    if (!anuncioEncontrado && datos.length > 0) {
      anuncioEncontrado = datos[0];
      Logger.log(`✓ Usando primer resultado (ID: ${anuncioEncontrado.id})`);
    }

    if (!anuncioEncontrado) {
      Logger.log(`❌ Anuncio no encontrado en resultados`);
      return {
        encontrado: false,
        error: null
      };
    }

    // Procesar el resultado con datos PÚBLICOS disponibles
    return {
      encontrado: true,
      id: anuncioEncontrado.id,
      pageName: anuncioEncontrado.page_name || 'N/A',
      pageId: anuncioEncontrado.page_id || 'N/A',
      creationTime: anuncioEncontrado.ad_creation_time || 'N/A',
      startTime: anuncioEncontrado.ad_delivery_start_time || 'N/A',
      stopTime: anuncioEncontrado.ad_delivery_stop_time || 'En circulación',
      snapshotUrl: anuncioEncontrado.ad_snapshot_url || 'N/A',

      // Creativos y textos (pueden no estar disponibles con campos mínimos)
      creativeBodies: anuncioEncontrado.ad_creative_bodies || ['Ver en snapshot'],
      linkTitles: anuncioEncontrado.ad_creative_link_titles || ['Ver en snapshot'],
      linkDescriptions: anuncioEncontrado.ad_creative_link_descriptions || ['Ver en snapshot'],
      linkCaptions: anuncioEncontrado.ad_creative_link_captions || [],

      // Plataformas y alcance (pueden no estar disponibles)
      platforms: anuncioEncontrado.publisher_platforms || ['N/A'],
      audienceSize: anuncioEncontrado.estimated_audience_size || 'N/A',

      datosCompletos: anuncioEncontrado
    };

  } catch (e) {
    Logger.log(`Error en consulta: ${e.message}`);
    return {
      error: e.message,
      encontrado: false
    };
  }
}

/**
 * Muestra el resultado de la consulta en un alert formateado
 * @param {Object} resultado - Datos del anuncio
 * @param {Object} ui - Objeto UI de Spreadsheet
 */
function mostrarResultadoBiblioteca(resultado, ui) {
  let mensaje = '📊 ANUNCIO DE LA COMPETENCIA\n';
  mensaje += '━'.repeat(50) + '\n\n';

  mensaje += `🆔 ID: ${resultado.id}\n\n`;

  // Información de la página
  mensaje += `📄 PÁGINA:\n`;
  mensaje += `   ${resultado.pageName}\n`;
  mensaje += `   ID: ${resultado.pageId}\n\n`;

  // Fechas
  if (resultado.startTime !== 'N/A') {
    const fechaInicio = new Date(resultado.startTime);
    mensaje += `📅 Inicio: ${fechaInicio.toLocaleDateString()}\n`;
  }

  if (resultado.stopTime !== 'En circulación') {
    const fechaFin = new Date(resultado.stopTime);
    mensaje += `📅 Fin: ${fechaFin.toLocaleDateString()}\n`;

    // Calcular días activo
    if (resultado.startTime !== 'N/A') {
      const dias = Math.floor((new Date(resultado.stopTime) - new Date(resultado.startTime)) / (1000 * 60 * 60 * 24));
      mensaje += `   (${dias} días activo)\n`;
    }
  } else {
    mensaje += `📅 Estado: ${resultado.stopTime}\n`;
  }
  mensaje += '\n';

  // Plataformas
  if (resultado.platforms.length > 0) {
    mensaje += `📱 Plataformas: ${resultado.platforms.join(', ')}\n\n`;
  }

  // Audiencia estimada
  if (resultado.audienceSize !== 'N/A') {
    mensaje += `👥 Audiencia estimada: ${resultado.audienceSize}\n\n`;
  }

  // Creativos
  mensaje += `📝 CREATIVOS:\n`;
  if (resultado.creativeBodies.length > 0) {
    const texto = resultado.creativeBodies[0].substring(0, 100);
    mensaje += `   Texto: ${texto}${resultado.creativeBodies[0].length > 100 ? '...' : ''}\n`;
  }

  if (resultado.linkTitles.length > 0) {
    mensaje += `   Título: ${resultado.linkTitles[0]}\n`;
  }

  if (resultado.linkDescriptions.length > 0) {
    mensaje += `   Descripción: ${resultado.linkDescriptions[0].substring(0, 80)}...\n`;
  }

  mensaje += '\n';
  mensaje += `🔗 Ver anuncio completo:\n`;
  mensaje += `   ${resultado.snapshotUrl}\n`;

  mensaje += '\n━'.repeat(50) + '\n';
  mensaje += '\n💡 DATOS GUARDADOS en "📚 Biblioteca de Anuncios"\n';
  mensaje += '\nℹ️ Datos de gasto/impresiones NO disponibles\n';
  mensaje += 'para anuncios de la competencia.\n';
  mensaje += '\nVer log completo: Extensiones > Apps Script > Ver registros';

  // Log completo
  Logger.log('=== ANUNCIO DE LA COMPETENCIA ===');
  Logger.log(JSON.stringify(resultado.datosCompletos, null, 2));

  ui.alert('✅ Anuncio Encontrado', mensaje, ui.ButtonSet.OK);

  // Escribir en hoja
  escribirResultadoEnHoja(resultado);
}

/**
 * Escribe el resultado en una hoja de cálculo
 * @param {Object} resultado - Datos del anuncio
 */
function escribirResultadoEnHoja(resultado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('📚 Biblioteca de Anuncios');

  // Crear hoja si no existe
  if (!hoja) {
    hoja = ss.insertSheet('📚 Biblioteca de Anuncios');
    hoja.setTabColor('#ff6d00');

    // Headers actualizados con campos públicos
    const headers = [
      'FECHA CONSULTA', 'ID ANUNCIO', 'PÁGINA', 'PAGE ID',
      'INICIO', 'FIN', 'DÍAS ACTIVO', 'ESTADO',
      'TEXTO PRINCIPAL', 'TÍTULO ENLACE', 'DESCRIPCIÓN',
      'PLATAFORMAS', 'AUDIENCIA EST.', 'SNAPSHOT URL'
    ];

    hoja.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground('#ff6d00')
         .setFontColor('white')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }

  // Calcular días activo
  let diasActivo = 'N/A';
  if (resultado.startTime !== 'N/A' && resultado.stopTime !== 'En circulación') {
    diasActivo = Math.floor((new Date(resultado.stopTime) - new Date(resultado.startTime)) / (1000 * 60 * 60 * 24));
  }

  // Estado
  const estado = resultado.stopTime === 'En circulación' ? 'ACTIVO' : 'FINALIZADO';

  // Agregar nueva fila con datos públicos
  const fechaConsulta = new Date();
  const nuevaFila = [
    fechaConsulta,
    resultado.id,
    resultado.pageName,
    resultado.pageId,
    resultado.startTime !== 'N/A' ? new Date(resultado.startTime) : 'N/A',
    resultado.stopTime !== 'En circulación' ? new Date(resultado.stopTime) : 'En circulación',
    diasActivo,
    estado,
    resultado.creativeBodies.length > 0 ? resultado.creativeBodies[0] : 'N/A',
    resultado.linkTitles.length > 0 ? resultado.linkTitles[0] : 'N/A',
    resultado.linkDescriptions.length > 0 ? resultado.linkDescriptions[0] : 'N/A',
    resultado.platforms.length > 0 ? resultado.platforms.join(', ') : 'N/A',
    resultado.audienceSize,
    resultado.snapshotUrl
  ];

  const ultimaFila = hoja.getLastRow();
  hoja.getRange(ultimaFila + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  // Ajustar columnas
  hoja.autoResizeColumns(1, nuevaFila.length);

  // Ancho específico para columnas de texto
  hoja.setColumnWidth(9, 300);  // Texto principal
  hoja.setColumnWidth(10, 200); // Título
  hoja.setColumnWidth(11, 250); // Descripción

  ss.toast('Resultado guardado en "📚 Biblioteca de Anuncios"', '✅ Completado', 3);
}

/**
 * FUNCIÓN SIMPLIFICADA DE PRUEBA
 * Solo busca el primer anuncio y muestra datos básicos
 */
function buscarPrimerAnuncio() {
  const ui = SpreadsheetApp.getUi();

  const respTermino = ui.prompt(
    '🔍 Búsqueda Simple - Prueba',
    'Ingresa un término de búsqueda:',
    ui.ButtonSet.OK_CANCEL
  );

  if (respTermino.getSelectedButton() !== ui.Button.OK) return;
  const termino = respTermino.getResponseText().trim();

  if (!termino) {
    ui.alert('Error', 'Debes ingresar un término', ui.ButtonSet.OK);
    return;
  }

  const respPais = ui.prompt(
    '🌍 País',
    'Código de país (2 letras):',
    ui.ButtonSet.OK_CANCEL
  );

  if (respPais.getSelectedButton() !== ui.Button.OK) return;
  const paisCode = respPais.getResponseText().trim().toUpperCase();

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_TOKEN');

  if (!token) {
    ui.alert('Error', 'No hay token', ui.ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Buscando...', '⏳', -1);

  try {
    // SOLO CAMPOS ABSOLUTAMENTE BÁSICOS
    const campos = 'id,page_name,ad_snapshot_url';

    const url = `https://graph.facebook.com/v22.0/ads_archive?` +
                `access_token=${token}&` +
                `ad_reached_countries=${encodeURIComponent(JSON.stringify([paisCode]))}&` +
                `search_terms=${encodeURIComponent(termino)}&` +
                `fields=${campos}&` +
                `limit=1`;

    Logger.log(`URL: ${url.replace(token, 'TOKEN')}`);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);

    Logger.log(`Response Code: ${responseCode}`);
    Logger.log(`Response: ${responseText}`);

    if (responseCode !== 200) {
      ui.alert('Error', `Error ${responseCode}\n\n${responseText}`, ui.ButtonSet.OK);
      return;
    }

    const json = JSON.parse(responseText);
    const datos = json.data || [];

    if (datos.length === 0) {
      ui.alert('Sin Resultados', `No se encontraron anuncios con "${termino}"`, ui.ButtonSet.OK);
      return;
    }

    const anuncio = datos[0];

    let mensaje = '✅ PRIMER ANUNCIO ENCONTRADO\n\n';
    mensaje += `📄 Página: ${anuncio.page_name || 'N/A'}\n`;
    mensaje += `🆔 ID: ${anuncio.id || 'N/A'}\n`;
    mensaje += `🔗 Snapshot: ${anuncio.ad_snapshot_url ? 'Disponible' : 'No disponible'}\n\n`;
    mensaje += `Ver datos completos en el log`;

    Logger.log('=== ANUNCIO COMPLETO ===');
    Logger.log(JSON.stringify(anuncio, null, 2));

    ui.alert('Anuncio Encontrado', mensaje, ui.ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
    Logger.log(`Error: ${e.message}\n${e.stack}`);
  }
}

/**
 * Función mejorada: Buscar anuncios de competencia por término
 * y extraer todos los datos públicos posibles
 */
function buscarAnunciosPorTermino() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const respTermino = ui.prompt(
    '🔍 Buscar Anuncios de Competencia',
    'Ingresa un término de búsqueda:\n\n' +
    'Ejemplos:\n' +
    '• Nombre de marca: "Nueva Acrópolis"\n' +
    '• Palabra clave: "cursos online"\n' +
    '• Nombre de producto: "iPhone 15"',
    ui.ButtonSet.OK_CANCEL
  );

  if (respTermino.getSelectedButton() !== ui.Button.OK) return;
  const termino = respTermino.getResponseText().trim();

  if (!termino) {
    ui.alert('Error', 'Debes ingresar un término de búsqueda', ui.ButtonSet.OK);
    return;
  }

  const respPais = ui.prompt(
    '🌍 Seleccionar País',
    'Ingresa el código del país (2 letras):\n\n' +
    'PE (Perú), MX (México), CO (Colombia),\n' +
    'AR (Argentina), US (Estados Unidos), etc.',
    ui.ButtonSet.OK_CANCEL
  );

  if (respPais.getSelectedButton() !== ui.Button.OK) return;
  const paisCode = respPais.getResponseText().trim().toUpperCase();

  if (paisCode.length !== 2) {
    ui.alert('Error', 'El código debe ser de 2 letras', ui.ButtonSet.OK);
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_TOKEN');

  if (!token) {
    ui.alert('❌ Error', 'No hay token configurado', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Buscando anuncios...', '⏳ Buscando', -1);

  try {
    // EMPEZAR CON CAMPOS MÍNIMOS que sabemos que funcionan
    // Si funciona, intentaremos agregar más campos después
    const campos = 'id,ad_snapshot_url,page_name,ad_delivery_start_time,ad_delivery_stop_time';

    // NOTA: Los campos de creative (copy, títulos, etc.) pueden causar error 500
    // para anuncios comerciales. Los probaremos después si estos funcionan.

    const countries = JSON.stringify([paisCode]);

    const params = {
      access_token: token,
      ad_reached_countries: countries,
      search_terms: termino,
      ad_type: 'ALL',
      fields: campos,
      limit: '50'
    };

    const queryString = Object.keys(params)
      .map(key => `${key}=${encodeURIComponent(params[key])}`)
      .join('&');

    const url = `https://graph.facebook.com/v22.0/ads_archive?${queryString}`;

    Logger.log(`=== BÚSQUEDA POR TÉRMINO ===`);
    Logger.log(`Término: "${termino}"`);
    Logger.log(`País: ${paisCode}`);
    Logger.log(`Campos solicitados: ${campos}`);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      const errorText = response.getContentText();
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);

      try {
        const errorData = JSON.parse(errorText);
        let errorMsg = errorData.error ? errorData.error.message : 'Error desconocido';

        if (errorMsg.includes('OAuth') || responseCode === 400) {
          errorMsg = `Error de permisos.\n\nGenera un token con permisos "ads_read" desde:\nhttps://developers.facebook.com/tools/explorer/\n\nError: ${errorMsg}`;
        }

        ui.alert('Error', errorMsg, ui.ButtonSet.OK);
      } catch (e) {
        ui.alert('Error', `Error ${responseCode}: ${errorText}`, ui.ButtonSet.OK);
      }
      return;
    }

    const json = JSON.parse(response.getContentText());
    const datos = json.data || [];

    ss.toast('', '', 1);

    Logger.log(`✅ Resultados obtenidos: ${datos.length}`);

    if (datos.length === 0) {
      ui.alert('Sin Resultados', `No se encontraron anuncios con el término "${termino}" en ${paisCode}`, ui.ButtonSet.OK);
      return;
    }

    // Log completo de datos
    Logger.log('=== DATOS COMPLETOS ===');
    Logger.log(JSON.stringify(datos, null, 2));

    // Guardar resultados en hoja
    const datosGuardados = guardarResultadosEnHoja(datos, termino, paisCode);

    // Preparar mensaje con resumen
    let mensaje = `✅ Encontrados ${datos.length} anuncios\n\n`;
    mensaje += `📊 DATOS EXTRAÍDOS:\n`;
    mensaje += `• Anuncios activos: ${datosGuardados.activos}\n`;
    mensaje += `• Anuncios finalizados: ${datos.length - datosGuardados.activos}\n\n`;

    mensaje += `Primeros 5 resultados:\n\n`;

    datos.slice(0, 5).forEach((ad, i) => {
      const estado = ad.ad_delivery_stop_time ? '⚪ Finalizado' : '🟢 Activo';
      const fechaInicio = ad.ad_delivery_start_time ? new Date(ad.ad_delivery_start_time).toLocaleDateString() : 'N/A';

      mensaje += `${i + 1}. ${ad.page_name || 'Sin nombre'} ${estado}\n`;
      mensaje += `   ID: ${ad.id}\n`;
      mensaje += `   Inicio: ${fechaInicio}\n\n`;
    });

    mensaje += `\n💾 Datos guardados en "📚 Biblioteca de Anuncios"\n`;
    mensaje += `\n📋 Ver todos los datos en el log:\n`;
    mensaje += `Extensiones > Apps Script > Ver registros`;

    ui.alert('✅ Búsqueda Completada', mensaje, ui.ButtonSet.OK);

  } catch (e) {
    ss.toast('', '', 1);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
    Logger.log(`❌ Error: ${e.message}`);
  }
}

/**
 * Guarda los resultados de búsqueda en la hoja
 */
function guardarResultadosEnHoja(datos, termino, paisCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('📚 Biblioteca de Anuncios');

  if (!hoja) {
    hoja = ss.insertSheet('📚 Biblioteca de Anuncios');
    hoja.setTabColor('#ff6d00');

    const headers = [
      'FECHA CONSULTA', 'TÉRMINO BÚSQUEDA', 'PAÍS', 'ID ANUNCIO',
      'PÁGINA', 'ESTADO', 'INICIO', 'FIN', 'DÍAS ACTIVO',
      'SNAPSHOT URL'
    ];

    hoja.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground('#ff6d00')
         .setFontColor('white')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }

  const fechaConsulta = new Date();
  let activos = 0;

  datos.forEach(ad => {
    // Calcular días activo
    let diasActivo = 'N/A';
    if (ad.ad_delivery_start_time && ad.ad_delivery_stop_time) {
      diasActivo = Math.floor((new Date(ad.ad_delivery_stop_time) - new Date(ad.ad_delivery_start_time)) / (1000 * 60 * 60 * 24));
    }

    const estado = ad.ad_delivery_stop_time ? 'FINALIZADO' : 'ACTIVO';
    if (estado === 'ACTIVO') activos++;

    const fila = [
      fechaConsulta,
      termino,
      paisCode,
      ad.id || 'N/A',
      ad.page_name || 'N/A',
      estado,
      ad.ad_delivery_start_time ? new Date(ad.ad_delivery_start_time) : 'N/A',
      ad.ad_delivery_stop_time ? new Date(ad.ad_delivery_stop_time) : 'En circulación',
      diasActivo,
      ad.ad_snapshot_url || 'N/A'
    ];

    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila + 1, 1, 1, fila.length).setValues([fila]);
  });

  // Formatear
  hoja.autoResizeColumns(1, 10);

  return {
    activos: activos
  };
}
