/**
 * ==================== CREADOR DE CAMPAÑAS DE WHATSAPP ====================
 * Sistema independiente para crear campañas de mensajes de WhatsApp en Meta Ads
 */

/**
 * Crea la hoja de configuración de campañas
 */
function crearHojaCampanas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('📲 Crear Campaña WhatsApp');

  if (!hoja) {
    hoja = ss.insertSheet('📲 Crear Campaña WhatsApp');
    hoja.setTabColor('#25d366'); // Verde WhatsApp
  }

  hoja.clear();

  // Estructura de la hoja
  const datos = [
    // TÍTULO
    ['🚀 CREADOR DE CAMPAÑAS - WHATSAPP', '', '', ''],
    ['Completa este formulario paso a paso. Los campos ROJOS son obligatorios, los AMARILLOS son opcionales.', '', '', ''],
    ['', '', '', ''],

    // SECCIÓN 1: INFORMACIÓN DE LA CAMPAÑA
    ['📋 PASO 1: INFORMACIÓN DE LA CAMPAÑA', '', '', ''],
    ['*Nombre de la Campaña', '👉 ESCRIBE AQUÍ', '¿Qué escribo aquí? Un nombre identificador para tu campaña. Ejemplo: "WhatsApp Enero 2025" o "Promo Verano 2025"', 'OBLIGATORIO'],
    ['Presupuesto Diario', '10', '¿Qué escribo aquí? Cuántos dólares quieres gastar POR DÍA. Ejemplo: 10 (sin símbolos). Mínimo: 1 USD', 'Opcional'],
    ['Estado Inicial de la Campaña', 'PAUSED', '¿Qué escribo aquí? Escribe PAUSED (campaña guardada como borrador, NO se publica) o ACTIVE (se publica inmediatamente). RECOMENDADO: PAUSED', 'Opcional'],
    ['', '', '', ''],

    // SECCIÓN 2: CONFIGURACIÓN DE WHATSAPP
    ['📱 PASO 2: CONFIGURACIÓN DE WHATSAPP', '', '', ''],
    ['Número de WhatsApp de tu negocio', '+51961048352', '¿Qué escribo aquí? El número de teléfono de WhatsApp donde recibirás mensajes. IMPORTANTE: Debe incluir + y código de país. Ejemplo: +51987654321 (Perú) o +57300123456 (Colombia)', 'Opcional'],
    ['*Page ID de Facebook', '👉 PEGA AQUÍ EL PAGE ID', '¿Qué es esto? Es un número largo que identifica tu página de Facebook. ¿Cómo lo obtengo? Ve al menú arriba: "Creador de Campañas WhatsApp" → "Obtener Page ID de Facebook". Ejemplo: 123456789012345', 'OBLIGATORIO'],
    ['', '', '', ''],

    // SECCIÓN 3: AUDIENCIA
    ['🎯 PASO 3: DEFINE TU AUDIENCIA (¿A quién le mostrarás el anuncio?)', '', '', ''],
    ['*País(es) donde mostrar el anuncio', '👉 ESCRIBE LOS CÓDIGOS DE PAÍS', '¿Qué escribo aquí? Códigos de país de 2 LETRAS separados por COMAS. Ejemplos comunes: PE (Perú), CO (Colombia), MX (México), AR (Argentina), CL (Chile), EC (Ecuador), BO (Bolivia), US (Estados Unidos), ES (España). Si quieres varios países escribe: PE,CO,MX', 'OBLIGATORIO'],
    ['Edad Mínima de tu audiencia', '18', '¿Qué escribo aquí? La edad mínima de personas que verán tu anuncio. Ejemplo: 18 (solo números, entre 18 y 65)', 'Opcional'],
    ['Edad Máxima de tu audiencia', '65', '¿Qué escribo aquí? La edad máxima de personas que verán tu anuncio. Ejemplo: 65 (solo números, entre 18 y 65)', 'Opcional'],
    ['Género de tu audiencia', 'ALL', '¿Qué escribo aquí? A qué género mostrar el anuncio. Opciones: ALL (todos), MALE (solo hombres), FEMALE (solo mujeres)', 'Opcional'],
    ['Intereses de tu audiencia', '', '¿Qué escribo aquí? OPCIONAL - Temas de interés de tu audiencia, separados por comas. Ejemplo: Marketing Digital, Emprendimiento, Negocios. DEJA VACÍO si no quieres filtrar por intereses', 'Opcional'],
    ['', '', '', ''],

    // SECCIÓN 4: CONTENIDO DEL ANUNCIO
    ['✍️ PASO 4: CREA EL CONTENIDO DE TU ANUNCIO', '', '', ''],
    ['*Nombre del Anuncio (solo para ti)', '👉 ESCRIBE AQUÍ', '¿Qué escribo aquí? Un nombre para identificar este anuncio (los clientes NO lo ven). Ejemplo: "Anuncio Promoción Enero" o "Ad Verano 2025"', 'OBLIGATORIO'],
    ['*Texto Principal del Anuncio', '👉 ESCRIBE AQUÍ EL MENSAJE', '¿Qué escribo aquí? El mensaje que las personas VERÁN en Facebook/Instagram. Ejemplo: "¡Descuento del 50% en todos nuestros productos! Contáctanos por WhatsApp" o "Aprende marketing digital desde cero. Escríbenos para más info"', 'OBLIGATORIO'],
    ['Botón de Llamada a la Acción', 'WHATSAPP_MESSAGE', '¿Qué escribo aquí? El texto del botón. Opciones: WHATSAPP_MESSAGE (muestra "Enviar mensaje"), LEARN_MORE ("Más información"), CONTACT_US ("Contáctanos"). RECOMENDADO: WHATSAPP_MESSAGE', 'Opcional'],
    ['Mensaje Automático de WhatsApp', 'Hola, quiero más información', '¿Qué escribo aquí? El mensaje que se escribirá AUTOMÁTICAMENTE cuando alguien haga clic en el botón de WhatsApp. Ejemplo: "Hola, vi tu anuncio y quiero más información" o "Hola, me interesa el descuento"', 'Opcional'],
    ['URL de Imagen o Video (opcional)', '', '¿Qué escribo aquí? OPCIONAL - Link público de una imagen o video. Si usas Google Drive, asegúrate que el enlace sea público. Ejemplo: https://drive.google.com/... o link directo a imagen. DEJA VACÍO si no quieres imagen/video', 'Opcional'],
    ['', '', '', ''],

    // SECCIÓN 5: CONFIGURACIÓN AVANZADA
    ['⚙️ CONFIGURACIÓN AVANZADA (YA ESTÁ CONFIGURADO - NO MODIFICAR)', '', '', ''],
    ['Objetivo de Optimización', 'CONVERSATIONS', 'NO MODIFICAR - Ya está configurado para optimizar conversaciones de WhatsApp', ''],
    ['Evento de Facturación', 'CONVERSATIONS', 'NO MODIFICAR - Ya está configurado para cobrar por conversaciones', ''],
    ['', '', '', ''],

    // INSTRUCCIONES
    ['📌 INSTRUCCIONES - LEE ANTES DE CREAR LA CAMPAÑA', '', '', ''],
    ['✅ PASO A PASO:', '', '', ''],
    ['1️⃣ Completa TODOS los campos ROJOS (marcados con *) - Son OBLIGATORIOS', '', '', ''],
    ['2️⃣ Completa los campos AMARILLOS que necesites - Son OPCIONALES', '', '', ''],
    ['3️⃣ Lee la columna C "¿Qué escribo aquí?" para entender cada campo', '', '', ''],
    ['4️⃣ Si necesitas el Page ID: Ve al menú → Creador de Campañas WhatsApp → Obtener Page ID de Facebook', '', '', ''],
    ['5️⃣ Códigos de país MÁS COMUNES: PE (Perú), CO (Colombia), MX (México), AR (Argentina), CL (Chile), EC (Ecuador), US (Estados Unidos)', '', '', ''],
    ['6️⃣ Una vez completado TODO, haz clic en el botón verde "▶️ CREAR CAMPAÑA" abajo', '', '', ''],
    ['', '', '', ''],
    ['⚠️ IMPORTANTE: Si te falta algún campo obligatorio, el sistema te avisará con un mensaje de error indicando qué falta.', '', '', ''],
    ['', '', '', ''],

    // BOTÓN DE ACCIÓN
    ['▶️ CREAR CAMPAÑA', 'Estado: Listo para crear', 'Haz clic aquí cuando hayas completado TODOS los campos obligatorios', '']
  ];

  // Escribir datos
  hoja.getRange(1, 1, datos.length, 4).setValues(datos);

  // FORMATEO
  // Título principal
  hoja.getRange(1, 1, 1, 4)
      .merge()
      .setBackground('#25d366')
      .setFontColor('white')
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 50);

  // Subtítulo explicativo
  hoja.getRange(2, 1, 1, 4)
      .merge()
      .setBackground('#e8f5e9')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 35);

  // Encabezados de sección (PASOS) - solo combinar columnas A-C, dejar D libre
  const seccionesRangos = [4, 9, 13, 20, 27, 31];
  seccionesRangos.forEach(fila => {
    hoja.getRange(fila, 1, 1, 3)
        .merge()
        .setBackground('#4a90e2')
        .setFontColor('white')
        .setFontWeight('bold')
        .setFontSize(12)
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
    // Dar formato también a la columna D en la misma fila
    hoja.getRange(fila, 4)
        .setBackground('#4a90e2')
        .setFontColor('#4a90e2'); // Mismo color para que sea invisible
    hoja.setRowHeight(fila, 40); // Altura para encabezados
  });

  // Instrucciones importantes - destacar
  hoja.getRange(31, 1, 1, 4).setBackground('#fff3cd').setFontWeight('bold').setFontSize(12); // Título instrucciones
  hoja.getRange(32, 1, 1, 4).setBackground('#e8f5e9').setFontWeight('bold'); // "PASO A PASO"
  hoja.getRange(33, 1, 9, 4).setBackground('#f9f9f9'); // Instrucciones numeradas

  // Botón de acción
  hoja.getRange(42, 1, 1, 2)
      .merge()
      .setBackground('#128c7e')
      .setFontColor('white')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
  hoja.setRowHeight(42, 45);

  // Columnas - hacer columna C más ancha para las explicaciones
  hoja.setColumnWidth(1, 300); // Campo nombre
  hoja.setColumnWidth(2, 280); // Campo valor
  hoja.setColumnWidth(3, 600); // Explicación "¿Qué escribo aquí?" - MUY ANCHA
  hoja.setColumnWidth(4, 120); // Indicador OBLIGATORIO/Opcional

  // Campos OBLIGATORIOS (fondo rojo claro + borde grueso)
  const camposObligatorios = [5, 11, 14, 21, 22];
  camposObligatorios.forEach(fila => {
    hoja.getRange(fila, 2)
        .setBackground('#ffcccc') // Rojo claro para campos obligatorios
        .setFontWeight('bold')
        .setFontColor('#cc0000')
        .setBorder(true, true, true, true, false, false, 'red', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // Campos opcionales (fondo amarillo suave)
  const camposOpcionales = [6, 7, 10, 15, 16, 17, 18, 23, 24, 25, 28, 29];
  camposOpcionales.forEach(fila => {
    hoja.getRange(fila, 2)
        .setBackground('#fff9e6') // Amarillo muy suave
        .setBorder(true, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  });

  // Hacer los labels de columna A más visibles
  hoja.getRange('A5:A29').setFontWeight('bold').setFontSize(10);

  // Columna D - indicadores OBLIGATORIO/Opcional
  hoja.getRange('D5:D29').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(9);

  // Agregar bordes y fondos a todas las celdas de campos para mayor claridad
  const todasLasFilasDeCampos = [5, 6, 7, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29];
  todasLasFilasDeCampos.forEach(fila => {
    // Borde para toda la fila de A a D
    hoja.getRange(fila, 1, 1, 4).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

    // Fondo gris muy claro para columna A (etiquetas)
    hoja.getRange(fila, 1).setBackground('#f5f5f5');

    // Fondo azul muy claro para columna C (explicaciones/ayuda)
    hoja.getRange(fila, 3).setBackground('#f0f8ff');
  });

  // Colorear indicadores
  for (let i = 5; i <= 29; i++) {
    const valor = hoja.getRange(i, 4).getValue();
    if (valor === 'OBLIGATORIO') {
      hoja.getRange(i, 4).setBackground('#ffcccc').setFontColor('#cc0000');
    } else if (valor === 'Opcional') {
      hoja.getRange(i, 4).setBackground('#e8f5e9').setFontColor('#666666');
    }
  }

  // Wrap text en columna C (explicaciones) y aumentar altura de filas
  hoja.getRange('C:C').setWrap(true).setVerticalAlignment('top');
  hoja.getRange('A:D').setVerticalAlignment('top');

  // Aumentar altura de las filas con campos para que se lean bien las explicaciones
  for (let i = 5; i <= 25; i++) {
    hoja.setRowHeight(i, 60); // Altura generosa para leer las explicaciones
  }

  SpreadsheetApp.getUi().alert(
    '✅ Hoja Creada Correctamente',
    '📋 CÓMO USAR ESTE FORMULARIO:\n\n' +
    '1️⃣ Lee la columna C "¿Qué escribo aquí?" - explica cada campo\n' +
    '2️⃣ Completa TODOS los campos ROJOS (obligatorios)\n' +
    '3️⃣ Completa los campos AMARILLOS que necesites (opcionales)\n' +
    '4️⃣ Los campos están organizados por PASOS numerados\n' +
    '5️⃣ Para obtener el Page ID: Menú → Creador de Campañas WhatsApp → Obtener Page ID\n' +
    '6️⃣ Cuando termines, haz clic en "▶️ CREAR CAMPAÑA"\n\n' +
    '📝 Ejemplo de países: PE,CO,MX (separados por comas)\n' +
    '📝 El número de WhatsApp debe incluir +: +51987654321\n\n' +
    '⚠️ Si tienes dudas sobre un campo, lee la columna C',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Lee la configuración de la hoja y crea la campaña en Meta
 */
function crearCampanaWhatsApp() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('📲 Crear Campaña WhatsApp');

  if (!hoja) {
    ui.alert('❌ Error', 'Primero debes crear la hoja de configuración desde el menú.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Leer configuración
    const config = obtenerConfiguracion();
    const token = config.META.ACCESS_TOKEN;

    if (!token) {
      ui.alert('❌ Error', 'No hay token de Meta configurado.', ui.ButtonSet.OK);
      return;
    }

    if (config.META.CUENTAS.length === 0) {
      ui.alert('❌ Error', 'No hay cuentas publicitarias configuradas.', ui.ButtonSet.OK);
      return;
    }

    const accountId = config.META.CUENTAS[0][0]; // Primera cuenta

    // Leer datos de la hoja (referencias CORREGIDAS)
    const campana = {
      nombre: hoja.getRange('B5').getValue(),
      presupuesto: parseFloat(hoja.getRange('B6').getValue()) || 10,
      estado: hoja.getRange('B7').getValue() || 'PAUSED',
      numeroWhatsApp: hoja.getRange('B10').getValue() || '+51961048352',
      pageId: hoja.getRange('B11').getValue(),
      paises: hoja.getRange('B14').getValue() || 'PE',
      edadMin: parseInt(hoja.getRange('B15').getValue()) || 18,
      edadMax: parseInt(hoja.getRange('B16').getValue()) || 65,
      genero: hoja.getRange('B17').getValue() || 'ALL',
      intereses: hoja.getRange('B18').getValue() || '',
      nombreAnuncio: hoja.getRange('B21').getValue(),
      textoPrincipal: hoja.getRange('B22').getValue(),
      callToAction: hoja.getRange('B23').getValue() || 'WHATSAPP_MESSAGE',
      mensajeBienvenida: hoja.getRange('B24').getValue(),
      urlMedia: hoja.getRange('B25').getValue(),
      optimizacion: hoja.getRange('B28').getValue() || 'CONVERSATIONS',
      facturacion: hoja.getRange('B29').getValue() || 'CONVERSATIONS'
    };

    // Validaciones básicas con mensajes mejorados
    if (!campana.nombre || String(campana.nombre).trim() === '' || campana.nombre === '👉 ESCRIBE AQUÍ') {
      ui.alert('❌ Campo Obligatorio Vacío',
        '⚠️ PASO 1: Nombre de la Campaña\n\n' +
        'Debes ingresar un nombre para tu campaña en la celda B5.\n\n' +
        '📝 Ejemplo: "WhatsApp Enero 2025" o "Promo Verano"\n\n' +
        'Lee la columna C para más ayuda.',
        ui.ButtonSet.OK);
      return;
    }

    if (!campana.nombreAnuncio || String(campana.nombreAnuncio).trim() === '' || campana.nombreAnuncio === '👉 ESCRIBE AQUÍ') {
      ui.alert('❌ Campo Obligatorio Vacío',
        '⚠️ PASO 4: Nombre del Anuncio\n\n' +
        'Debes ingresar un nombre para tu anuncio en la celda B21.\n\n' +
        '📝 Ejemplo: "Anuncio Promoción Enero"\n\n' +
        'Lee la columna C para más ayuda.',
        ui.ButtonSet.OK);
      return;
    }

    if (!campana.textoPrincipal || String(campana.textoPrincipal).trim() === '' || campana.textoPrincipal === '👉 ESCRIBE AQUÍ EL MENSAJE') {
      ui.alert('❌ Campo Obligatorio Vacío',
        '⚠️ PASO 4: Texto Principal del Anuncio\n\n' +
        'Debes ingresar el mensaje que verán las personas en la celda B22.\n\n' +
        '📝 Ejemplo: "¡Descuento del 50%! Contáctanos por WhatsApp"\n\n' +
        'Lee la columna C para más ayuda.',
        ui.ButtonSet.OK);
      return;
    }

    if (!campana.pageId || String(campana.pageId).trim() === '' || campana.pageId === '👉 PEGA AQUÍ EL PAGE ID') {
      ui.alert('❌ Campo Obligatorio Vacío',
        '⚠️ PASO 2: Page ID de Facebook\n\n' +
        'Debes ingresar el Page ID de tu página de Facebook en la celda B11.\n\n' +
        '🔍 ¿Cómo obtenerlo?\n' +
        'Ve al menú: Creador de Campañas WhatsApp → Obtener Page ID de Facebook\n\n' +
        '📝 Es un número largo, ejemplo: 123456789012345',
        ui.ButtonSet.OK);
      return;
    }

    if (!campana.paises || String(campana.paises).trim() === '' || campana.paises === '👉 ESCRIBE LOS CÓDIGOS DE PAÍS') {
      ui.alert('❌ Campo Obligatorio Vacío',
        '⚠️ PASO 3: País(es) donde mostrar el anuncio\n\n' +
        'Debes especificar al menos un país en la celda B14.\n\n' +
        '📝 Usa códigos de 2 LETRAS separados por COMAS\n\n' +
        'Ejemplos:\n' +
        '• PE (solo Perú)\n' +
        '• PE,CO,MX (Perú, Colombia y México)\n' +
        '• US (Estados Unidos)\n\n' +
        'Lee la columna C para ver más países.',
        ui.ButtonSet.OK);
      return;
    }

    // Actualizar estado en la fila del botón
    hoja.getRange('B42').setValue('⏳ Creando campaña...');
    SpreadsheetApp.flush();

    // PASO 1: Crear Campaña
    const campaignId = crearCampanaBase(accountId, token, campana);

    // PASO 2: Crear Ad Set (Conjunto de Anuncios)
    const adsetId = crearAdSet(accountId, campaignId, token, campana);

    // PASO 3: Crear Anuncio (requiere creative)
    const adCreativeId = crearAdCreative(accountId, token, campana);
    const adId = crearAd(accountId, adsetId, adCreativeId, token, campana);

    // Actualizar estado final
    hoja.getRange('B42').setValue('✅ Campaña creada exitosamente');

    ui.alert(
      '✅ Campaña Creada Exitosamente',
      `✅ Tu campaña de WhatsApp ha sido creada correctamente\n\n` +
      `📋 Nombre: ${campana.nombre}\n\n` +
      `🆔 IDs Generados:\n` +
      `• Campaign ID: ${campaignId}\n` +
      `• AdSet ID: ${adsetId}\n` +
      `• Ad ID: ${adId}\n\n` +
      `📊 Estado: ${campana.estado === 'PAUSED' ? 'PAUSADO (Borrador)' : 'ACTIVO'}\n\n` +
      `💡 Puedes ver y editar tu campaña en Meta Ads Manager:\n` +
      `https://business.facebook.com/adsmanager`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    hoja.getRange('B42').setValue('❌ Error: ' + error.message);
    ui.alert(
      '❌ Error al Crear Campaña',
      `⚠️ Ocurrió un error al crear la campaña:\n\n${error.message}\n\n` +
      `📝 Revisa que:\n` +
      `• Todos los campos obligatorios estén completos\n` +
      `• El Page ID sea correcto\n` +
      `• Los códigos de país sean válidos (PE, CO, MX, etc.)\n` +
      `• El número de WhatsApp incluya + y código de país\n\n` +
      `Si el problema persiste, revisa el log del script.`,
      ui.ButtonSet.OK
    );
    Logger.log('Error completo: ' + error.stack);
  }
}

/**
 * Crea la campaña base en Meta
 */
function crearCampanaBase(accountId, token, campana) {
  const endpoint = `https://graph.facebook.com/v22.0/${accountId}/campaigns`;

  const payload = {
    name: campana.nombre,
    objective: 'OUTCOME_ENGAGEMENT', // Objetivo correcto para WhatsApp/Mensajes
    status: campana.estado,
    special_ad_categories: JSON.stringify([]), // Debe ser string JSON
    access_token: token
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error('Error creando campaña: ' + result.error.message);
  }

  Logger.log('Campaña creada: ' + result.id);
  return result.id;
}

/**
 * Crea el conjunto de anuncios (Ad Set)
 */
function crearAdSet(accountId, campaignId, token, campana) {
  const endpoint = `https://graph.facebook.com/v22.0/${accountId}/adsets`;

  // Validar y procesar países
  const paisesStr = campana.paises ? campana.paises.toString().trim() : '';

  if (!paisesStr || paisesStr === '') {
    throw new Error(
      '⚠️ PASO 3: Campo "País(es) donde mostrar el anuncio" está vacío\n\n' +
      'Debes especificar al menos un código de país en la celda B14.\n\n' +
      '📝 Formato: Códigos de 2 LETRAS separados por COMAS\n\n' +
      'Ejemplos:\n' +
      '• PE (solo Perú)\n' +
      '• PE,CO,MX (varios países)\n' +
      '• US (Estados Unidos)\n\n' +
      'Códigos comunes: PE, CO, MX, AR, CL, EC, BO, US, ES'
    );
  }

  const paisesArray = paisesStr.split(',').map(p => p.trim().toUpperCase()).filter(p => p.length > 0);

  if (paisesArray.length === 0) {
    throw new Error(
      '⚠️ PASO 3: No se encontraron países válidos\n\n' +
      'El formato debe ser códigos de 2 letras separados por comas.\n\n' +
      '📝 Ejemplos correctos:\n' +
      '• PE (un país)\n' +
      '• PE,CO,MX (varios países)\n\n' +
      '❌ Ejemplos incorrectos:\n' +
      '• Peru (debe ser PE)\n' +
      '• pe,co (deben ser mayúsculas, pero el sistema los convierte)'
    );
  }

  // Validar códigos de país (deben ser 2 letras)
  const paisesInvalidos = paisesArray.filter(p => p.length !== 2);
  if (paisesInvalidos.length > 0) {
    throw new Error(
      `⚠️ PASO 3: Códigos de país inválidos encontrados\n\n` +
      `Códigos incorrectos: ${paisesInvalidos.join(', ')}\n\n` +
      `📝 Los códigos de país deben ser exactamente 2 LETRAS\n\n` +
      `Ejemplos correctos:\n` +
      `• PE = Perú\n` +
      `• CO = Colombia\n` +
      `• MX = México\n` +
      `• AR = Argentina\n` +
      `• CL = Chile\n` +
      `• EC = Ecuador\n` +
      `• US = Estados Unidos\n` +
      `• ES = España\n\n` +
      `Formato: PE,CO,MX (separados por comas, SIN espacios antes/después)`
    );
  }

  // Construir targeting
  const targeting = {
    geo_locations: {
      countries: paisesArray
    },
    age_min: campana.edadMin,
    age_max: campana.edadMax
  };

  if (campana.genero !== 'ALL') {
    targeting.genders = campana.genero === 'MALE' ? [1] : [2];
  }

  const payload = {
    name: campana.nombre + ' - AdSet',
    campaign_id: campaignId,
    daily_budget: Math.round(campana.presupuesto * 100), // En centavos
    billing_event: 'IMPRESSIONS',
    optimization_goal: 'CONVERSATIONS',
    bid_strategy: 'LOWEST_COST_WITHOUT_CAP',
    targeting: JSON.stringify(targeting),
    status: campana.estado,
    promoted_object: JSON.stringify({
      page_id: campana.pageId || ''
    }),
    access_token: token
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error('Error creando AdSet: ' + result.error.message);
  }

  Logger.log('AdSet creado: ' + result.id);
  return result.id;
}

/**
 * Crea el creative (diseño del anuncio)
 */
function crearAdCreative(accountId, token, campana) {
  const endpoint = `https://graph.facebook.com/v22.0/${accountId}/adcreatives`;

  // Limpiar número de WhatsApp (quitar + y espacios)
  const whatsappNumber = campana.numeroWhatsApp.replace(/[\+\s]/g, '');

  // Construir el link de WhatsApp
  const mensajeBienvenida = campana.mensajeBienvenida || 'Hola, quiero más información';
  const whatsappLink = `https://wa.me/${whatsappNumber}?text=${encodeURIComponent(mensajeBienvenida)}`;

  const linkData = {
    message: campana.textoPrincipal,
    link: whatsappLink,
    call_to_action: {
      type: campana.callToAction,
      value: {
        link: whatsappLink
      }
    }
  };

  // Agregar imagen/video si existe
  if (campana.urlMedia) {
    if (campana.urlMedia.includes('.mp4') || campana.urlMedia.includes('video')) {
      linkData.video_id = campana.urlMedia;
    } else {
      linkData.picture = campana.urlMedia;
    }
  }

  const objectStorySpec = {
    page_id: campana.pageId,
    link_data: linkData
  };

  const payload = {
    name: campana.nombreAnuncio,
    object_story_spec: JSON.stringify(objectStorySpec),
    degrees_of_freedom_spec: JSON.stringify({
      creative_features_spec: {
        standard_enhancements: {
          enroll_status: 'OPT_OUT'
        }
      }
    }),
    access_token: token
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error('Error creando Creative: ' + result.error.message);
  }

  Logger.log('Creative creado: ' + result.id);
  return result.id;
}

/**
 * Crea el anuncio final
 */
function crearAd(accountId, adsetId, creativeId, token, campana) {
  const endpoint = `https://graph.facebook.com/v22.0/${accountId}/ads`;

  const payload = {
    name: campana.nombreAnuncio,
    adset_id: adsetId,
    creative: JSON.stringify({ creative_id: creativeId }),
    status: campana.estado,
    access_token: token
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error('Error creando Ad: ' + result.error.message);
  }

  Logger.log('Ad creado: ' + result.id);
  return result.id;
}

/**
 * Muestra todas las páginas de Facebook disponibles para el usuario
 * Utiliza el token de Meta configurado en el sistema
 */
function mostrarPageIDsDisponibles() {
  const ui = SpreadsheetApp.getUi();

  try {
    const config = obtenerConfiguracion();
    const token = config.META.ACCESS_TOKEN;

    if (!token) {
      ui.alert('❌ Error', 'No hay token de Meta configurado.', ui.ButtonSet.OK);
      return;
    }

    // Primero intentar obtener páginas del usuario
    const endpoint = `https://graph.facebook.com/v22.0/me/accounts?access_token=${token}`;
    const response = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      mostrarAlternativaPageID(result.error.message);
      return;
    }

    if (!result.data || result.data.length === 0) {
      mostrarAlternativaPageID('No se encontraron páginas vinculadas al token.');
      return;
    }

    let mensaje = '📄 PÁGINAS DE FACEBOOK DISPONIBLES:\n\n';
    result.data.forEach((page, index) => {
      mensaje += `${index + 1}. ${page.name}\n`;
      mensaje += `   📋 Page ID: ${page.id}\n`;
      mensaje += `   ${page.access_token ? '✅ Acceso correcto' : '⚠️ Sin acceso'}\n\n`;
    });

    mensaje += '──────────────────────────────\n\n';
    mensaje += '📝 ¿CÓMO USAR ESTE PAGE ID?\n\n';
    mensaje += '1️⃣ Copia el Page ID de la página que quieres usar\n';
    mensaje += '2️⃣ Ve a la hoja "📲 Crear Campaña WhatsApp"\n';
    mensaje += '3️⃣ Pega el Page ID en la celda B11 (PASO 2)\n\n';
    mensaje += '⚠️ Asegúrate de copiar SOLO los números del Page ID';

    ui.alert('📄 Páginas Disponibles', mensaje, ui.ButtonSet.OK);

    // También escribir en la hoja para referencia
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName('📲 Crear Campaña WhatsApp');

    if (hoja) {
      let fila = 44; // Después del botón de crear campaña (fila 42) + 1 espacio
      hoja.getRange(fila, 1, 1, 4)
        .merge()
        .setValue('📄 TUS PÁGINAS DE FACEBOOK DISPONIBLES')
        .setBackground('#e3f2fd')
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center');
      fila++;

      hoja.getRange(fila, 1).setValue('Nombre de la Página').setFontWeight('bold').setBackground('#f0f0f0');
      hoja.getRange(fila, 2).setValue('Page ID (Copia este número)').setFontWeight('bold').setBackground('#f0f0f0');
      hoja.getRange(fila, 3).setValue('Estado de Acceso').setFontWeight('bold').setBackground('#f0f0f0');
      fila++;

      result.data.forEach(page => {
        hoja.getRange(fila, 1).setValue(page.name);
        hoja.getRange(fila, 2).setValue(page.id).setFontWeight('bold').setFontColor('#0066cc');
        hoja.getRange(fila, 3).setValue(page.access_token ? '✅ Acceso correcto' : '⚠️ Sin acceso');
        fila++;
      });

      hoja.getRange(fila, 1, 1, 3)
        .merge()
        .setValue('👆 Copia el Page ID de la página que quieres usar y pégalo en la celda B11 arriba')
        .setBackground('#fff9e6')
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
    }

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
    Logger.log('Error obteniendo pages: ' + error.stack);
  }
}

/**
 * Muestra instrucciones alternativas para obtener el Page ID
 */
function mostrarAlternativaPageID(errorMsg) {
  const ui = SpreadsheetApp.getUi();

  const mensaje = `⚠️ ${errorMsg}\n\n` +
    '═══════════════════════════════════\n\n' +
    '📌 CÓMO OBTENER TU PAGE ID MANUALMENTE:\n\n' +
    '🔹 MÉTODO 1 - Desde tu Página de Facebook:\n' +
    '1️⃣ Ve a tu página de Facebook en el navegador\n' +
    '2️⃣ Haz clic en "Acerca de" en el menú lateral izquierdo\n' +
    '3️⃣ Desplázate hasta el final de la sección "Acerca de"\n' +
    '4️⃣ Busca "ID de página" - es un número de 15-16 dígitos\n' +
    '5️⃣ Copia ese número\n\n' +
    '🔹 MÉTODO 2 - Desde Configuración:\n' +
    '1️⃣ Ve a facebook.com/tu-pagina\n' +
    '2️⃣ Haz clic en "Configuración" (⚙️)\n' +
    '3️⃣ El Page ID aparece en la parte superior de la página\n\n' +
    '🔹 MÉTODO 3 - Herramienta Online:\n' +
    '1️⃣ Ve a: https://findmyfbid.in/\n' +
    '2️⃣ Pega la URL de tu página de Facebook\n' +
    '3️⃣ Copia el ID que te muestra\n\n' +
    '═══════════════════════════════════\n\n' +
    '📝 Una vez que tengas el Page ID:\n' +
    '✅ Pégalo en la celda B11 de la hoja "📲 Crear Campaña WhatsApp"\n' +
    '✅ Debe ser un número largo (ejemplo: 123456789012345)\n' +
    '✅ NO incluyas espacios ni caracteres especiales';

  ui.alert('ℹ️ Cómo Obtener el Page ID', mensaje, ui.ButtonSet.OK);
}

/**
 * Valida si un Page ID es válido consultando la API de Meta
 */
function validarPageID() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('📲 Crear Campaña WhatsApp');

  if (!hoja) {
    ui.alert('❌ Error', 'Primero crea la hoja de configuración.', ui.ButtonSet.OK);
    return;
  }

  const pageId = hoja.getRange('B11').getValue();

  if (!pageId || pageId === '👉 PEGA AQUÍ EL PAGE ID') {
    ui.alert(
      '❌ Error',
      '⚠️ PASO 2: Page ID de Facebook\n\n' +
      'Primero debes ingresar un Page ID en la celda B11.\n\n' +
      '🔍 ¿Cómo obtenerlo?\n' +
      'Ve al menú: Creador de Campañas WhatsApp → Obtener Page ID de Facebook',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const config = obtenerConfiguracion();
    const token = config.META.ACCESS_TOKEN;

    const endpoint = `https://graph.facebook.com/v22.0/${pageId}?fields=name,id&access_token=${token}`;
    const response = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      ui.alert('❌ Page ID Inválido', `El Page ID "${pageId}" no es válido o no tienes acceso.\n\nError: ${result.error.message}`, ui.ButtonSet.OK);
      return;
    }

    ui.alert('✅ Page ID Válido', `Página: ${result.name}\nID: ${result.id}\n\n¡El Page ID es correcto!`, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}
