/**
 * ==================== MENÚ PRINCIPAL Y SETUP ====================
 * Punto de entrada del sistema.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Reporte de Ventas')
    .addItem('🔄 Actualizar Reporte Completo', 'actualizarReporteCompleto')
    .addSeparator()
    .addItem('📈 Solo Meta Ads', 'actualizarDatosMeta')
    .addItem('💰 Solo Ventas', 'actualizarDatosVentas')
    .addItem('🎯 Extraer Gasto por Producto', 'extraerGastoPorAnuncio')
    .addItem('📊 Extraer TODOS los Rangos', 'extraerTodosLosRangos')
    // --- NUEVO PUNTO 6: ACCESO AL ANALIZADOR DE CREATIVOS ---
    .addSeparator()
    .addItem('🎨 Panel de Creativos (Ad-Level)', 'runCreativeAnalysis') // Llama a la nueva función
    // --------------------------------------------------------
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 Diagnósticos')
      .addItem('🛒 Consultar Ventas por Producto', 'consultarVentasPorProducto')
      .addSeparator()
      .addItem('📊 Matching de Ventas', 'diagnosticarMatchingVentas')
      .addItem('📈 Extracción de Rangos', 'diagnosticarExtraccionRangos')
      .addSeparator()
      .addItem('💬 Contar Mensajes AYER PERÚ → X1', 'contarMensajesAyerPeru')
      .addItem('🎯 Contar Mensajes AD ID AYER → X2', 'contarMensajesPorAdId')
      .addItem('📊 Contar Mensajes AD ID TOTAL → X3', 'contarMensajesTotalesPorAdId')
      .addSeparator()
      .addItem('🔗 Clics Únicos en Enlace → X1', 'extraerClicsUnicosEnlaceDesdeMetaParaAdId')
      .addItem('📋 Ver Todos los Campos de Clics', 'diagnosticarCamposClicsMeta')
      .addSeparator()
      .addItem('🆔 Buscar AD ID Específico', 'diagnosticarAdId')
      .addItem('🔬 Diagnóstico Completo AD ID', 'diagnosticarAdIdCompleto'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 Consultar Gasto de Anuncio')
      .addItem('✅ Obtener Gasto por AD ID (Recomendado)', 'obtenerGastoPorAdId')
      .addSeparator()
      .addItem('🧪 PRUEBA: Buscar Primer Anuncio', 'buscarPrimerAnuncio')
      .addSeparator()
      .addItem('📚 Buscar en Biblioteca (Requiere permisos)', 'buscarGastoEnBibliotecaAnuncios')
      .addItem('🔎 Buscar por Término (Requiere permisos)', 'buscarAnunciosPorTermino'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📲 Creador de Campañas WhatsApp')
      .addItem('📋 Crear Hoja de Configuración', 'crearHojaCampanas')
      .addItem('📄 Obtener Page ID de Facebook', 'mostrarPageIDsDisponibles')
      .addItem('✅ Validar Page ID', 'validarPageID')
      .addSeparator()
      .addItem('▶️ Crear Campaña', 'crearCampanaWhatsApp'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Configuración')
      .addItem('🔑 Configurar Token de Meta', 'configurarToken')
      .addItem('📊 Configurar Cuentas Publicitarias', 'configurarCuentas')
      .addItem('📄 Configurar Spreadsheet de Ventas', 'configurarSpreadsheetVentas')
      .addItem('🔤 Configurar Palabras Clave', 'configurarPalabrasClave')
      .addSeparator()
      .addItem('👁️ Actualizar Panel Visual', 'actualizarHojaConfiguracion'))
    .addSeparator()
    .addItem('🔍 Diagnosticar Campañas', 'diagnosticarCampanasYPalabras')
    .addItem('🏗️ Inicializar Sistema', 'instalarSistema')
    .addToUi();


}

function actualizarReporteCompleto() {
  actualizarDatosMeta();
  actualizarDatosVentas();
}

function diagnosticarCampanasYPalabras() {
  const ui = SpreadsheetApp.getUi();
  const config = obtenerConfiguracion();
  if (!config.META.ACCESS_TOKEN) return ui.alert('Falta Token');
  try {
    const cta = config.META.CUENTAS[0][0];
    const url = `${config.META.ENDPOINT}/${config.META.API_VERSION}/${cta}/campaigns?fields=name&access_token=${config.META.ACCESS_TOKEN}&limit=5`;
    const r = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
    ui.alert('Muestra Campañas', r.getContentText().substring(0,500), ui.ButtonSet.OK);
  } catch(e) { ui.alert(e.message); }
}

function instalarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojas = [
    // Agregar la nueva hoja al inicializador para que siempre exista
    { nombre: '🎨 Panel de Creativos', color: '#6aa84f' },
    { nombre: '📊 Dashboard', color: '#4285f4' },
    { nombre: '📈 Datos Meta Ads', color: '#34a853' },
    { nombre: '💰 Datos Ventas', color: '#fbbc04' },
    { nombre: '🎯 Reporte General', color: '#ea4335' },
    { nombre: '⚙️ Configuración', color: '#9e9e9e' }
  ];

  hojas.forEach(h => {
    let sheet = ss.getSheetByName(h.nombre);
    if (!sheet) sheet = ss.insertSheet(h.nombre);
    sheet.setTabColor(h.color);
  });

  escribirResumenMetaEnHoja({});
  actualizarHojaConfiguracion();

  SpreadsheetApp.getUi().alert('✅ Sistema Instalado con Panel de Configuración');
}

/**
 * Función para ejecutar análisis de creativos (Ad-Level)
 * TODO: Implementar la lógica de análisis de creativos
 */
function runCreativeAnalysis() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '🎨 Panel de Creativos',
    'Esta funcionalidad está en desarrollo.\n\nPróximamente podrás analizar el rendimiento de tus creativos a nivel de anuncio.',
    ui.ButtonSet.OK
  );
}