/**
 * ==================== LÓGICA META ADS (FILTRO PAÍS) ====================
 */

function actualizarDatosMeta() {
  const ui = SpreadsheetApp.getUi();
  const config = obtenerConfiguracion();
  
  if (!config.META.ACCESS_TOKEN) return ui.alert('Falta Token');
  if (config.META.CUENTAS.length === 0) return ui.alert('Faltan Cuentas');
  
  try {
    const resumen = extraerYAgruparPorProducto();
    escribirResumenMetaEnHoja(resumen);
  } catch (error) {
    ui.alert('Error Meta', error.message, ui.ButtonSet.OK);
  }
}

function extraerYAgruparPorProducto() {
  const config = obtenerConfiguracion();
  const resumen = {};
  const tcActual = obtenerTasaActual(); 

  // Inicializar contadores por cada llave (Producto - País)
  for (const p in config.PRODUCTOS) {
    resumen[p] = {
      alcanceTotal: 0, clicsTotal: 0, gastoTotal: 0, impresionesTotal: 0,
      campanasActivas: 0, totalCampanas: 0
    };
  }
  
  config.META.CUENTAS.forEach((cuenta) => {
    const accId = cuenta[0];
    const accMoneda = cuenta[1];
    
    try {
      // Pedimos nombre y estado
      const url = `${config.META.ENDPOINT}/${config.META.API_VERSION}/${accId}/campaigns?fields=id,name,status&access_token=${config.META.ACCESS_TOKEN}&limit=500`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) return;
      
      const data = JSON.parse(resp.getContentText()).data || [];
      
      data.forEach(camp => {
        // AQUI ESTA LA MAGIA: Pasamos el objeto completo de config para validar país
        const productoKey = identificarProductoDeCampana(camp.name, config.PRODUCTOS);
        
        if (productoKey) {
          resumen[productoKey].totalCampanas++;
          if (camp.status === 'ACTIVE') resumen[productoKey].campanasActivas++;
          
          const insights = obtenerInsightsCampana(camp.id, config.META.ACCESS_TOKEN, config.META);
          if (insights) {
            resumen[productoKey].alcanceTotal += parseInt(insights.reach) || 0;
            resumen[productoKey].impresionesTotal += parseInt(insights.impressions) || 0;
            resumen[productoKey].clicsTotal += parseInt(insights.clicks) || 0; 
            
            let gasto = parseFloat(insights.spend) || 0;
            if (accMoneda !== 'USD') gasto = gasto / tcActual;
            resumen[productoKey].gastoTotal += gasto;
          }
        }
      });
    } catch (e) { Logger.log(e); }
  });
  
  return resumen;
}

function identificarProductoDeCampana(nombreCampana, productosConfig) {
  if (!nombreCampana) return null;
  const nombreUpper = nombreCampana.toUpperCase(); // Todo a mayúsculas para comparar
  
  for (const key in productosConfig) {
    const config = productosConfig[key];
    const keywords = config.palabrasClave;
    const paisRequerido = config.pais ? config.pais.toUpperCase() : null;
    
    // 1. Verificar si tiene el PAIS (si está configurado)
    if (paisRequerido && !nombreUpper.includes(paisRequerido)) {
      continue; // Si la campaña no dice "PERU" y el config pide "PERU", saltamos
    }

    // 2. Verificar Keywords
    const matchKeyword = keywords.some(kw => nombreUpper.includes(kw.toUpperCase().trim()));
    
    if (matchKeyword) return key; // Retorna ej: "KIT FINANZAS - PERU"
  }
  return null;
}

function obtenerInsightsCampana(campaignId, token, metaConfig) {
  const url = `${metaConfig.ENDPOINT}/${metaConfig.API_VERSION}/${campaignId}/insights?fields=spend,reach,impressions,clicks&date_preset=maximum&access_token=${token}`;
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const d = JSON.parse(r.getContentText());
    return (d.data && d.data.length > 0) ? d.data[0] : null;
  } catch (e) { return null; }
}