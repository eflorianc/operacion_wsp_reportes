/**
 * ==================== LÓGICA VENTAS (COMPATIBLE CON ALIAS) ====================
 */

function actualizarDatosVentas() {
  const ui = SpreadsheetApp.getUi();
  const config = obtenerConfiguracion();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDestino = ss.getSheetByName('📈 Datos Meta Ads');
  
  if (config.VENTAS_BOT.HOJAS.length === 0) {
    ui.alert('⚠️ No hay hojas de ventas configuradas.');
    return;
  }
  
  try {
    const tasasCambio = obtenerTasasDeCambio();
    const ventasConsolidadas = obtenerVentasDeTodosLosSpreadsheets(config, tasasCambio);
    
    const lastRow = sheetDestino.getLastRow();
    if (lastRow < 2) return;
    
    const rango = sheetDestino.getRange(2, 1, lastRow - 1, 18);
    const valores = rango.getValues();
    
    let tFact = 0, tGasto = 0, tUtil = 0, tAlc = 0, tCli = 0;
    
    const limpiar = (t) => t ? t.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() : "";
    
    const nuevos = valores.map(fila => {
      const prodReporte = limpiar(fila[1]); 
      if (prodReporte === 'TOTALES' || prodReporte === '') return fila;

      const gasto = parseFloat(fila[9]) || 0; 
      let factUSD = 0;
      let tasaRef = 0;
      
      for (const v of ventasConsolidadas) {
        const nombreCoincide = prodReporte.includes(v.producto);
        const paisCoincide = v.pais && prodReporte.includes(v.pais); 
        
        if (nombreCoincide && paisCoincide) {
           factUSD += v.montoUSD;
           tasaRef = v.tasaUsada; 
        }
      }
      
      const roas = gasto > 0 ? factUSD / gasto : 0;
      const util = factUSD - gasto;
      const roi = gasto > 0 ? util / gasto : 0;
      
      tFact += factUSD; tGasto += gasto; tUtil += util;
      tAlc += (parseInt(fila[4])||0); tCli += (parseInt(fila[5])||0);

      fila[10] = factUSD;
      fila[11] = tasaRef > 0 ? tasaRef : ''; 
      fila[12] = roas;
      fila[13] = util;
      fila[14] = roi;
      return fila;
    });
    
    rango.setValues(nuevos);
    
    if (nuevos.length > 0) {
      const formats = new Array(nuevos.length).fill(['$#,##0.00', '0.00', '0.00', '$#,##0.00', '0.00%']);
      sheetDestino.getRange(2, 11, nuevos.length, 5).setNumberFormats(formats);
    }
    
    actualizarTotales(sheetDestino, tFact, tUtil, tGasto, tAlc, tCli);
    SpreadsheetApp.flush();
    ss.toast(`✅ Ventas de ${config.VENTAS_BOT.HOJAS.length} hojas procesadas.`, 'Éxito', 5);
    
  } catch (e) {
    ui.alert('Error Fatal', e.message + '\n' + e.stack, ui.ButtonSet.OK);
  }
}

function obtenerVentasDeTodosLosSpreadsheets(config, tasasCambio) {
  let todasLasVentas = [];
  const hojas = config.VENTAS_BOT.HOJAS; // Lista de objetos {id, nombre}
  const mapaMonedas = config.MONEDAS_PAIS; 
  
  hojas.forEach(hojaObj => {
    // Soportamos tanto el objeto nuevo {id: '...'} como el string viejo por si acaso
    const id = hojaObj.id || hojaObj; 
    
    try {
      const ss = SpreadsheetApp.openById(id);
      const hoja = ss.getSheetByName(config.VENTAS_BOT.HOJA_COMPRAS);
      if (!hoja) return; 
      
      const lastRow = hoja.getLastRow();
      if (lastRow < 2) return;
      
      const cols = config.VENTAS_BOT.COLUMNAS_COMPRAS;
      const maxCol = Math.max(cols.VALOR, cols.NOMBRE_PRODUCTO, cols.PAIS);
      const datos = hoja.getRange(2, 1, lastRow - 1, maxCol).getValues();
      
      const limpiar = (t) => t ? t.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() : "";
      
      datos.forEach(fila => {
        const valorLocal = parseFloat(fila[cols.VALOR - 1]) || 0;
        const prodRaw = fila[cols.NOMBRE_PRODUCTO - 1]; 
        const paisRaw = fila[cols.PAIS - 1]; 
        
        if (prodRaw && valorLocal > 0) {
          const paisLimpio = limpiar(paisRaw);
          const prodLimpio = limpiar(prodRaw);
          
          const codigoMoneda = mapaMonedas[paisLimpio] || 'USD';
          const tasa = tasasCambio[codigoMoneda] || 1; 
          
          todasLasVentas.push({
            producto: prodLimpio,
            pais: paisLimpio,
            montoUSD: valorLocal / tasa,
            tasaUsada: tasa
          });
        }
      });
      
    } catch (error) {
      Logger.log(`Error leyendo spreadsheet ${id}: ${error.message}`);
    }
  });
  
  return todasLasVentas;
}