/**
 * ==================== VISUALIZACIÓN (ANCHO AJUSTADO PARA NOMBRES LARGOS) ====================
 */

function escribirResumenMetaEnHoja(resumen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📈 Datos Meta Ads');
  if (!sheet) return;
  
  const headers = [
    'Fecha', 'Producto (y País)', 'Campaign ID', 'Campaign Name', 
    'Alcance', 'Clics Únicos', 'Conversaciones', 
    'Gasto Base', 'IGV (18%)', 'Gasto Total', 
    'Facturación USD', 'T.C.', 'ROAS', 'Utilidad', 'ROI %',
    'Moneda Original', 'CPM', 'Estado'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground('#0b5394').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  const filas = [];
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  for (const p in resumen) {
    const d = resumen[p];
    const cpm = d.impresionesTotal > 0 ? (d.gastoTotal / d.impresionesTotal) * 1000 : 0;
    const estado = d.campanasActivas > 0 ? 'ACTIVO' : (d.totalCampanas > 0 ? 'PAUSADO' : 'N/A');
    const igv = d.gastoTotal * 0.18;

    filas.push([
      hoy, p, `${d.totalCampanas} camps`, 
      d.campanasActivas > 0 ? `${d.campanasActivas} activas` : 'Todas pausadas',
      d.alcanceTotal, d.clicsTotal, 0, 
      d.gastoTotal, igv, d.gastoTotal + igv,
      '', '', '', '', '', 'USD', cpm, estado 
    ]);
  }
  
  if (filas.length > 0) {
    const lr = sheet.getLastRow();
    if (lr > 1) sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).clearContent().clearFormat();
    sheet.getRange(2, 1, filas.length, headers.length).setValues(filas);
    sheet.getRange(2, 8, filas.length, 3).setNumberFormat('$#,##0.00');
    sheet.getRange(2, 17, filas.length, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(2, 5, filas.length, 2).setNumberFormat('#,##0');
  }
  SpreadsheetApp.flush();
}

function actualizarTotales(sheet, fact, util, gasto, alc, clic) {
  let row = sheet.getLastRow();
  const lastVal = sheet.getRange(row, 2).getValue();
  if (lastVal !== 'TOTALES') row++;
  
  const gBase = gasto / 1.18;
  const igv = gasto - gBase;
  const roas = gasto > 0 ? fact/gasto : 0;
  const roi = gasto > 0 ? util/gasto : 0;
  
  const vals = ['', 'TOTALES', '', '', alc, clic, 0, gBase, igv, gasto, fact, '', roas, util, roi, '', '', ''];
  const r = sheet.getRange(row, 1, 1, 18);
  r.setValues([vals]);
  r.setFontWeight('bold').setBackground('#e0e0e0').setBorder(true, false, true, false, false, false);
  sheet.getRange(row, 8, 1, 3).setNumberFormat('$#,##0.00'); 
  sheet.getRange(row, 11, 1, 1).setNumberFormat('$#,##0.00'); 
  sheet.getRange(row, 13, 1, 1).setNumberFormat('0.00');
  sheet.getRange(row, 14, 1, 1).setNumberFormat('$#,##0.00'); 
  sheet.getRange(row, 15, 1, 1).setNumberFormat('0.00%'); 
  SpreadsheetApp.flush();
}

function actualizarHojaConfiguracion() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('⚙️ Configuración');
    if (!sheet) sheet = ss.insertSheet('⚙️ Configuración');
    
    const config = obtenerConfiguracion();
    sheet.clear(); 
    
    const rows = [];
    const bgs = []; const fcs = []; const fws = [];
    const fill = (bg, fc, fw) => { bgs.push([bg,bg,bg,bg,bg]); fcs.push([fc,fc,fc,fc,fc]); fws.push([fw,fw,fw,fw,fw]); };

    // TÍTULO
    rows.push(['⚙️ PANEL DE CONFIGURACIÓN', '', '', '', '']); fill('#4285f4', '#fff', 'bold');
    rows.push(['', '', '', '', '']); fill(null, null, null);

    // TOKEN
    rows.push(['🔑 TOKEN META', '', '', '', '']); fill('#e8f0fe', '#000', 'bold');
    rows.push(['Estado:', config.META.ACCESS_TOKEN ? '✅ Listo' : '❌ Falta', '', '', '🔧 CONFIG']);
    bgs.push([null,null,null,null,'#34a853']); fcs.push([null, config.META.ACCESS_TOKEN?'green':'red',null,null,'#fff']); fws.push([null,'bold',null,null,'bold']);
    rows.push(['', '', '', '', '']); fill(null, null, null);

    // CUENTAS
    rows.push(['📊 CUENTAS PUBLICITARIAS', '', '', '', '']); fill('#e8f0fe', '#000', 'bold');
    if (!config.META.CUENTAS.length) { rows.push(['⚠️ Vacío', '', '', '', '']); fill('#fff3cd', '#000', 'normal'); }
    else { config.META.CUENTAS.forEach((c,i)=> { rows.push([i+1, c[0], '', '', c[1]]); fill(null, '#000', 'normal'); }); }
    rows.push(['', '', '', '', '']); fill(null, null, null);

    // --- SECCIÓN SPREADSHEETS (NOMBRE LARGO) ---
    rows.push(['📄 HOJAS DE VENTAS', '', '', '', '']); fill('#e8f0fe', '#000', 'bold');
    const hojas = config.VENTAS_BOT.HOJAS;
    
    if (hojas.length === 0) {
      rows.push(['Estado:', '❌ Ninguna conectada', '', '', '']); 
      fill(null, 'red', 'bold'); 
    } else {
      hojas.forEach((h, i) => {
         const nombre = h.nombre || `Hoja ${i+1}`;
         // Ya no mostramos plataforma en columna B porque está en el nombre
         const idCorto = '...' + h.id.substring(h.id.length - 6);
         
         rows.push([nombre, '', `✅ Conectado (${idCorto})`, '', '']);
         fill(null, 'green', 'bold'); 
      });
    }
    rows.push(['', '', '', '', '']); fill(null, null, null);
    // ------------------------------------------

    // PRODUCTOS
    rows.push(['🎯 PRODUCTOS Y PAÍSES', '', '', '', '']); fill('#e8f0fe', '#000', 'bold');
    rows.push(['PRODUCTO (LLAVE)', 'KEYWORDS', 'PAÍS', '', '']); fill('#f3f6f4', '#000', 'bold'); 
    for (const p in config.PRODUCTOS) {
      const pais = config.PRODUCTOS[p].pais || 'GLOBAL';
      rows.push([p, config.PRODUCTOS[p].palabrasClave.join(', '), pais, '', '']);
      fill(null, '#000', 'normal');
    }

    if (rows.length > 0) {
      const rng = sheet.getRange(1, 1, rows.length, 5);
      rng.setValues(rows);
      rng.setBackgrounds(bgs); rng.setFontColors(fcs); rng.setFontWeights(fws);
      
      sheet.getRange(1, 1, 1, 5).merge().setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(18);
      sheet.setRowHeight(1, 50);
      
      // AJUSTE DE ANCHO DE COLUMNA A PARA NOMBRES LARGOS
      sheet.setColumnWidth(1, 350); 
      
      rows.forEach((r, i) => {
        const idx = i + 1;
        const val = r[0] ? r[0].toString() : '';
        
        if (val.includes('PANEL') || val.includes('TOKEN') || val.includes('CUENTAS') || val.includes('HOJAS DE VENTAS') || val.includes('PRODUCTOS')) 
           sheet.getRange(idx, 1, 1, 5).merge();
           
        if (val === 'Estado:' || val === 'ID:') sheet.getRange(idx, 2, 1, 3).merge();
        
        // Merge para filas de hojas: Como el nombre está en A, y el estado en C, fusionamos B y C para que el estado se pegue o quede limpio.
        // O mejor: Dejamos A sola (Nombre largo) y fusionamos B+C+D para el estado
        const esHoja = hojas.some(h => h.nombre === val || `Hoja ${i}` === val); 
        if (esHoja) {
           sheet.getRange(idx, 3, 1, 3).merge(); // Fusiona el estado en C, D, E
        } else if (typeof r[0] === 'number') {
           sheet.getRange(idx, 2, 1, 3).merge(); // Cuentas
        }
      });
      sheet.getRange(4, 5).setHorizontalAlignment('center');
    }
    SpreadsheetApp.flush();
  } catch (e) { Logger.log(e); }
}