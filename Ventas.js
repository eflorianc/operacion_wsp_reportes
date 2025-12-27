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

      // LÓGICA SIMPLIFICADA DE MATCHING
      // Formato reporte: "PRODUCTO X - PERU"
      // Formato ventas: "PRODUCTO X" (solo producto)
      // País ventas: "PERU" (en columna separada)

      for (const v of ventasConsolidadas) {
        // REGLA 1: El producto de ventas debe estar contenido en el nombre del reporte
        if (prodReporte.includes(v.producto)) {

          // REGLA 2: Si hay país en ventas, también debe estar en el reporte
          if (v.pais && v.pais !== '') {
            // Si el país de ventas está en el reporte, es un match válido
            if (prodReporte.includes(v.pais)) {
              factUSD += v.montoUSD;
              tasaRef = v.tasaUsada;
            }
            // Si el país NO está en el reporte, NO es match (evita duplicados)
          } else {
            // Si no hay país en ventas, asumir que es match válido
            factUSD += v.montoUSD;
            tasaRef = v.tasaUsada;
          }
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
            montoLocal: valorLocal,
            moneda: codigoMoneda,
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

/**
 * Consulta ventas de un producto específico
 */
function consultarVentasPorProducto() {
  const ui = SpreadsheetApp.getUi();
  const config = obtenerConfiguracion();

  if (config.VENTAS_BOT.HOJAS.length === 0) {
    ui.alert('⚠️ No hay hojas de ventas configuradas.');
    return;
  }

  // Pedir nombre del producto
  const resp = ui.prompt(
    '🔍 Consultar Ventas por Producto',
    'Ingresa el nombre del producto (ej: KIT AMIGURUMIS CHESPIRITO - PERU):\n\nPuedes incluir el país en el nombre para filtrar.',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const busqueda = resp.getResponseText().trim();
  if (!busqueda) {
    ui.alert('❌ Error', 'Debes ingresar un nombre de producto.', ui.ButtonSet.OK);
    return;
  }

  try {
    const tasasCambio = obtenerTasasDeCambio();
    const ventasConsolidadas = obtenerVentasDeTodosLosSpreadsheets(config, tasasCambio);

    const limpiar = (t) => t ? t.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() : "";
    const busquedaLimpia = limpiar(busqueda);

    // Extraer país si está en la búsqueda
    let paisBuscado = '';
    const paises = ['PERU', 'COLOMBIA', 'MEXICO', 'ARGENTINA', 'CHILE', 'ECUADOR', 'BOLIVIA'];
    for (const pais of paises) {
      if (busquedaLimpia.includes(pais)) {
        paisBuscado = pais;
        break;
      }
    }

    // Filtrar ventas que coincidan
    const ventasFiltradas = ventasConsolidadas.filter(v => {
      const productoCoincide = v.producto.includes(busquedaLimpia) || busquedaLimpia.includes(v.producto);

      if (paisBuscado) {
        // Si hay país en la búsqueda, debe coincidir
        return productoCoincide && v.pais === paisBuscado;
      } else {
        // Si no hay país en la búsqueda, solo verificar producto
        return productoCoincide;
      }
    });

    if (ventasFiltradas.length === 0) {
      ui.alert(
        '❌ No se encontraron ventas',
        `No se encontraron ventas para "${busqueda}".\n\nIntenta con otro nombre o verifica que el producto exista en el spreadsheet.`,
        ui.ButtonSet.OK
      );
      return;
    }

    // Agrupar por país y calcular totales
    const totalesPorPais = {};
    let totalVentas = 0;
    let totalUSD = 0;

    ventasFiltradas.forEach(v => {
      const pais = v.pais || 'SIN PAÍS';

      if (!totalesPorPais[pais]) {
        totalesPorPais[pais] = {
          ventas: 0,
          totalLocal: 0,
          totalUSD: 0,
          moneda: v.moneda,
          tasa: v.tasaUsada
        };
      }

      totalesPorPais[pais].ventas++;
      totalesPorPais[pais].totalLocal += v.montoLocal;
      totalesPorPais[pais].totalUSD += v.montoUSD;
      totalVentas++;
      totalUSD += v.montoUSD;
    });

    // Construir mensaje
    let mensaje = `📊 VENTAS DE: "${busqueda}"\n\n`;
    mensaje += '━'.repeat(50) + '\n\n';

    mensaje += `📈 RESUMEN GLOBAL:\n`;
    mensaje += `• Total de ventas: ${totalVentas}\n`;
    mensaje += `• Total en USD: $${totalUSD.toFixed(2)}\n\n`;

    mensaje += '━'.repeat(50) + '\n\n';
    mensaje += '🌎 DETALLE POR PAÍS:\n\n';

    for (const pais in totalesPorPais) {
      const datos = totalesPorPais[pais];
      mensaje += `${pais}:\n`;
      mensaje += `  • Número de ventas: ${datos.ventas}\n`;
      mensaje += `  • Total: ${datos.totalLocal.toFixed(2)} ${datos.moneda}\n`;
      mensaje += `  • Tasa de cambio: ${datos.tasa.toFixed(2)}\n`;
      mensaje += `  • Total USD: $${datos.totalUSD.toFixed(2)}\n\n`;
    }

    // Mostrar primeras 5 ventas como ejemplo
    mensaje += '━'.repeat(50) + '\n\n';
    mensaje += '📋 PRIMERAS 5 VENTAS ENCONTRADAS:\n\n';

    ventasFiltradas.slice(0, 5).forEach((v, i) => {
      mensaje += `${i + 1}. ${v.producto} (${v.pais})\n`;
      mensaje += `   ${v.montoLocal.toFixed(2)} ${v.moneda} = $${v.montoUSD.toFixed(2)} USD\n\n`;
    });

    if (ventasFiltradas.length > 5) {
      mensaje += `... y ${ventasFiltradas.length - 5} ventas más\n\n`;
    }

    Logger.log(mensaje);

    // Mensaje corto para el alert
    let alertMsg = `📊 Producto: "${busqueda}"\n\n`;
    alertMsg += `📈 Total de ventas: ${totalVentas}\n`;
    alertMsg += `💵 Total USD: $${totalUSD.toFixed(2)}\n\n`;
    alertMsg += 'Detalles por país:\n';
    for (const pais in totalesPorPais) {
      const datos = totalesPorPais[pais];
      alertMsg += `• ${pais}: ${datos.ventas} ventas\n`;
    }
    alertMsg += '\n📝 Ver log completo: Extensiones > Apps Script > Ver registros';

    ui.alert('✅ Consulta Completada', alertMsg, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', e.message + '\n' + e.stack, ui.ButtonSet.OK);
  }
}

/**
 * Función de diagnóstico para ver qué ventas se están matcheando
 */
function diagnosticarMatchingVentas() {
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

    let mensaje = '🔍 DIAGNÓSTICO DE MATCHING DE VENTAS\n\n';
    mensaje += `Total de ventas encontradas: ${ventasConsolidadas.length}\n\n`;

    // Mostrar primeras 10 ventas
    mensaje += '📊 PRIMERAS 10 VENTAS DEL SPREADSHEET:\n';
    mensaje += '━'.repeat(50) + '\n';

    ventasConsolidadas.slice(0, 10).forEach((v, i) => {
      mensaje += `${i + 1}. Producto: "${v.producto}"\n`;
      mensaje += `   País: "${v.pais}"\n`;
      mensaje += `   Monto Local: ${v.montoLocal.toFixed(2)} ${v.moneda}\n`;
      mensaje += `   Tasa: ${v.tasaUsada.toFixed(2)}\n`;
      mensaje += `   USD: $${v.montoUSD.toFixed(2)}\n\n`;
    });

    // Obtener productos del reporte
    const lastRow = sheetDestino.getLastRow();
    if (lastRow >= 2) {
      mensaje += '\n📈 PRODUCTOS EN EL REPORTE:\n';
      mensaje += '━'.repeat(50) + '\n';

      const rango = sheetDestino.getRange(2, 2, Math.min(lastRow - 1, 10), 1);
      const productos = rango.getValues();

      productos.forEach((p, i) => {
        if (p[0] && p[0] !== 'TOTALES') {
          mensaje += `${i + 1}. "${p[0]}"\n`;
        }
      });
    }

    // Mostrar configuración de productos
    mensaje += '\n\n⚙️ PRODUCTOS CONFIGURADOS:\n';
    mensaje += '━'.repeat(50) + '\n';

    const productosConfig = config.PRODUCTOS || {};
    if (Object.keys(productosConfig).length === 0) {
      mensaje += 'No hay productos configurados\n';
    } else {
      for (const prod in productosConfig) {
        mensaje += `• ${prod} (${productosConfig[prod].pais || 'SIN PAÍS'})\n`;
        mensaje += `  Keywords: ${productosConfig[prod].palabrasClave.join(', ')}\n\n`;
      }
    }

    // Mostrar tasas de cambio
    mensaje += '\n💱 TASAS DE CAMBIO ACTUALES:\n';
    mensaje += '━'.repeat(50) + '\n';
    mensaje += `PEN: ${tasasCambio.PEN || 'N/A'}\n`;
    mensaje += `COP: ${tasasCambio.COP || 'N/A'}\n`;
    mensaje += `MXN: ${tasasCambio.MXN || 'N/A'}\n`;

    Logger.log(mensaje);
    ui.alert('Diagnóstico Completo', 'Revisa el Log del script (Ver > Registros) para ver el diagnóstico completo.\n\nPrimeras ventas:\n' + mensaje.substring(0, 500) + '...', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', e.message + '\n' + e.stack, ui.ButtonSet.OK);
  }
}