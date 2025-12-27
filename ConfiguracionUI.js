/**
 * ==================== FUNCIONES DE CONFIGURACIÓN ====================
 */

/**
 * Configura el token de acceso de Meta Ads
 */
function configurarToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔑 Token de Acceso de Meta Ads:', 'Ingrese su token de acceso:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token && token.length > 5) {
      PropertiesService.getScriptProperties().setProperty('META_TOKEN', token);

      const savedToken = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
      if (savedToken && savedToken.length > 5) {
        ui.alert('✅ Token guardado correctamente');
        actualizarHojaConfiguracion();
      } else {
        ui.alert('❌ Error al guardar el token');
      }
    } else {
      ui.alert('❌ Token inválido');
    }
  }
}

/**
 * Configura las cuentas publicitarias de Meta Ads
 */
function configurarCuentas() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const cuentasActuales = JSON.parse(props.getProperty('META_CUENTAS') || '[]');
  let mensaje = 'Cuentas actuales:\n\n';

  if (cuentasActuales.length === 0) {
    mensaje = 'No hay cuentas configuradas.\n\n';
  } else {
    cuentasActuales.forEach((c, i) => {
      mensaje += `${i + 1}. ${c[0]} - ${c[1]}\n`;
    });
  }

  mensaje += '\n¿Desea agregar una nueva cuenta?';

  const respuesta = ui.alert('📊 Configurar Cuentas', mensaje, ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    const idCuenta = ui.prompt('ID de Cuenta Publicitaria:', 'Ingrese el ID (act_XXXXXXXXX):', ui.ButtonSet.OK_CANCEL);

    if (idCuenta.getSelectedButton() == ui.Button.OK) {
      const id = idCuenta.getResponseText().trim();

      const nombreCuenta = ui.prompt('Nombre de la Cuenta:', 'Ingrese un nombre descriptivo:', ui.ButtonSet.OK_CANCEL);

      if (nombreCuenta.getSelectedButton() == ui.Button.OK) {
        const nombre = nombreCuenta.getResponseText().trim();

        if (id && nombre) {
          cuentasActuales.push([id, nombre]);
          props.setProperty('META_CUENTAS', JSON.stringify(cuentasActuales));
          ui.alert('✅ Cuenta agregada correctamente');
          actualizarHojaConfiguracion();
        } else {
          ui.alert('❌ Datos inválidos');
        }
      }
    }
  }
}

/**
 * Configura los spreadsheets de ventas
 */
function configurarSpreadsheetVentas() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const hojasActuales = JSON.parse(props.getProperty('VENTAS_SPREADSHEET_IDS') || '[]');
  let mensaje = 'Hojas de ventas actuales:\n\n';

  if (hojasActuales.length === 0) {
    mensaje = 'No hay hojas de ventas configuradas.\n\n';
  } else {
    hojasActuales.forEach((h, i) => {
      const nombre = h.nombre || `Hoja ${i + 1}`;
      mensaje += `${i + 1}. ${nombre}\n`;
    });
  }

  mensaje += '\n¿Desea agregar una nueva hoja de ventas?';

  const respuesta = ui.alert('📄 Configurar Hojas de Ventas', mensaje, ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    const idHoja = ui.prompt('ID del Spreadsheet:', 'Ingrese el ID del spreadsheet de ventas:', ui.ButtonSet.OK_CANCEL);

    if (idHoja.getSelectedButton() == ui.Button.OK) {
      const id = idHoja.getResponseText().trim();

      const nombreHoja = ui.prompt('Nombre de la Hoja:', 'Ingrese un nombre descriptivo (ej: PERÚ, COLOMBIA):', ui.ButtonSet.OK_CANCEL);

      if (nombreHoja.getSelectedButton() == ui.Button.OK) {
        const nombre = nombreHoja.getResponseText().trim();

        if (id && nombre) {
          hojasActuales.push({ id: id, nombre: nombre });
          props.setProperty('VENTAS_SPREADSHEET_IDS', JSON.stringify(hojasActuales));
          ui.alert('✅ Hoja de ventas agregada correctamente');
          actualizarHojaConfiguracion();
        } else {
          ui.alert('❌ Datos inválidos');
        }
      }
    }
  }
}

/**
 * Configura palabras clave y productos
 */
function configurarPalabrasClave() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const productosActuales = JSON.parse(props.getProperty('PRODUCTOS_CONFIG') || '{}');
  let mensaje = 'Productos actuales:\n\n';

  if (Object.keys(productosActuales).length === 0) {
    mensaje = 'No hay productos configurados.\n\n';
  } else {
    for (const prod in productosActuales) {
      const pais = productosActuales[prod].pais || 'GLOBAL';
      const palabras = productosActuales[prod].palabrasClave.join(', ');
      mensaje += `• ${prod} (${pais})\n  Keywords: ${palabras}\n\n`;
    }
  }

  mensaje += '¿Desea agregar un nuevo producto?';

  const respuesta = ui.alert('🎯 Configurar Productos', mensaje, ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    const nombreProducto = ui.prompt('Nombre del Producto (LLAVE):', 'Ingrese el nombre EXACTO del producto:', ui.ButtonSet.OK_CANCEL);

    if (nombreProducto.getSelectedButton() == ui.Button.OK) {
      const producto = nombreProducto.getResponseText().trim().toUpperCase();

      const paisProducto = ui.prompt('País del Producto:', 'Ingrese el país (PERÚ, COLOMBIA, MÉXICO, etc.):', ui.ButtonSet.OK_CANCEL);

      if (paisProducto.getSelectedButton() == ui.Button.OK) {
        const pais = paisProducto.getResponseText().trim().toUpperCase();

        const palabrasProducto = ui.prompt('Palabras Clave:', 'Ingrese las palabras clave separadas por comas:', ui.ButtonSet.OK_CANCEL);

        if (palabrasProducto.getSelectedButton() == ui.Button.OK) {
          const palabras = palabrasProducto.getResponseText().split(',').map(p => p.trim().toUpperCase()).filter(p => p);

          if (producto && palabras.length > 0) {
            productosActuales[producto] = {
              pais: pais,
              palabrasClave: palabras
            };

            props.setProperty('PRODUCTOS_CONFIG', JSON.stringify(productosActuales));
            ui.alert('✅ Producto agregado correctamente');
            actualizarHojaConfiguracion();
          } else {
            ui.alert('❌ Datos inválidos');
          }
        }
      }
    }
  }
}

/**
 * Almacena el token de Meta Ads en las propiedades del script.
 * Esta función se usa para forzar la escritura y lectura del token
 * si hay problemas de almacenamiento.
 */
function forzarAlmacenamientoToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔑 RE-INGRESE SU Token de Acceso de Meta Ads (Necesario para el análisis de creativos):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const token = response.getResponseText();
    if (token && token.length > 5) {
      // 1. Almacenamiento forzado en la clave principal
      PropertiesService.getScriptProperties().setProperty('META_TOKEN', token);

      // 2. Verificación inmediata
      const savedToken = PropertiesService.getScriptProperties().getProperty('META_TOKEN');

      if (savedToken && savedToken.length > 5) {
        ui.alert('✅ TOKEN GUARDADO', 'El token se ha guardado y verificado correctamente. El sistema ahora debería poder leerlo.', ui.ButtonSet.OK);
      } else {
        ui.alert('❌ ERROR AL GUARDAR', 'Hubo un error al almacenar el token. Por favor, revise los permisos o inténtelo de nuevo.', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Token inválido. Inténtelo de nuevo.', ui.ButtonSet.OK);
    }
  }
}