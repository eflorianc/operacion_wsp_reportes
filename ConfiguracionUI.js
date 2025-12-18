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