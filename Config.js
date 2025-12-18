/**
 * ==================== CONFIGURACIÓN CENTRAL (ALIAS DE PAÍSES) ====================
 */

function obtenerConfiguracion() {
  const props = PropertiesService.getScriptProperties();
  
  // Recuperamos la lista. Hacemos una migración automática si antes eran solo IDs sueltos.
  const rawIds = props.getProperty('VENTAS_SPREADSHEET_IDS');
  let listaHojas = [];
  
  try {
    const parsed = rawIds ? JSON.parse(rawIds) : [];
    
    // Normalización: Si el dato guardado es antiguo (solo ID), lo convertimos a objeto
    listaHojas = parsed.map((item, index) => {
      if (typeof item === 'string') {
        return { id: item, nombre: `Hoja ${index + 1}` }; // Nombre genérico temporal
      }
      return item; // Ya es un objeto con nombre
    });
    
  } catch (e) {
    listaHojas = [];
  }

  const config = {
    META: {
      ACCESS_TOKEN: props.getProperty('META_TOKEN') || '',
      API_VERSION: 'v21.0',
      ENDPOINT: 'https://graph.facebook.com',
      CUENTAS: JSON.parse(props.getProperty('META_CUENTAS') || '[]')
    },
    
    VENTAS_BOT: {
      HOJAS: listaHojas, // Nueva estructura: [{id: '...', nombre: 'PERÚ'}, ...]
      HOJA_COMPRAS: 'Compras',
      
      COLUMNAS_COMPRAS: {
        FECHA: 3,             
        NOMBRE: 2,            
        VALOR: 4,             
        NOMBRE_PRODUCTO: 10,  
        PAIS: 11              
      }
    },
    
    EXCHANGE: {
      API_URL: 'https://api.exchangerate-api.com/v4/latest/USD',
      CACHE_DURATION: 21600 
    },
    
    MONEDAS_PAIS: {
      'PERU': 'PEN',
      'COLOMBIA': 'COP',
      'MEXICO': 'MXN',
      'CHILE': 'CLP',
      'ARGENTINA': 'ARS',
      'ECUADOR': 'USD',
      'PANAMA': 'USD',
      'ESTADOS UNIDOS': 'USD'
    },
    
    PRODUCTOS: JSON.parse(props.getProperty('PRODUCTOS_CONFIG') || '{}')
  };
  
  return config;
}