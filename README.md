# 📊 Sistema de Reportes - Meta Ads & Ventas

Sistema automatizado de reportes integrado con Meta Ads Manager para análisis de campañas publicitarias y facturación de ventas multi-país.

## 🚀 Características

- **Integración con Meta Ads API**: Extrae datos de rendimiento de campañas en tiempo real
- **Multi-país**: Soporte para Perú, Colombia, México, Chile, Argentina, Ecuador, Panamá y Estados Unidos
- **Conversión de Monedas**: Cálculo automático de tasas de cambio (PEN, COP, MXN, CLP, ARS, USD)
- **Análisis de Creativos**: Reportes a nivel de anuncio (Ad-Level) con métricas clave
- **Cálculo de ROI y ROAS**: Métricas financieras automáticas incluyendo IGV
- **Panel de Configuración Visual**: Interfaz amigable para gestionar credenciales y productos

## 📋 Requisitos Previos

- Cuenta de Google con acceso a Google Sheets
- Token de acceso de Meta Ads Manager
- IDs de cuentas publicitarias de Meta
- Node.js instalado (para usar clasp)
- [Clasp](https://github.com/google/clasp) - Herramienta CLI de Google Apps Script

```bash
npm install -g @google/clasp
```

## 🛠️ Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/eflorianc/operacion_wsp_reportes.git
cd operacion_wsp_reportes
```

### 2. Autenticarse con clasp

```bash
clasp login
```

### 3. Crear un nuevo proyecto de Google Apps Script

```bash
clasp create --type sheets --title "Sistema de Reportes"
```

### 4. Subir el código a Google Apps Script

```bash
clasp push
```

### 5. Abrir el proyecto en el navegador

```bash
clasp open
```

## ⚙️ Configuración

### 1. Inicializar el Sistema

1. Abre tu Google Spreadsheet vinculado
2. Verás el menú **"📊 Reporte de Ventas"**
3. Ve a: **🏗️ Inicializar Sistema**

Esto creará las hojas necesarias:
- 📈 Datos Meta Ads
- ⚙️ Configuración

### 2. Configurar Token de Meta

1. Ve a: **⚙️ Configuración → 🔑 Configurar Token de Meta**
2. Ingresa tu token de acceso de Meta Ads

**¿Dónde obtener el token?**
- Meta Business Suite → Configuración → Herramientas empresariales → Tokens de acceso

### 3. Configurar Cuentas Publicitarias

1. Ve a: **⚙️ Configuración → 📊 Configurar Cuentas Publicitarias**
2. Ingresa los IDs de tus cuentas (formato: `act_123456789`)
3. Indica la moneda de cada cuenta (USD, PEN, etc.)

### 4. Configurar Hojas de Ventas

1. Ve a: **⚙️ Configuración → 📄 Configurar Spreadsheet de Ventas**
2. Ingresa los IDs de las hojas de cálculo donde están tus ventas
3. Asigna un nombre a cada hoja (ej: "PERÚ", "COLOMBIA")

### 5. Configurar Palabras Clave de Productos

1. Ve a: **⚙️ Configuración → 🔤 Configurar Palabras Clave**
2. Define productos y sus palabras clave de identificación
3. Asigna países a cada producto (opcional)

**Ejemplo:**
```json
{
  "KIT FINANZAS - PERU": {
    "palabrasClave": ["FINANZAS", "KIT"],
    "pais": "PERU"
  }
}
```

## 📊 Uso

### Actualizar Reportes

**Reporte Completo:**
```
📊 Reporte de Ventas → 🔄 Actualizar Reporte Completo
```

**Solo Meta Ads:**
```
📊 Reporte de Ventas → 📈 Solo Meta Ads
```

**Solo Ventas:**
```
📊 Reporte de Ventas → 💰 Solo Ventas
```

### Extraer Todos los Rangos

```
📊 Reporte de Ventas → 📊 Extraer TODOS los Rangos
```

Genera un reporte consolidado con múltiples rangos de tiempo en una sola tabla, incluyendo:
- Datos por anuncio
- Totales por rango
- Filtros automáticos
- Colores por rango

## 📁 Estructura del Proyecto

```
📦 operacion_wsp_reportes
├── 📄 Config.js                 # Configuración central (países, monedas)
├── 📄 ConfiguracionUI.js        # Funciones de configuración de UI
├── 📄 ExtraccionCreativos.js    # Extracción de datos de Meta Ads
├── 📄 Main.js                   # Menú principal y punto de entrada
├── 📄 MetaAds.js                # Lógica de integración con Meta API
├── 📄 Reportes.js               # Generación de reportes visuales
├── 📄 Utilidades.js             # Tasas de cambio y utilidades
├── 📄 Ventas.js                 # Procesamiento de datos de ventas
├── 📄 appsscript.json           # Configuración del proyecto Apps Script
├── 📄 .clasp.json               # Configuración de clasp
└── 📄 README.md                 # Este archivo
```

## 📈 Métricas Calculadas

### Columnas de Gasto
- **GASTO**: Inversión publicitaria base (USD)
- **IGV**: 18% del gasto
- **GASTO TOTAL**: GASTO + IGV

### Columnas de Rendimiento
- **ALCANCE**: Usuarios únicos alcanzados
- **CLICS**: Clics en anuncios
- **IMPRESIONES**: Total de visualizaciones
- **CPM**: Costo por mil impresiones (GASTO / IMPRESIONES × 1000)

### Columnas de Facturación
- **FACT USD**: Facturación en dólares americanos
- **# VENTAS**: Cantidad de ventas
- **T.C.**: Tasa de cambio aplicada

### Métricas Financieras
- **ROAS**: Return on Ad Spend (FACT USD / GASTO TOTAL)
- **UTILIDAD**: FACT USD - GASTO TOTAL
- **ROI**: Return on Investment (UTILIDAD / GASTO TOTAL)

## 🌍 Países y Monedas Soportadas

| País | Código Moneda |
|------|---------------|
| Perú | PEN |
| Colombia | COP |
| México | MXN |
| Chile | CLP |
| Argentina | ARS |
| Ecuador | USD |
| Panamá | USD |
| Estados Unidos | USD |

## 🔧 Funciones Principales

### `extraerTodosLosRangos()`
Extrae datos de múltiples rangos de tiempo en una sola ejecución.

### `actualizarReporteCompleto()`
Actualiza tanto datos de Meta Ads como de ventas.

### `obtenerComprasPorAdId()`
Vincula ventas con anuncios mediante AD ID para cálculo de ROAS.

## 🐛 Diagnóstico

### Diagnosticar Campañas

```
📊 Reporte de Ventas → 🔍 Diagnosticar Campañas
```

Muestra información de las primeras 5 campañas para verificar la conexión con Meta API.

### Diagnosticar AD ID Específico

Desde el editor de Apps Script, ejecuta:
```javascript
diagnosticarAdId()
```

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -m 'feat: Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📝 Convenciones de Commits

Este proyecto sigue [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` Nueva funcionalidad
- `fix:` Corrección de bugs
- `docs:` Cambios en documentación
- `style:` Formateo de código
- `refactor:` Refactorización de código
- `test:` Agregar o modificar tests
- `chore:` Tareas de mantenimiento

## ⚠️ Notas Importantes

- **Permisos**: En la primera ejecución, Google solicitará autorización para acceder a hojas de cálculo y hacer peticiones HTTP externas.
- **Límites de API**: Meta Ads tiene límites de rate limiting. El sistema maneja paginación automáticamente.
- **Tasas de Cambio**: Se actualizan cada 6 horas mediante caché. Fuente: [Exchange Rate API](https://exchangerate-api.com/)
- **IGV**: Calculado al 18% (configurable en el código si es necesario)

## 📧 Soporte

Para reportar problemas o solicitar funcionalidades, abre un [Issue](https://github.com/eflorianc/operacion_wsp_reportes/issues).

## 📜 Licencia

Este proyecto es privado. Todos los derechos reservados.

---

⚡ Desarrollado con Google Apps Script y ❤️
