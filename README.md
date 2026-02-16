# 📊 Calculom-tricas - Sistema de Análisis de Reservas

Sistema web para cargar y analizar archivos Excel con datos de reservas, proporcionando métricas detalladas por instancia, tasas de conversión y rankings de rendimiento.

## 🚀 Características

- **Carga de Archivos Excel**: Interfaz drag-and-drop para cargar archivos .xlsx y .xls
- **📅 Filtro por Fecha**: Filtra reservas por rango de fechas de creación (Columna Z)
- **Análisis por Instancia**: Desglose completo de reservas por estado (Confirmada, Pendiente, Fallida, Procesando)
- **Métricas de Conversión**: Cálculo de tasas de conversión cotizaciones/reservas
- **Top 5 Ranking**: Visualización de las instancias con más ventas
- **📄 Exportación a PDF**: Genera reportes completos en PDF con todos los datos, gráficos y estadísticas
- **Diseño Responsivo**: Funciona en desktop y móviles

## 📋 Formato del Archivo Excel

El archivo Excel debe contener las siguientes columnas:

| Columna | Campo |
|---------|-------|
| A | ID Reserva |
| B | **Instancia** (obligatorio) |
| C | Email Instancia |
| D | Zona Horaria |
| E | Nombre Huésped |
| F | Email Huésped |
| G | Teléfono Huésped |
| H | Fecha Check-in |
| I | Fecha Check-out |
| J | Noches |
| K | Habitaciones |
| L | Total Huéspedes |
| M | Detalle Habitaciones |
| N | Precio Total |
| O | Moneda |
| P | Monto Pagado |
| Q | Monto Pendiente |
| R | Depósito |
| S | Método de Pago |
| T | Estado de Pago |
| U | **Estado de Reserva** (obligatorio: Confirmada/Pendiente/Fallida/Procesando) |
| V | Canal |
| W | ID Canal Reserva |
| X | Source |
| Y | Creado Por |
| Z | **Fecha Creación** (usado para filtros de fecha) |
| AA | Fecha Actualización |

## 🛠️ Instalación

### Opción 1: Uso Local

1. Clona el repositorio:
```bash
git clone https://github.com/alesoander/calculom-tricas.git
cd calculom-tricas
```

2. Abre `index.html` directamente en tu navegador

### Opción 2: Servidor Web

```bash
# Usando Python
python -m http.server 8000

# Usando Node.js
npx serve

# Usando PHP
php -S localhost:8000
```

Luego visita `http://localhost:8000` en tu navegador.

### Opción 3: GitHub Pages

1. Ve a Settings > Pages en tu repositorio
2. Selecciona la rama `main` como fuente
3. Tu sitio estará disponible en `https://alesoander.github.io/calculom-tricas/`

## 📖 Uso

1. **Cargar Archivo**: Haz clic en "Seleccionar Archivo" o arrastra tu Excel a la zona de carga
2. **📅 Filtrar por Fecha (Opcional)**: Usa el filtro de rango de fechas para analizar períodos específicos
   - Selecciona fecha de inicio ("Desde") y fecha final ("Hasta")
   - Haz clic en "Aplicar Filtro" para ver solo las reservas en ese rango
   - Usa "Limpiar Filtro" para restaurar todos los datos
3. **Ver Resumen**: Revisa las estadísticas generales de todas las instancias (o filtradas)
4. **Top 5**: Identifica las instancias con mejor rendimiento
5. **Ingresar Cotizaciones**: Para cada instancia, ingresa el número de cotizaciones
6. **Ver Conversiones**: El sistema calculará automáticamente las tasas de conversión
7. **📄 Exportar PDF**: Haz clic en el botón "Exportar PDF" para generar un reporte completo

### 📅 Filtro por Fecha

El sistema incluye un filtro de rango de fechas que permite:

- **Filtrar por Fecha de Creación**: Analiza reservas creadas en un período específico (Columna Z)
- **Formato Flexible**: Soporta fechas en formato de texto y números de serie de Excel
- **Actualización en Tiempo Real**: Todas las estadísticas, gráficos y métricas se actualizan automáticamente
- **Integración con PDF**: Los reportes PDF incluyen información del filtro aplicado
- **Validación**: El sistema valida que existan resultados antes de aplicar el filtro

### 📥 Exportación de PDF

El sistema permite generar reportes PDF profesionales que incluyen:

- **Cabecera**: Título del reporte, fecha de generación, nombre del archivo cargado y rango de filtro (si aplica)
- **Resumen General**: Todas las estadísticas globales (total reservas, confirmadas, pendientes, fallidas, procesando, instancias)
- **Top 5 Instancias**: Gráfico visual y tabla con las 5 instancias con más ventas
- **Tasas de Conversión Globales**: Total de cotizaciones y porcentajes de conversión
- **Detalles por Instancia**: Información completa de cada instancia:
  - Total de reservas y desglose por estado
  - Cantidad de cotizaciones
  - Porcentajes de conversión (cotizaciones/total y cotizaciones/confirmadas)
- **Pie de Página**: Números de página, timestamp de generación y marca del sistema

El PDF se descarga automáticamente con un nombre único basado en la fecha y hora: `reporte-reservas-YYYY-MM-DD-HHMMSS.pdf`


## 📊 Métricas Calculadas

### Por Instancia:
- Total de reservas
- Reservas por estado (Confirmada, Pendiente, Fallida, Procesando)
- Cotizaciones / Total Reservas (%)
- Cotizaciones / Reservas Confirmadas (%)

### Globales:
- Total de reservas en todas las instancias
- Total de confirmadas, pendientes, fallidas y procesando
- Total Cotizaciones / Total Confirmadas (%)
- Top 5 instancias con más ventas

## 🎨 Tecnologías

- **HTML5**: Estructura semántica
- **CSS3**: Diseño moderno y responsivo
- **JavaScript (ES6+)**: Lógica de procesamiento
- **SheetJS (xlsx)**: Procesamiento de archivos Excel
- **Chart.js**: Visualización de datos
- **jsPDF**: Generación de documentos PDF
- **html2canvas**: Captura de gráficos para PDF

## 🔧 Dependencias

Las siguientes librerías se cargan desde CDN (no requieren instalación):

- SheetJS (xlsx) v0.18.5
- Chart.js v4.x
- jsPDF v2.5.1
- html2canvas v1.4.1

## 📱 Compatibilidad

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Safari 14+
- ✅ Edge 90+
- ✅ Dispositivos móviles (iOS/Android)

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu función (`git checkout -b feature/NuevaFuncion`)
3. Commit tus cambios (`git commit -m 'Agregar nueva función'`)
4. Push a la rama (`git push origin feature/NuevaFuncion`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es de código abierto y está disponible bajo la Licencia MIT.

## 👤 Autor

**alesoander**

## 🐛 Reportar Problemas

Si encuentras algún bug o tienes sugerencias, por favor abre un [Issue](https://github.com/alesoander/calculom-tricas/issues).

## 📝 Notas

- El procesamiento del archivo se realiza completamente en el navegador (client-side)
- No se envían datos a ningún servidor
- Los datos se mantienen en memoria solo durante la sesión
- Compatible con archivos Excel de cualquier tamaño (dentro de los límites del navegador)