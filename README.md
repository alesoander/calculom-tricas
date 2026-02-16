# 📊 Calculom-tricas - Sistema de Análisis de Reservas

Sistema web para cargar y analizar archivos Excel con datos de reservas, proporcionando métricas detalladas por instancia, tasas de conversión y rankings de rendimiento.

## 🚀 Características

- **Carga de Archivos Excel**: Interfaz drag-and-drop para cargar archivos .xlsx y .xls
- **Análisis por Instancia**: Desglose completo de reservas por estado (Confirmada, Pendiente, Fallida, Procesando)
- **Métricas de Conversión**: Cálculo de tasas de conversión cotizaciones/reservas
- **Top 5 Ranking**: Visualización de las instancias con más ventas
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
| Z | Fecha Creación |
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
2. **Ver Resumen**: Revisa las estadísticas generales de todas las instancias
3. **Top 5**: Identifica las instancias con mejor rendimiento
4. **Ingresar Cotizaciones**: Para cada instancia, ingresa el número de cotizaciones
5. **Ver Conversiones**: El sistema calculará automáticamente las tasas de conversión

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

## 🔧 Dependencias

Las siguientes librerías se cargan desde CDN (no requieren instalación):

- SheetJS (xlsx) v0.18.5
- Chart.js v4.x

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