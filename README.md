# 📦 Gestión de Stock

Aplicación web para que emprendedores y pequeños negocios controlen su stock, ventas, gastos y clientes desde un solo lugar. Funciona directamente en el navegador, sin necesidad de instalar nada ni de un servidor.

## ✨ Funcionalidades

- **Panel de control (Inicio):** métricas de productos, unidades, valor del stock, ventas del mes, ganancia estimada y alertas de stock bajo.
- **Stock:** alta, edición y baja de productos con semáforo de stock (verde / amarillo / rojo), búsqueda y filtros por marca y estado.
- **Entrada de mercadería:** panel lateral para cargar stock de productos existentes o crear productos nuevos sobre la marcha.
- **Ventas:** registro de ventas con descuento automático de stock y asignación opcional a un cliente.
- **Gastos:** registro de gastos por categoría con filtros.
- **Mini CRM:** fichas de clientes con contacto, notas, total comprado y última compra.
- **Configuración:** nombre del negocio, logo, color de fondo, color de la barra lateral, moneda y umbrales de alerta de stock.
- **Respaldo en Excel:** exportación de todos los datos a un archivo `.xlsx` con varias hojas (Resumen, Productos, Ventas, Gastos, Clientes, Marcas y Configuración).
- **Importar respaldo:** restauración desde un archivo Excel o desde un respaldo JSON.

## 🚀 Cómo usarla

1. Descargá o cloná el proyecto.
2. Abrí el archivo `index.html` en cualquier navegador moderno (Chrome, Edge, Firefox).
3. Listo. Empezá a cargar tus productos.

> También se puede publicar gratis con **GitHub Pages** y usarla desde una URL.

## 💾 Dónde se guardan los datos

Los datos se almacenan localmente en el navegador (`localStorage`). No se envían a ningún servidor.

- Para **respaldar** tu información usá *Exportar a Excel*.
- Para **recuperarla** en otro dispositivo usá *Importar respaldo*.

> ⚠️ Si borrás los datos del navegador o del sitio, se pierde la información que no hayas exportado. Hacé respaldos seguido.

## 🗂️ Estructura del proyecto

```
gestion/
├── index.html            # Estructura de la aplicación
├── styles.css            # Estilos y diseño
├── app.js                # Lógica de la aplicación
├── libs/
│   └── xlsx.full.min.js  # Librería SheetJS para exportar/importar Excel
└── README.md
```

## 🛠️ Tecnologías

- HTML5, CSS3 y JavaScript vanilla (sin frameworks).
- [SheetJS](https://sheetjs.com/) para la exportación e importación de archivos Excel.
- `localStorage` para la persistencia de datos.

## 📄 Licencia

Uso libre para emprendedores.
