# Comprobantes de Pago - Perú

Aplicación web para la generación automatizada de **Boletas de Pago** (comprobantes de pago) de **Clarke, Modet & Co. Peru S.A.C.**, desarrollada por **Solutions & Payroll**.

## ¿Qué hace?

### Generación de Excel

Genera un libro Excel multi-hoja con boletas de pago por empleado a partir de tres archivos de entrada:

1. **Detallado de Nómina** — conceptos salariales por empleado (devengados y deducciones)
2. **Datos ESSALUD** — aportes al seguro social de salud
3. **Base de Empleados** — datos complementarios (AFP, etc.)

### División de PDF

Permite dividir un PDF generado desde el Excel en archivos individuales por empleado. Cada PDF se nombra automáticamente con el formato `CLARKEPE_AAAAMM_Comprobantes_CÓDIGO.pdf` y se descarga en un archivo ZIP.

Cada hoja del Excel generado contiene:

- Datos del empleador (Clarke, Modet & Co. Peru S.A.C. — RUC 20256396000)
- Información del empleado (código, nombre, DNI, cargo, área, salario, fecha de ingreso)
- AFP y número CUSPP
- Remuneraciones (ingresos)
- Descuentos del trabajador
- Aportaciones del empleador (ESSALUD, crédito EPS)
- Totales: bruto, deducciones, neto a pagar
- Imágenes corporativas y líneas de firma

## Tecnologías

| Tecnología | Uso |
|---|---|
| **React 18** | UI y lógica de negocio |
| **Vite 5** | Build tool y servidor de desarrollo |
| **SheetJS (xlsx)** | Lectura de archivos Excel de entrada |
| **ExcelJS** | Generación del libro Excel de salida con formato profesional |
| **JSZip** | Empaquetado ZIP para descarga de PDFs individuales |
| **pdf-lib** | División de PDF en archivos por página |
| **pdfjs-dist** | Extracción de texto del PDF (mes y año) |

> Toda la lógica corre en el navegador. **No requiere backend.**

## Requisitos previos

- **Node.js** >= 18
- **npm** >= 9

## Instalación

```bash
npm install
```

## Ejecución

### Desarrollo

```bash
npm run dev
```

Abre http://localhost:5173 en el navegador.

### Build de producción

```bash
npm run build
```

El output se genera en la carpeta `dist/` y puede servirse con cualquier servidor web estático.

### Preview del build

```bash
npm run preview
```

## Cómo usar

### Generar comprobantes (Excel)

1. Carga los tres archivos Excel requeridos arrastrándolos a cada zona de carga o haciendo clic:
   - **Detallado de Nómina** (.xlsx)
   - **Datos ESSALUD** (.xlsx)
   - **Base de Empleados** (.xlsx)
2. Haz clic en **"Generar Comprobantes"**.
3. La aplicación procesa los datos y descarga automáticamente el archivo `Comprobantes de Pago.xlsx` con una hoja por empleado.

### Dividir en PDFs individuales

1. Abre el Excel generado en Excel y guárdalo como PDF (Archivo → Guardar como → PDF).
2. En la misma sesión de la aplicación, arrastra el PDF a la sección **"Dividir PDF en Comprobantes Individuales"**.
3. Haz clic en **"Dividir PDF"**.
4. La aplicación divide el PDF por páginas, nombra cada archivo como `CLARKEPE_AAAAMM_Comprobantes_CÓDIGO.pdf` y descarga un ZIP con todos los PDFs.

> **Importante:** El PDF debe dividirse en la misma sesión en que se generó el Excel, ya que la aplicación usa el orden de los empleados del Excel generado para nombrar los archivos.

## Estructura del proyecto

```
├── public/                          # Assets estáticos
│   ├── Logo syp.png                 # Logo Solutions & Payroll
│   ├── Icono Clarke.png             # Logo Clarke, Modet & Co.
│   ├── Representante legal.png      # Firma representante legal
│   └── Plantilla Comprobantes.xlsx  # Plantillas de referencia
├── src/
│   ├── main.jsx                     # Entry point React
│   ├── App.jsx                      # Componente principal (UI + lógica)
│   ├── App.css                      # Estilos de la aplicación
│   └── index.css                    # Estilos globales
├── index.html                       # Shell HTML
├── vite.config.js                   # Configuración Vite
└── package.json                     # Dependencias y scripts
```

## Funcionamiento

### Generación de Excel

1. **Carga de archivos**: Tres zonas drag & drop para los archivos Excel de entrada.
2. **Parseo**: SheetJS lee los archivos en memoria y los convierte a arrays 2D.
3. **Procesamiento**: Agrupa los conceptos por empleado, cruza los datos entre las tres fuentes.
4. **Generación**: ExcelJS construye el libro de salida hoja por hoja, con celdas fusionadas, imágenes incrustadas, bordes, formatos de número y estilos profesionales.
5. **Descarga**: El libro se convierte a `ArrayBuffer`, se crea un Blob y se descarga automáticamente.

### División de PDF

1. **Carga del PDF**: Se sube el PDF exportado desde el Excel.
2. **Extracción**: pdfjs-dist extrae el mes y año desde el texto del PDF.
3. **División**: pdf-lib divide el PDF por páginas, una por empleado.
4. **Nombrado**: Cada PDF se nombra usando el código del empleado (obtenido del Excel generado previamente en la misma sesión) y el período extraído.
5. **Empaquetado**: JSZip agrupa todos los PDFs en un archivo ZIP y se descarga.

## Licencia

© 2026 Solutions & Payroll. Uso interno.
