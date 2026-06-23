# Comprobantes de Pago - Perú

Aplicación web para la generación automatizada de **Boletas de Pago** (comprobantes de pago) de **Clarke, Modet & Co. Peru S.A.C.**, desarrollada por **Solutions & Payroll**.

## ¿Qué hace?

Genera un libro Excel multi-hoja con boletas de pago por empleado a partir de tres archivos de entrada:

1. **Detallado de Nómina** — conceptos salariales por empleado (devengados y deducciones)
2. **Datos ESSALUD** — aportes al seguro social de salud
3. **Base de Empleados** — datos complementarios (AFP, etc.)

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
| **JSZip** | Compresión ZIP (dependencia interna de ExcelJS) |

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

1. Carga los tres archivos Excel requeridos arrastrándolos a cada zona de carga o haciendo clic:
   - **Detallado de Nómina** (.xlsx)
   - **Datos ESSALUD** (.xlsx)
   - **Base de Empleados** (.xlsx)
2. Haz clic en **"Generar Comprobantes"**.
3. La aplicación procesa los datos y descarga automáticamente el archivo `Comprobantes de Pago.xlsx`.
4. El Excel generado contiene una hoja por cada empleado con su boleta de pago formateada.

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

1. **Carga de archivos**: Tres zonas drag & drop para los archivos Excel de entrada.
2. **Parseo**: SheetJS lee los archivos en memoria y los convierte a arrays 2D.
3. **Procesamiento**: Agrupa los conceptos por empleado, cruza los datos entre las tres fuentes.
4. **Generación**: ExcelJS construye el libro de salida hoja por hoja, con celdas fusionadas, imágenes incrustadas, bordes, formatos de número y estilos profesionales.
5. **Descarga**: El libro se convierte a `ArrayBuffer`, se crea un Blob y se descarga automáticamente.

## Licencia

© 2026 Solutions & Payroll. Uso interno.
