import { useState, useRef, useCallback } from 'react'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs/dist/exceljs.min.js'
import './App.css'

function extractMonthYear(periodStr) {
  if (!periodStr || typeof periodStr !== 'string') return { month: '', year: '' }
  return { year: periodStr.substring(0, 4), month: periodStr.substring(4, 6) }
}

function getLastDayOfPeriod(periodStr) {
  if (!periodStr || typeof periodStr !== 'string') return null
  const year = parseInt(periodStr.substring(0, 4), 10)
  const month = parseInt(periodStr.substring(4, 6), 10)
  const lastDay = new Date(year, month, 0).getDate()
  return Math.floor(new Date(year, month - 1, lastDay).getTime() / 86400000) + 25569
}

function excelSerialToDMY(serial) {
  if (!serial || typeof serial !== 'number') return ''
  const dt = new Date((serial - 25569) * 86400 * 1000)
  const day = String(dt.getUTCDate()).padStart(2, '0')
  const month = String(dt.getUTCMonth() + 1).padStart(2, '0')
  const year = dt.getUTCFullYear()
  return `${day}/${month}/${year}`
}

const FONT_TITLE = { name: 'Arial', family: 2, bold: true, size: 12, color: { theme: 1 } }
const FONT_HEADER = { name: 'Arial', family: 2, bold: true, size: 12, color: { indexed: 8 } }
const FONT_DATA_SM = { name: 'Arial', family: 2, size: 10, color: { theme: 1 } }
const FONT_DATA = { name: 'Arial', family: 2, size: 11, color: { theme: 1 } }
const FONT_DATA_LG = { name: 'Arial', family: 2, size: 12, color: { theme: 1 } }
const FILL_WHITE = { type: 'pattern', pattern: 'solid', fgColor: { theme: 0 }, bgColor: { indexed: 64 } }
const FILL_GRAY = { type: 'pattern', pattern: 'solid', fgColor: { theme: 0, tint: -0.1499984740745262 }, bgColor: { indexed: 64 } }
const ALIGN_CENTER = { horizontal: 'center', vertical: 'middle' }
const ALIGN_LEFT = { horizontal: 'left', vertical: 'middle' }
const BORDER = { style: 'thin', color: { indexed: 64 } }
const BORDER_BOX = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }
const BORDER_TOP = { top: BORDER }

function colName(n) {
  let s = ''
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) }
  return s
}

function merge(ws, r1, c1, r2, c2) {
  ws.mergeCells(`${colName(c1)}${r1}:${colName(c2)}${r2}`)
}

function sc(ws, r, c, value, font, fill, align, border) {
  const cell = ws.getCell(r, c)
  cell.value = value
  cell.font = font || FONT_DATA
  cell.fill = fill || FILL_WHITE
  if (align) cell.alignment = align
  if (border !== null) cell.border = border !== undefined ? border : BORDER_BOX
}

function scNum(ws, r, c, value, font, fill, align) {
  const cell = ws.getCell(r, c)
  cell.value = value
  cell.numFmt = '#,##0.00'
  cell.font = font || FONT_DATA
  cell.fill = fill || FILL_WHITE
  if (align) cell.alignment = align
  cell.border = BORDER_BOX
}

function generarExcel(detalladoFile, essaludFile, baseEmpleadosFile) {
  return new Promise((resolve, reject) => {
    const rn = new FileReader()
    const re = new FileReader()
    const rb = new FileReader()
    let nd = null, ed = null, bd = null, err = null

    const check = () => {
      if (err) return
      if (nd && ed && bd) buildWorkbook(nd, ed, bd).then(resolve).catch(reject)
    }

    rn.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' })
        nd = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' })
        check()
      } catch (e) { err = e; reject(e) }
    }
    rn.onerror = () => reject(new Error('Error al leer Detallado Nómina'))
    rn.readAsArrayBuffer(detalladoFile)

    re.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' })
        ed = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' })
        check()
      } catch (e) { err = e; reject(e) }
    }
    re.onerror = () => reject(new Error('Error al leer Datos ESSALUD'))
    re.readAsArrayBuffer(essaludFile)

    rb.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' })
        bd = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' })
        check()
      } catch (e) { err = e; reject(e) }
    }
    rb.onerror = () => reject(new Error('Error al leer Base Empleados'))
    rb.readAsArrayBuffer(baseEmpleadosFile)
  })
}

async function buildWorkbook(nominaRows, essaludRows, baseRows) {
  const wb = new ExcelJS.Workbook()
  const dataRows = nominaRows.slice(3)
  const empMap = {}
  const empOrder = []

  for (const row of dataRows) {
    const emp = String(row[1] || '').trim()
    if (!emp) continue
    if (!empMap[emp]) {
      empMap[emp] = {
        empleado: emp, docIdent: row[3], nombreCompleto: row[2],
        fAntiguedad: typeof row[4] === 'number' ? row[4] : null,
        salario: row[5], periodoMes: row[8],
        oficio: row[18], descOficio: row[19], area: row[20], descArea: row[21],
        cussp: String(row[25] || '').trim(), conceptos: []
      }
      empOrder.push(emp)
    }
    empMap[emp].conceptos.push({
      concepto: String(row[9] || '').trim(),
      descConcepto: String(row[10] || '').trim(),
      horas: typeof row[11] === 'number' ? row[11] : 0,
      valorDevengado: typeof row[12] === 'number' ? row[12] : 0,
      valorDeduccion: typeof row[13] === 'number' ? row[13] : 0
    })
  }

  const essMap = {}
  for (const row of essaludRows.slice(1)) {
    const c = String(row[1] || '').trim()
    if (!c) continue
    essMap[c] = { aporteEssalud: row[6], creditoEpsTrab: row[7], totalAportes: row[8] }
  }

  // AFP map from Base Empleados
  const baseHeaders = baseRows[0]
  const afpCol = baseHeaders.indexOf('AFP')
  const empCol = baseHeaders.indexOf('EMPLEADO')
  const afpMap = {}
  for (const row of baseRows.slice(1)) {
    const emp = String(row[empCol] || '').trim()
    if (emp) afpMap[emp] = String(row[afpCol] || '').trim()
  }

  for (let i = 0; i < empOrder.length; i++) {
    const ed = empMap[empOrder[i]]
    const es = essMap[empOrder[i]] || {}
    const afp = afpMap[empOrder[i]] || ''
    const { month, year } = extractMonthYear(ed.periodoMes)
    const idx = i + 1
    const name = idx === 1 ? 'PS1' : `PS1 (${idx})`
    buildSheet(wb, name, ed, es, month, year, afp)
  }

  return wb
}

function buildSheet(wb, name, ed, es, month, year, afp) {
  const ws = wb.addWorksheet(name, {
    properties: { tabColor: { argb: 'FF92D050' } },
    pageSetup: { fitToPage: true },
    views: [{ showGridLines: false }]
  })

  const cw = [15.66, 29.33, 15.66, 15.66, 15.66, 30.44, 15.66, 16.33, 15.66, 15.66, 15.66, 15.66]
  for (let i = 0; i < cw.length; i++) ws.getColumn(i + 1).width = cw[i]

  ws.getRow(1).height = 54
  for (let r = 2; r <= 26; r++) ws.getRow(r).height = 24.6
  ws.getRow(10).height = 28.5
  for (let r = 27; r <= 36; r++) ws.getRow(r).height = 15.6

  // Row 1: Title A-J + Image area K-L
  sc(ws, 1, 1, 'BOLETA DE PAGO', FONT_TITLE, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 1, 1, 1, 10)
  sc(ws, 1, 11, '', FONT_TITLE, FILL_WHITE)
  merge(ws, 1, 11, 1, 12)

  // Row 2
  sc(ws, 2, 1, 'EMPRESA', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  merge(ws, 2, 1, 3, 1)
  sc(ws, 2, 2, 'CLARKE, MODET & CO. PERU S.A.C.', FONT_DATA_LG, FILL_WHITE, ALIGN_LEFT)
  merge(ws, 2, 2, 3, 4)
  sc(ws, 2, 5, 'R.U.C.', FONT_HEADER, FILL_GRAY)
  sc(ws, 2, 6, '20256396000', FONT_DATA_LG, FILL_WHITE)
  sc(ws, 2, 7, 'MES', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 2, 8, month, FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 2, 9, 'AÑO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 2, 10, year, FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 2, 11, 'F. INGRESO', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  sc(ws, 2, 12, ed.fAntiguedad ? excelSerialToDMY(ed.fAntiguedad) : '', FONT_DATA_LG, FILL_WHITE, ALIGN_LEFT)

  // Row 3
  sc(ws, 3, 5, 'PERIODO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 3, 6, month, FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 3, 7, 'DEL', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 3, 8, `01/${month}/${year}`, FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 3, 9, 'AL', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  const ls = getLastDayOfPeriod(ed.periodoMes)
  sc(ws, 3, 10, ls ? excelSerialToDMY(ls) : '', FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 3, 11, 'F. RETIRO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 3, 12, '', FONT_DATA_LG, FILL_WHITE, ALIGN_CENTER)

  // Row 5-6: Employee info
  sc(ws, 5, 1, 'CÓDIGO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 5, 2, 'NOMBRES Y APELLIDOS', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 5, 2, 5, 5)
  sc(ws, 5, 6, 'CARGO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 5, 6, 5, 9)
  sc(ws, 5, 10, 'AREA', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 5, 10, 5, 12)

  sc(ws, 6, 1, ed.empleado, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 6, 2, ed.nombreCompleto, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 6, 2, 6, 5)
  sc(ws, 6, 6, ed.descOficio, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 6, 6, 6, 9)
  sc(ws, 6, 10, ed.descArea, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 6, 10, 6, 12)

  // Row 7-8: ID data
  sc(ws, 7, 1, 'D.N.I.', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 7, 2, 'CUSPP AFP', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 7, 3, '# AG ESSALUD', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 7, 3, 7, 4)
  sc(ws, 7, 5, 'SALARIO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 7, 5, 7, 6)
  sc(ws, 7, 7, 'MONEDA', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 7, 7, 7, 8)
  sc(ws, 7, 9, 'CODIGO EMPRESA', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 7, 9, 7, 10)
  sc(ws, 7, 11, 'A.F.P.', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 7, 11, 7, 12)

  sc(ws, 8, 1, ed.docIdent, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 8, 2, ed.cussp, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  sc(ws, 8, 3, ed.cussp, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 8, 3, 8, 4)
  sc(ws, 8, 5, ed.salario, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 8, 5, 8, 6)
  sc(ws, 8, 7, 'SOL', FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 8, 7, 8, 8)
  sc(ws, 8, 9, ed.docIdent, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 8, 9, 8, 10)
  sc(ws, 8, 11, afp, FONT_DATA, FILL_WHITE, ALIGN_CENTER)
  merge(ws, 8, 11, 8, 12)

  // Row 10: Section headers
  sc(ws, 10, 1, 'REMUNERACIONES', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 10, 1, 10, 4)
  sc(ws, 10, 5, 'DESCUENTOS TRABAJADOR', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 10, 5, 10, 8)
  sc(ws, 10, 9, 'APORTACIONES EMPLEADOR', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 10, 9, 10, 12)

  // Row 11: Concept headers
  sc(ws, 11, 1, 'CONCEPTO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 11, 1, 11, 2)
  sc(ws, 11, 3, 'HORAS', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 11, 4, 'VALOR', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 11, 5, 'CONCEPTO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 11, 5, 11, 6)
  sc(ws, 11, 7, 'HORAS', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 11, 8, 'VALOR', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 11, 9, 'CONCEPTO', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  merge(ws, 11, 9, 11, 10)
  sc(ws, 11, 11, 'HORAS', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)
  sc(ws, 11, 12, 'VALOR', FONT_HEADER, FILL_GRAY, ALIGN_CENTER)

  // Rows 12-14: Concepts
  const dev = ed.conceptos.filter(c => c.valorDevengado > 0)
  const ded = ed.conceptos.filter(c => c.valorDeduccion > 0)

  for (let i = 0; i < 3; i++) {
    const r = 12 + i
    if (i < dev.length) {
      const c = dev[i]
      sc(ws, r, 1, `${c.concepto} ${c.descConcepto}`, FONT_DATA_SM, FILL_WHITE)
      sc(ws, r, 3, c.horas, FONT_DATA_SM, FILL_WHITE)
      sc(ws, r, 4, c.valorDevengado, FONT_DATA_SM, FILL_WHITE)
    }
    if (i < ded.length) {
      const c = ded[i]
      sc(ws, r, 5, `${c.concepto} ${c.descConcepto}`, FONT_DATA_SM, FILL_WHITE)
      sc(ws, r, 7, c.horas, FONT_DATA_SM, FILL_WHITE)
      sc(ws, r, 8, c.valorDeduccion, FONT_DATA_SM, FILL_WHITE)
    }
  }

  const ae = typeof es.aporteEssalud === 'number' ? es.aporteEssalud : 0
  const ce = typeof es.creditoEpsTrab === 'number' ? es.creditoEpsTrab : 0

  sc(ws, 12, 9, 'APORTE A ESSALUD', FONT_DATA_SM, FILL_WHITE)
  sc(ws, 12, 11, 0, FONT_DATA_SM, FILL_WHITE)
  sc(ws, 12, 12, r2(ae), FONT_DATA_SM, FILL_WHITE)

  sc(ws, 13, 9, 'CREDITO EPS', FONT_DATA_SM, FILL_WHITE)
  sc(ws, 13, 11, 0, FONT_DATA_SM, FILL_WHITE)
  sc(ws, 13, 12, r2(ce), FONT_DATA_SM, FILL_WHITE)

  // Row 25: Total aporte empleador + top border A-H
  const tae = typeof es.totalAportes === 'number' ? es.totalAportes : (ae + ce)
  for (let c = 1; c <= 8; c++) {
    const cell = ws.getCell(25, c)
    cell.value = cell.value || ''
    cell.font = cell.font || FONT_DATA_LG
    cell.fill = cell.fill || FILL_WHITE
    cell.border = { top: BORDER }
  }
  sc(ws, 25, 9, 'TOTAL APORTE EMPLEADOR', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  merge(ws, 25, 9, 25, 11)
  scNum(ws, 25, 12, r2(tae), FONT_DATA_LG, FILL_WHITE)

  // Row 26: Totals
  const tr = ed.conceptos.reduce((s, c) => s + c.valorDevengado, 0)
  const td = ed.conceptos.reduce((s, c) => s + c.valorDeduccion, 0)
  sc(ws, 26, 1, 'REMUNERACION TOTAL', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  merge(ws, 26, 1, 26, 3)
  scNum(ws, 26, 4, tr, FONT_DATA_LG, FILL_WHITE)
  sc(ws, 26, 5, 'DESCUENTOS TOTALES', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  merge(ws, 26, 5, 26, 7)
  scNum(ws, 26, 8, r2(td), FONT_DATA_LG, FILL_WHITE)
  sc(ws, 26, 9, 'REMUNERACION NETA', FONT_HEADER, FILL_GRAY, ALIGN_LEFT)
  merge(ws, 26, 9, 26, 11)
  scNum(ws, 26, 12, r2(tr - td), FONT_DATA_LG, FILL_WHITE)

  // Rows 12-24: merge A-B, E-F, I-J and apply lateral borders
  for (let r = 12; r <= 24; r++) {
    merge(ws, r, 1, r, 2)
    merge(ws, r, 5, r, 6)
    merge(ws, r, 9, r, 10)

    for (let c = 1; c <= 12; c++) {
      const cell = ws.getCell(r, c)
      if (cell.value === null || cell.value === undefined) {
        cell.value = ''
        cell.font = FONT_DATA_SM
        cell.fill = FILL_WHITE
      }
      if (c === 1 || c === 5 || c === 9) cell.border = { left: BORDER }
      else if (c === 3 || c === 4 || c === 7 || c === 8 || c === 11 || c === 12) cell.border = { left: BORDER, right: BORDER }
    }
  }

  // Row 36: Signatures
  sc(ws, 36, 6, 'REPRESENTANTE LEGAL', FONT_DATA_LG, FILL_WHITE, ALIGN_LEFT, BORDER_TOP)
  merge(ws, 36, 6, 36, 8)
  sc(ws, 36, 10, 'RECIBI CONFORME TRABAJADOR', FONT_DATA_LG, FILL_WHITE, ALIGN_LEFT, BORDER_TOP)
  merge(ws, 36, 10, 36, 12)

  // Rows 27-36: right border on column L + row 36 bottom border
  for (let r = 27; r <= 36; r++) {
    const cell = ws.getCell(r, 12)
    cell.border = { ...cell.border, right: BORDER }
  }
  for (let c = 1; c <= 12; c++) {
    const cell = ws.getCell(36, c)
    cell.border = { ...cell.border, bottom: BORDER }
  }

  // L4 and L9: right border
  ws.getCell(4, 12).border = { right: BORDER }
  ws.getCell(9, 12).border = { right: BORDER }
}

function r2(n) {
  return Math.round(n * 100) / 100
}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1048576).toFixed(1)} MB`
}

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [detalladoFile, setDetalladoFile] = useState(null)
  const [essaludFile, setEssaludFile] = useState(null)
  const [baseFile, setBaseFile] = useState(null)
  const [isDraggingDetallado, setIsDraggingDetallado] = useState(false)
  const [isDraggingEssalud, setIsDraggingEssalud] = useState(false)
  const [isDraggingBase, setIsDraggingBase] = useState(false)
  const [isGenerating, setIsGenerating] = useState(false)
  const [modal, setModal] = useState(null)

  const detalladoInputRef = useRef(null)
  const essaludInputRef = useRef(null)
  const baseInputRef = useRef(null)

  const handleDragOver = useCallback((e, setDragging) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(true)
  }, [])

  const handleDragLeave = useCallback((e, setDragging) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(false)
  }, [])

  const handleDrop = useCallback((e, setDragging, setFile) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(false)
    const file = e.dataTransfer.files[0]
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setFile(file)
    } else {
      setModal({ type: 'error', title: 'Formato inválido', message: 'Por favor, sube un archivo Excel (.xlsx o .xls)' })
    }
  }, [])

  const handleFileSelect = useCallback((e, setFile) => {
    const file = e.target.files[0]
    if (file) setFile(file)
  }, [])

  const handleRemoveFile = useCallback((setFile, inputRef) => {
    setFile(null)
    if (inputRef.current) inputRef.current.value = ''
  }, [])

  const handleGenerate = useCallback(async () => {
    setIsGenerating(true)
    try {
      const wb = await generarExcel(detalladoFile, essaludFile, baseFile)

      const imgRes = await fetch('/Icono Clarke.png')
      const imgBlob = await imgRes.blob()
      const base64 = await new Promise((resolve) => {
        const reader = new FileReader()
        reader.onloadend = () => resolve(reader.result.split(',')[1])
        reader.readAsDataURL(imgBlob)
      })
      const imageId = wb.addImage({ base64, extension: 'png' })
      wb.worksheets.forEach(ws => {
        ws.addImage(imageId, {
          tl: { col: 10, colOff: 350000, row: 0 },
          ext: { width: 167, height: 68 }
        })
      })

      const repRes = await fetch('/Representante legal.png')
      const repBlob = await repRes.blob()
      const repBase64 = await new Promise((resolve) => {
        const reader = new FileReader()
        reader.onloadend = () => resolve(reader.result.split(',')[1])
        reader.readAsDataURL(repBlob)
      })
      const repImageId = wb.addImage({ base64: repBase64, extension: 'png' })
      wb.worksheets.forEach(ws => {
        ws.addImage(repImageId, {
          tl: { col: 5, row: 28 },
          ext: { width: 324, height: 130 }
        })
      })

      const buffer = await wb.xlsx.writeBuffer()
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'Comprobantes de Pago.xlsx'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
      setModal({ type: 'success', title: 'Generado correctamente', message: 'El archivo de comprobantes de pago se ha descargado.' })
    } catch (err) {
      setModal({ type: 'error', title: 'Error', message: err.message || 'Ocurrió un error al generar el archivo.' })
    } finally {
      setIsGenerating(false)
    }
  }, [detalladoFile, essaludFile, baseFile])

  return (
    <div className="app">
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img 
                  src="/Logo syp.png" 
                  alt="Solutions & Payroll Logo" 
                  width="60" 
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Comprobantes de Pago</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      <main className="main-content">
        <div className="container">
          <div className="help-section">
            <button 
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg 
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" 
                height="20" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Sube el Detallado de Nómina</strong>
                    <p>Adjunta el archivo Excel con el detallado de nómina del mes</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Sube los Datos ESSALUD</strong>
                    <p>Adjunta el archivo Excel con los aportes de ESSALUD</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Genera los Comprobantes</strong>
                    <p>Haz clic en Generar para descargar el archivo de comprobantes de pago</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <h2>Generar Comprobantes de Pago</h2>
              <p className="description">
                Sube los archivos de nómina y ESSALUD para generar las boletas de pago
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Detallado Nómina
                  </label>
                  <input ref={detalladoInputRef} type="file" accept=".xlsx,.xls" className="file-input" id="detallado-file" onChange={(e) => handleFileSelect(e, setDetalladoFile)} />
                  <div
                    className={`drop-zone ${isDraggingDetallado ? 'drag-active' : ''} ${detalladoFile ? 'has-file' : ''}`}
                    onClick={() => detalladoInputRef.current?.click()}
                    onDragOver={(e) => handleDragOver(e, setIsDraggingDetallado)}
                    onDragLeave={(e) => handleDragLeave(e, setIsDraggingDetallado)}
                    onDrop={(e) => handleDrop(e, setIsDraggingDetallado, setDetalladoFile)}
                  >
                    {detalladoFile ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="16" y1="13" x2="8" y2="13"/>
                            <line x1="16" y1="17" x2="8" y2="17"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <div className="file-name">{detalladoFile.name}</div>
                          <div className="file-size">{formatFileSize(detalladoFile.size)}</div>
                        </div>
                        <button type="button" className="btn-remove" onClick={(e) => { e.stopPropagation(); handleRemoveFile(setDetalladoFile, detalladoInputRef) }} title="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra aquí el archivo</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionar</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    )}
                  </div>
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Datos ESSALUD
                  </label>
                  <input ref={essaludInputRef} type="file" accept=".xlsx,.xls" className="file-input" id="essalud-file" onChange={(e) => handleFileSelect(e, setEssaludFile)} />
                  <div
                    className={`drop-zone ${isDraggingEssalud ? 'drag-active' : ''} ${essaludFile ? 'has-file' : ''}`}
                    onClick={() => essaludInputRef.current?.click()}
                    onDragOver={(e) => handleDragOver(e, setIsDraggingEssalud)}
                    onDragLeave={(e) => handleDragLeave(e, setIsDraggingEssalud)}
                    onDrop={(e) => handleDrop(e, setIsDraggingEssalud, setEssaludFile)}
                  >
                    {essaludFile ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="16" y1="13" x2="8" y2="13"/>
                            <line x1="16" y1="17" x2="8" y2="17"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <div className="file-name">{essaludFile.name}</div>
                          <div className="file-size">{formatFileSize(essaludFile.size)}</div>
                        </div>
                        <button type="button" className="btn-remove" onClick={(e) => { e.stopPropagation(); handleRemoveFile(setEssaludFile, essaludInputRef) }} title="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra aquí el archivo</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionar</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    )}
                  </div>
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/>
                      <path d="M9 3v18M15 3v18M3 9h18M3 15h18"/>
                    </svg>
                    Base Empleados
                  </label>
                  <input ref={baseInputRef} type="file" accept=".xlsx,.xls" className="file-input" id="base-file" onChange={(e) => handleFileSelect(e, setBaseFile)} />
                  <div
                    className={`drop-zone ${isDraggingBase ? 'drag-active' : ''} ${baseFile ? 'has-file' : ''}`}
                    onClick={() => baseInputRef.current?.click()}
                    onDragOver={(e) => handleDragOver(e, setIsDraggingBase)}
                    onDragLeave={(e) => handleDragLeave(e, setIsDraggingBase)}
                    onDrop={(e) => handleDrop(e, setIsDraggingBase, setBaseFile)}
                  >
                    {baseFile ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="16" y1="13" x2="8" y2="13"/>
                            <line x1="16" y1="17" x2="8" y2="17"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <div className="file-name">{baseFile.name}</div>
                          <div className="file-size">{formatFileSize(baseFile.size)}</div>
                        </div>
                        <button type="button" className="btn-remove" onClick={(e) => { e.stopPropagation(); handleRemoveFile(setBaseFile, baseInputRef) }} title="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra aquí el archivo</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionar</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    )}
                  </div>
                </div>

                <button className="btn-primary" onClick={handleGenerate}                   disabled={!detalladoFile || !essaludFile || !baseFile || isGenerating}>
                  {isGenerating ? (
                    <>
                      <svg className="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                      </svg>
                      Generando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Generar Comprobantes
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      </main>

      {modal && (
        <div className="modal-overlay" onClick={() => setModal(null)}>
          <div className={`modal-content ${modal.type}`} onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon">
              {modal.type === 'success' ? (
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                  <polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
              ) : (
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="15" y1="9" x2="9" y2="15"/>
                  <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
              )}
            </div>
            <h3 className="modal-title">{modal.title}</h3>
            <p className="modal-message">{modal.message}</p>
            <button className="modal-button" onClick={() => setModal(null)}>
              {modal.type === 'success' ? 'Continuar' : 'Cerrar'}
            </button>
            {modal.type === 'success' && (
              <div className="modal-progress">
                <div className="modal-progress-bar"/>
              </div>
            )}
          </div>
        </div>
      )}

      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App
