import * as XLSX from 'xlsx'
import type { HeaderBandCell, PropositionTable } from '@/lib/excel-propositions-types'

export type { PropositionTable, HeaderBandCell } from '@/lib/excel-propositions-types'
export { DEMO_CUSTOMER_INTEL_XLSX } from '@/lib/excel-propositions-types'

function normalizeSheetTitle(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, ' ')
}

/** Match "Proposition 1", "Preposition 2", etc. (common workbook typo). */
function sheetPrepositionIndex(name: string): number | null {
  const n = normalizeSheetTitle(name)
  const m = n.match(/^(preposition|proposition)\s*(\d)\s*$/)
  if (!m) return null
  const idx = Number(m[2]) - 1
  return idx >= 0 && idx < 3 ? idx : null
}

/** Prefer sheets named Proposition/Preposition 1–3; otherwise skip generic tabs and take next three. */
export function pickPropositionSheetNames(sheetNames: string[]): string[] {
  const picked: (string | undefined)[] = [undefined, undefined, undefined]

  for (const original of sheetNames) {
    const idx = sheetPrepositionIndex(original)
    if (idx !== null && !picked[idx]) picked[idx] = original
  }

  if (picked.every(Boolean)) return picked as string[]

  const usable = sheetNames.filter((s) => {
    const n = normalizeSheetTitle(s)
    return !['home', 'index', 'readme', 'cover', 'toc'].includes(n)
  })

  for (let i = 0; i < 3; i++) {
    if (!picked[i]) picked[i] = usable[i]
  }

  return picked.map((p, i) => p ?? `__missing_${i}`)
}

function cellToString(cell: unknown): string {
  if (cell === null || cell === undefined) return ''
  if (cell instanceof Date) return cell.toISOString().slice(0, 10)
  return String(cell).replace(/\r\n/g, '\n').trim()
}

/** First body row: serial in column A and multiple populated cells (skips title rows). */
function findFirstDataRowIndex(stringGrid: string[][]): number {
  for (let r = 0; r < stringGrid.length; r++) {
    const row = stringGrid[r] ?? []
    const first = String(row[0] ?? '').trim()
    if (!/^\d+$/.test(first)) continue
    const n = Number(first)
    if (n < 1 || n > 1_000_000) continue
    const nonEmpty = row.filter((c) => c !== '').length
    if (nonEmpty >= 3) return r
  }
  return -1
}

/** Title block in row 0, column A (often multi-line). */
function extractSectionIntro(stringGrid: string[][]): {
  sectionHeading: string
  sectionDescription: string
} {
  const raw = stringGrid[0]?.[0] ?? ''
  if (!raw.trim()) return { sectionHeading: '', sectionDescription: '' }
  const lines = raw.split(/\n/).map((l) => l.trim()).filter(Boolean)
  if (lines.length === 0) return { sectionHeading: '', sectionDescription: '' }
  if (lines.length === 1) return { sectionHeading: lines[0], sectionDescription: '' }
  return {
    sectionHeading: lines[0],
    sectionDescription: lines.slice(1).join(' '),
  }
}

/** Leaf header row: fill blanks upward within header band. */
function buildHeadersFromGrid(stringGrid: string[][], headerRowIdx: number, scanTop: number): string[] {
  const lengths = stringGrid.slice(scanTop, headerRowIdx + 1).map((row) => row?.length ?? 0)
  const width = lengths.length ? Math.max(...lengths) : 0
  const headerRow = stringGrid[headerRowIdx] ?? []

  return Array.from({ length: width }, (_, j) => {
    const direct = headerRow[j] ?? ''
    if (direct !== '') return direct
    for (let up = headerRowIdx - 1; up >= scanTop; up--) {
      const cell = stringGrid[up]?.[j]
      if (cell !== undefined && cell !== '') return cell
    }
    return `Column ${j + 1}`
  })
}

/** Rows directly above the first data row: stop at first blank row or title-only row. */
function findHeaderBandRowIndices(stringGrid: string[][], dataRowIdx: number): number[] {
  if (dataRowIdx <= 0) return []
  const indices: number[] = []
  for (let r = dataRowIdx - 1; r >= 0; r--) {
    const row = stringGrid[r] ?? []
    const nonEmpty = row.filter((c) => c !== '').length
    if (nonEmpty === 0) break
    const widest = row.reduce((m, c) => Math.max(m, String(c).length), 0)
    const onlyTitleLike = nonEmpty === 1 && widest > 160
    if (onlyTitleLike) break
    indices.push(r)
  }
  return indices.reverse()
}

function sheetColumnCount(sheet: XLSX.WorkSheet, stringGrid: string[][]): number {
  const ref = sheet['!ref']
  if (ref) {
    try {
      const range = XLSX.utils.decode_range(ref)
      return Math.max(range.e.c + 1, 1)
    } catch {
      /* fall through */
    }
  }
  let m = 0
  for (const row of stringGrid) {
    m = Math.max(m, row?.length ?? 0)
  }
  return Math.max(m, 1)
}

/** Build `<thead>` rows honoring Excel !merges over the header band. */
function buildMergedHeaderBand(
  sheet: XLSX.WorkSheet,
  stringGrid: string[][],
  headerRowIndices: number[],
  maxCol: number
): HeaderBandCell[][] {
  if (!headerRowIndices.length || maxCol <= 0) return []

  const merges = sheet['!merges'] ?? []
  const mergeOrigin = new Map<string, { colspan: number; rowspan: number }>()
  const skip = new Set<string>()

  for (const m of merges) {
    const { s, e } = m
    mergeOrigin.set(`${s.r},${s.c}`, {
      colspan: e.c - s.c + 1,
      rowspan: e.r - s.r + 1,
    })
    for (let r = s.r; r <= e.r; r++) {
      for (let c = s.c; c <= e.c; c++) {
        if (r === s.r && c === s.c) continue
        skip.add(`${r},${c}`)
      }
    }
  }

  const rows: HeaderBandCell[][] = []
  for (const r of headerRowIndices) {
    const rowCells: HeaderBandCell[] = []
    for (let c = 0; c < maxCol; c++) {
      if (skip.has(`${r},${c}`)) continue
      const mo = mergeOrigin.get(`${r},${c}`)
      rowCells.push({
        text: stringGrid[r]?.[c] ?? '',
        colspan: mo?.colspan ?? 1,
        rowspan: mo?.rowspan ?? 1,
        startCol: c,
      })
    }
    rows.push(rowCells)
  }
  return rows
}

function emptyTableBase(): Omit<PropositionTable, 'title' | 'sheetName'> {
  return {
    sectionHeading: '',
    sectionDescription: '',
    headers: [],
    headerRows: [],
    rows: [],
  }
}

/** Parse workbook bytes (Node Buffer or browser Uint8Array). */
export function workbookToPropositionTables(input: Buffer | Uint8Array): PropositionTable[] {
  let bytes: Uint8Array
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
    bytes = new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
  } else {
    bytes = input instanceof Uint8Array ? input : new Uint8Array(input as unknown as ArrayBuffer)
  }
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true })
  const sheetNames = workbook.SheetNames
  if (!sheetNames.length) {
    return [
      { title: 'Proposition 1', sheetName: '', ...emptyTableBase() },
      { title: 'Proposition 2', sheetName: '', ...emptyTableBase() },
      { title: 'Proposition 3', sheetName: '', ...emptyTableBase() },
    ]
  }

  const pickedNames = pickPropositionSheetNames(sheetNames)

  return [1, 2, 3].map((num, i) => {
    const sheetName = pickedNames[i]
    const sheet =
      sheetName && workbook.Sheets[sheetName] ? workbook.Sheets[sheetName] : undefined

    if (!sheet) {
      return {
        title: `Proposition ${num}`,
        sheetName: sheetName || '',
        ...emptyTableBase(),
      }
    }

    const grid = XLSX.utils.sheet_to_json<(string | number | boolean | Date | null | undefined)[]>(
      sheet,
      { header: 1, defval: '', raw: false }
    ) as unknown[][]

    const stringGrid = grid.map((row) =>
      (Array.isArray(row) ? row : []).map((c) => cellToString(c))
    )

    const { sectionHeading, sectionDescription } = extractSectionIntro(stringGrid)

    const dataRowIdx = findFirstDataRowIndex(stringGrid)
    const headerBandIndices = findHeaderBandRowIndices(stringGrid, dataRowIdx)

    let headerRowIdx = 0
    if (headerBandIndices.length) {
      headerRowIdx = headerBandIndices[headerBandIndices.length - 1]
    } else if (dataRowIdx > 0) {
      headerRowIdx = dataRowIdx - 1
    } else {
      for (let r = 0; r < Math.min(stringGrid.length, 40); r++) {
        const row = stringGrid[r]
        const nonEmpty = row.filter((c) => c !== '').length
        if (nonEmpty >= 2) {
          headerRowIdx = r
          break
        }
      }
    }

    const scanTop = headerBandIndices.length ? headerBandIndices[0] : Math.max(0, headerRowIdx - 12)
    const normalizedHeaders = buildHeadersFromGrid(stringGrid, headerRowIdx, scanTop)

    const sheetCols = sheetColumnCount(sheet, stringGrid)
    const maxCols = Math.max(sheetCols, normalizedHeaders.length, 1)

    const headerRows = buildMergedHeaderBand(sheet, stringGrid, headerBandIndices, maxCols)

    const paddedHeaders = Array.from({ length: maxCols }, (_, j) => normalizedHeaders[j] ?? `Column ${j + 1}`)

    const bodyStart = dataRowIdx > 0 ? dataRowIdx : headerRowIdx + 1
    const rows = stringGrid.slice(bodyStart).map((row) =>
      Array.from({ length: maxCols }, (_, j) => row[j] ?? '')
    )

    const trimmedRows = rows.filter((row) => row.some((c) => c !== ''))

    return {
      title: `Proposition ${num}`,
      sheetName,
      sectionHeading,
      sectionDescription,
      headers: paddedHeaders,
      headerRows,
      rows: trimmedRows,
    }
  })
}
