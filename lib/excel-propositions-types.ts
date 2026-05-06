export const DEMO_CUSTOMER_INTEL_XLSX =
  'Demo Customer Intelligence Database_Dead Burned Magnesia Buyers_CMI.xlsx'

/** One cell in a merged Excel header row (exported with colspan / rowspan). */
export type HeaderBandCell = {
  text: string
  colspan: number
  rowspan: number
  /** Leftmost Excel column index (0-based); used for S.No. column tint. */
  startCol: number
}

export type PropositionTable = {
  title: string
  sheetName: string
  /** Pulled from the workbook title block (first rows). */
  sectionHeading: string
  sectionDescription: string
  /** Leaf column labels (aligned with data columns), filled downward where blanks. */
  headers: string[]
  /** Multi-row thead matching workbook merges / tiers when present. */
  headerRows: HeaderBandCell[][]
  rows: string[][]
}
