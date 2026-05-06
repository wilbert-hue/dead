import type { PropositionTable } from '@/lib/excel-propositions-types'
import { DEMO_CUSTOMER_INTEL_XLSX } from '@/lib/excel-propositions-types'

/**
 * When `/api/excel-propositions` is unreachable (e.g. dev server stopped), load the
 * same workbook from `public/data/` and parse in the browser.
 */
export async function tryLoadTablesFromPublicWorkbook(): Promise<PropositionTable[] | null> {
  const name = DEMO_CUSTOMER_INTEL_XLSX
  const urls = [`/data/${encodeURIComponent(name)}`]

  for (const url of urls) {
    try {
      const res = await fetch(url, { cache: 'no-store' })
      if (!res.ok) continue
      const ab = await res.arrayBuffer()
      const { workbookToPropositionTables } = await import('@/lib/excel-propositions')
      const tables = workbookToPropositionTables(new Uint8Array(ab))
      if (Array.isArray(tables) && tables.length === 3) return tables
    } catch {
      /* try next */
    }
  }
  return null
}
