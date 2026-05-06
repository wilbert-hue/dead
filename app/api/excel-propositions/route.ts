import { NextResponse } from 'next/server'
import * as fs from 'fs/promises'
import * as path from 'path'
import type { PropositionTable } from '@/lib/excel-propositions-types'
import { DEMO_CUSTOMER_INTEL_XLSX } from '@/lib/excel-propositions-types'
import { workbookToPropositionTables } from '@/lib/excel-propositions'

export const dynamic = 'force-dynamic'

function emptyPropositionTables(): PropositionTable[] {
  return [1, 2, 3].map((n) => ({
    title: `Proposition ${n}`,
    sheetName: '',
    sectionHeading: '',
    sectionDescription: '',
    headers: [],
    headerRows: [],
    rows: [],
  }))
}

async function resolveWorkbookAbsolutePath(): Promise<string | null> {
  const cwd = process.cwd()
  const candidates = [
    path.join(cwd, 'public', 'data', DEMO_CUSTOMER_INTEL_XLSX),
    path.join(cwd, DEMO_CUSTOMER_INTEL_XLSX),
  ]
  for (const p of candidates) {
    try {
      await fs.access(p)
      return p
    } catch {
      /* try next */
    }
  }
  return null
}

export async function GET() {
  try {
    const filePath = await resolveWorkbookAbsolutePath()
    if (!filePath) {
      return NextResponse.json(
        {
          error: 'Excel file not found',
          expectedPaths: [
            `public/data/${DEMO_CUSTOMER_INTEL_XLSX}`,
            DEMO_CUSTOMER_INTEL_XLSX + ' (project root)',
          ],
          tables: emptyPropositionTables(),
        },
        { status: 404 }
      )
    }
    const buffer = await fs.readFile(filePath)
    const tables = workbookToPropositionTables(new Uint8Array(buffer))
    return NextResponse.json({
      tables,
      fileName: DEMO_CUSTOMER_INTEL_XLSX,
      resolvedFrom: path.relative(process.cwd(), filePath).replace(/\\/g, '/'),
    })
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e)
    return NextResponse.json(
      {
        error: 'Failed to read workbook',
        details: message,
        tables: emptyPropositionTables(),
      },
      { status: 500 }
    )
  }
}
