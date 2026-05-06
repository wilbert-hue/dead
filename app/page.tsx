'use client'

import { useEffect, useState } from 'react'
import Image from 'next/image'
import { ChevronDown, ChevronUp, UserRound } from 'lucide-react'
import type { PropositionTable } from '@/lib/excel-propositions-types'
import { DEMO_CUSTOMER_INTEL_XLSX } from '@/lib/excel-propositions-types'
import { tryLoadTablesFromPublicWorkbook } from '@/lib/load-propositions-client'

const NAV_TEAL = 'rgb(77 182 172)'
const NAV_TEAL_BG = 'rgb(232 247 245)'
const BANNER_BG = '#3b4b5e'
const HDR_GREEN = '#e2efda'
const HDR_YELLOW = '#fde073'
const DASHBOARD_SUBTITLE = 'Dead burned Magnesia Buyers'

function emptyTablesFallback(): PropositionTable[] {
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

export default function DashboardPage() {
  const [tables, setTables] = useState<PropositionTable[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [expectedHints, setExpectedHints] = useState<string[]>([])
  const [expanded, setExpanded] = useState<Record<string, boolean>>({})

  useEffect(() => {
    let cancelled = false
    ;(async () => {
      try {
        setLoading(true)

        let parsed: PropositionTable[] | null = null
        let message: string | null = null
        let hints: string[] = []

        try {
          const res = await fetch('/api/excel-propositions', { cache: 'no-store' })
          const json = await res.json().catch(() => ({}))
          if (Array.isArray(json.tables)) {
            parsed = json.tables.slice(0, 3) as PropositionTable[]
          }
          if (!res.ok) {
            message =
              [json.error, json.details].filter(Boolean).join(' — ') || `HTTP ${res.status}`
            hints = Array.isArray(json.expectedPaths) ? json.expectedPaths : []
          }
        } catch (e) {
          const isFetch =
            e instanceof TypeError &&
            /fetch|network|failed/i.test(String((e as Error).message))
          message = isFetch
            ? 'Failed to fetch — check that `npm run dev` is running and open http://localhost:3000 (not a saved HTML file).'
            : e instanceof Error
              ? e.message
              : 'Failed to load data'
        }

        const tablesEmpty = (t: PropositionTable[]) =>
          !t.length ||
          !t.some((x) => x.headers.length > 0 || x.headerRows.length > 0 || x.rows.length > 0)

        if (!parsed || parsed.length !== 3 || tablesEmpty(parsed)) {
          const fromPublic = await tryLoadTablesFromPublicWorkbook()
          if (fromPublic && !tablesEmpty(fromPublic)) {
            parsed = fromPublic
            message = null
            hints = []
          }
        }

        if (!cancelled) {
          const finalTables = parsed && parsed.length === 3 ? parsed : emptyTablesFallback()
          setTables(finalTables)
          setError(tablesEmpty(finalTables) ? message : null)
          setExpectedHints(tablesEmpty(finalTables) ? hints : [])
        }
      } catch (e) {
        if (!cancelled) {
          setError(e instanceof Error ? e.message : 'Failed to load data')
          setExpectedHints([])
          setTables(emptyTablesFallback())
        }
      } finally {
        if (!cancelled) setLoading(false)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  const toggleSection = (key: string) => {
    setExpanded((prev) => ({ ...prev, [key]: !(prev[key] ?? true) }))
  }

  const isOpen = (key: string) => expanded[key] ?? true

  if (loading) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-[#f8f9fa]">
        <p className="text-lg text-slate-600">Loading dashboard…</p>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-[#f8f9fa] text-slate-900">
      {/* Top header */}
      <header className="border-b border-slate-200/80 bg-white">
        <div className="mx-auto flex max-w-[1600px] items-start gap-6 px-5 py-5 lg:items-center">
          <div className="relative shrink-0 pt-1">
            <Image
              src="/logo.png"
              alt="Coherent Market Insights"
              width={220}
              height={72}
              className="h-auto w-[min(42vw,220px)] object-contain object-left"
              priority
            />
          </div>

          <div className="flex flex-1 flex-col items-center justify-center text-center">
            <h1 className="text-2xl font-bold tracking-tight text-slate-900 md:text-[1.65rem]">
              Coherent Dashboard
            </h1>
            <p className="mt-1 text-sm font-medium text-slate-700">{DASHBOARD_SUBTITLE}</p>
          </div>

          <div className="hidden w-[220px] shrink-0 lg:block" aria-hidden />
        </div>
      </header>

      <div className="mx-auto flex max-w-[1600px] flex-col gap-6 px-5 py-8 lg:flex-row lg:items-start">
        {/* Sidebar */}
        <aside className="w-full shrink-0 lg:w-56 xl:w-64">
          <div className="rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="border-b border-slate-100 px-4 py-3">
              <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                Chart View
              </p>
            </div>
            <nav className="p-3">
              <div
                className="rounded-md px-3 py-3"
                style={{
                  backgroundColor: NAV_TEAL_BG,
                  borderLeft: `3px solid ${NAV_TEAL}`,
                }}
              >
                <div className="flex items-center gap-2">
                  <UserRound className="h-5 w-5 shrink-0" style={{ color: NAV_TEAL }} />
                  <span className="text-sm font-semibold text-slate-900">
                    Customer Intelligence
                  </span>
                </div>
                <p className="mt-2 pl-7 text-xs leading-snug text-slate-600">
                  Customer database with proposition tables.
                </p>
              </div>
            </nav>
          </div>
        </aside>

        {/* Main */}
        <main className="min-w-0 flex-1">
          {error && (
            <div className="mb-6 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-950">
              <p className="font-semibold">Could not load workbook.</p>
              <p className="mt-1">
                Add <span className="font-mono">{DEMO_CUSTOMER_INTEL_XLSX}</span> to{' '}
                <span className="font-mono">public/data/</span> or the project root, then refresh.
              </p>
              {expectedHints.length > 0 && (
                <ul className="mt-2 list-inside list-disc font-mono text-xs">
                  {expectedHints.map((h) => (
                    <li key={h}>{h}</li>
                  ))}
                </ul>
              )}
              <p className="mt-2 text-xs opacity-90">{error}</p>
            </div>
          )}

          <h2 className="mb-5 text-xl font-bold text-slate-900">
            Customer Intelligence Database
          </h2>

          <div className="space-y-6">
            {tables.map((table) => {
              const open = isOpen(table.title)
              const duplicateIntro =
                tables.length > 1 &&
                tables.every(
                  (t) =>
                    t.sectionHeading === tables[0].sectionHeading &&
                    t.sectionDescription === tables[0].sectionDescription
                )

              const introParagraph =
                [table.sectionHeading, table.sectionDescription].filter(Boolean).join(' ').trim() ||
                'Database covering buyers by segment, application, end-use industry, and technical requirements.'

              const heading = duplicateIntro
                ? table.title
                : table.sectionHeading.trim() || `${table.title} — Buyer segments`

              const description =
                duplicateIntro && table.title !== tables[0]?.title
                  ? 'Same database scope as Proposition 1; columns and coverage follow this proposition sheet.'
                  : introParagraph

              return (
                <section
                  key={table.title}
                  className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm"
                >
                  <button
                    type="button"
                    onClick={() => toggleSection(table.title)}
                    className="flex w-full items-start justify-between gap-4 px-5 py-4 text-left transition-colors hover:bg-slate-50/90"
                  >
                    <span className="text-base font-bold leading-snug text-slate-900">
                      {heading}
                    </span>
                    <span className="mt-0.5 shrink-0 text-slate-500">
                      {open ? (
                        <ChevronUp className="h-5 w-5" aria-hidden />
                      ) : (
                        <ChevronDown className="h-5 w-5" aria-hidden />
                      )}
                    </span>
                  </button>

                  {open && (
                    <div className="w-full border-t border-slate-200">
                      <div className="px-5 pb-5 pt-4">
                      {!(
                        table.sectionHeading.trim() || table.sectionDescription.trim()
                      ) && (
                        <p className="mb-4 text-sm italic leading-relaxed text-slate-600">
                          {description}
                        </p>
                      )}

                      <div className="overflow-x-auto rounded-md border border-neutral-900">
                        {table.headers.length === 0 ? (
                          <p className="bg-white px-4 py-8 text-center text-sm text-slate-500">
                            No columns in this sheet.
                          </p>
                        ) : (
                          <div className="w-max min-w-full">
                            {(table.sectionHeading.trim() || table.sectionDescription.trim()) && (
                              <div
                                className="box-border w-full border-b border-black px-4 py-4 text-center"
                                style={{ backgroundColor: BANNER_BG }}
                              >
                                {table.sectionHeading.trim() ? (
                                  <p className="text-base font-bold leading-snug text-white md:text-lg">
                                    {table.sectionHeading}
                                  </p>
                                ) : null}
                                {table.sectionDescription.trim() ? (
                                  <p className="mt-2 text-xs font-normal leading-snug text-white/90 md:text-sm">
                                    {table.sectionDescription}
                                  </p>
                                ) : null}
                              </div>
                            )}

                            <table className="min-w-full border-collapse text-sm">
                              <thead>
                                {table.headerRows.length > 0
                                  ? table.headerRows.map((hr, tri) => (
                                      <tr key={tri}>
                                        {hr.map((cell, hci) => (
                                          <th
                                            key={`${tri}-${hci}-${cell.startCol}`}
                                            colSpan={cell.colspan}
                                            rowSpan={cell.rowspan}
                                            className="border border-black px-2 py-2 text-center text-xs font-bold text-neutral-900"
                                            style={{
                                              backgroundColor:
                                                cell.startCol === 0 ? HDR_YELLOW : HDR_GREEN,
                                            }}
                                          >
                                            <span className="inline-block max-w-none whitespace-pre-wrap leading-snug">
                                              {cell.text}
                                            </span>
                                          </th>
                                        ))}
                                      </tr>
                                    ))
                                  : (
                                      <tr>
                                        {table.headers.map((h, hi) => (
                                          <th
                                            key={hi}
                                            className="border border-black px-2 py-2 text-center text-xs font-bold text-neutral-900"
                                            style={{
                                              backgroundColor: hi === 0 ? HDR_YELLOW : HDR_GREEN,
                                            }}
                                          >
                                            {h}
                                          </th>
                                        ))}
                                      </tr>
                                    )}
                              </thead>
                              <tbody className="bg-white">
                                {table.rows.length === 0 ? (
                                  <tr>
                                    <td
                                      colSpan={Math.max(table.headers.length, 1)}
                                      className="border border-black px-4 py-8 text-center text-slate-500"
                                    >
                                      No data rows.
                                    </td>
                                  </tr>
                                ) : (
                                  table.rows.map((row, ri) => (
                                    <tr key={ri}>
                                      {row.map((cell, ci) => (
                                        <td
                                          key={ci}
                                          className={`border border-black px-2 py-2 text-sm text-neutral-900 ${
                                            ci === 0
                                              ? 'text-center'
                                              : ci === 1
                                                ? 'text-left'
                                                : 'text-center'
                                          }`}
                                        >
                                          <span className="block whitespace-pre-wrap">
                                            {cell}
                                          </span>
                                        </td>
                                      ))}
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                      </div>
                    </div>
                  )}
                </section>
              )
            })}
          </div>
        </main>
      </div>
    </div>
  )
}
