import type { PropositionTable } from '@/lib/excel-propositions-types'

const ROW_COUNT = 40

function fitRow(row: string[], width: number): string[] {
  const next = [...row]
  while (next.length < width) next.push('—')
  return next.slice(0, width)
}

/** Fictitious company labels — no “Demo” wording; combinations vary by row index. */
function syntheticCompany(i: number): string {
  const a = [
    'Atlas', 'Euro', 'Pacific', 'Harbor', 'Andes', 'Arctic', 'Desert', 'Crescent',
    'Meridian', 'Vertex', 'Sterling', 'Granite', 'Flint', 'Titanium', 'Zenith',
    'Nordic', 'Samba', 'Orient', 'Sahel', 'Baltic',
  ]
  const b = [
    'Refractory', 'Magnesia', 'Kiln', 'Steel', 'Flux', 'Lining', 'Ceramic',
    'Brick', 'Thermal', 'Industrial', 'Fusion', 'Carbon', 'Mineral', 'Alloy',
  ]
  const c = ['Works', 'Holdings', 'Industries', 'Partners', 'Group', 'Ltd.', 'Corp.', 'Co.', 'Technologies']
  return `${a[i % a.length]} ${b[(i * 3) % b.length]} ${c[(i * 7) % c.length]}`
}

const HEADQUARTERS = [
  'Pittsburgh, PA, USA',
  'Essen, Germany',
  'Busan, South Korea',
  'Rotterdam, Netherlands',
  'São Paulo, Brazil',
  'Bergen, Norway',
  'Dubai, UAE',
  'İzmir, Turkey',
  'Jamshedpur, India',
  'Wuhan, China',
  'Hamilton, Ontario, Canada',
  'Newcastle, UK',
  'Port Hedland, Australia',
  'Monterrey, Mexico',
  'Gdansk, Poland',
  'Liege, Belgium',
  'Le Havre, France',
  'Kaohsiung, Taiwan',
  'Kitakyushu, Japan',
  'Taranto, Italy',
  'Port Talbot, UK',
  'Burnie, Australia',
  'Saldanha Bay, South Africa',
  'Pecem, Brazil',
  'Tarragona, Spain',
  'Ulsan, South Korea',
  'Johor, Malaysia',
  'Rayong, Thailand',
  'Klang, Malaysia',
  'Cartagena, Colombia',
  'Veracruz, Mexico',
  'Sukinda, India',
  'Nagoya, Japan',
  'Linz, Austria',
  'Brisbane, Australia',
  'Mobile, AL, USA',
  'Gary, IN, USA',
  'Port Kembla, Australia',
  'Constanta, Romania',
  'Iskenderun, Turkey',
  'Port Said, Egypt',
]

function hq(i: number): string {
  return HEADQUARTERS[i % HEADQUARTERS.length]
}

function buildP1Rows(): string[][] {
  const industries = ['Iron & Steel', 'Cement & Lime', 'Refractory Manufacturing', 'Others']
  const subs = ['Ladle Linings', 'Cement Kiln Linings', 'Basic Brick Mfg', 'Others']
  const status = ['Operating', 'Expansion', 'New Project', 'Operating']
  return Array.from({ length: ROW_COUNT }, (_, i) => {
    return [
      String(i + 1),
      syntheticCompany(i),
      String(1985 + (i % 32)),
      hq(i),
      `$${140 + ((i * 23) % 520)}`,
      industries[i % industries.length],
      subs[i % subs.length],
      status[i % status.length],
    ]
  })
}

function buildP2Rows(): string[][] {
  const grades = ['High purity DBM', 'Standard DBM grades', 'MgO-rich blends', 'Basic mixes', 'Custom sizing']
  const forms = ['Slip casting', 'Dry pressing', 'Iso-static', 'Casting', 'Extrusion']
  const roles = ['Technical sales lead', 'Procurement manager', 'Plant engineer', 'Buyer', 'Category manager']
  const stages = ['Qualified', 'Pilot trial', 'Discovery', 'Qualified', 'Pilot trial']
  const sizes = ['Large', 'Medium', 'Small', 'Medium', 'Large']

  return Array.from({ length: ROW_COUNT }, (_, i) => {
    return [
      String(i + 1),
      syntheticCompany(i + 11),
      String(1988 + (i % 28)),
      hq(i + 3),
      `$${130 + ((i * 19) % 480)}`,
      i % 3 === 0 ? 'Iron & Steel' : i % 3 === 1 ? 'Cement & Lime' : 'Others',
      grades[i % grades.length],
      forms[i % forms.length],
      roles[i % roles.length],
      stages[i % stages.length],
      sizes[i % sizes.length],
    ]
  })
}

function buildP3Rows(): string[][] {
  const grades = ['DBM 97%', 'DBM 95%', 'DBM HT', 'DBM standard', 'Custom MgO', 'DBM 96%', 'DBM 94%', 'Fine DBM']
  const density = ['>3.10', '>3.05', '>3.12', '>3.00', '>3.08', '>3.06', '>3.04', '>3.11']
  const forms = ['Brick', 'Castable', 'Gunning', 'Brick', 'Powder', 'Brick', 'Castable', 'Powder']
  const routes = ['Direct', 'Distributor', 'Direct', 'Multi-vendor', 'Agent', 'Direct', 'EPC', 'Direct']
  const buys = ['Annual tender', 'Spot buys', 'Framework', 'RFQ stage', 'Project-based', 'Maintenance budget', 'Annual', 'R&D samples']
  const stages = ['Qualified', 'Pilot', 'Qualified', 'Discovery', 'Pilot', 'Qualified', 'Qualified', 'Pilot']

  return Array.from({ length: ROW_COUNT }, (_, i) => {
    const j = i % grades.length
    return [
      String(i + 1),
      syntheticCompany(i + 23),
      String(1990 + (i % 26)),
      hq(i + 5),
      `$${150 + ((i * 21) % 500)}`,
      i % 4 === 0 ? 'Iron & Steel' : i % 4 === 1 ? 'Cement & Lime' : i % 4 === 2 ? 'Refractory Manufacturing' : 'Others',
      grades[j],
      density[j],
      forms[j],
      routes[j],
      buys[j],
      stages[j],
    ]
  })
}

function rowsForProposition(pi: number): string[][] {
  if (pi === 0) return buildP1Rows()
  if (pi === 1) return buildP2Rows()
  return buildP3Rows()
}

/** Replace body rows with illustrative fictional values; keeps Excel headers as-is. */
export function applyDemoRows(tables: PropositionTable[]): PropositionTable[] {
  return tables.map((t, pi) => {
    const w = t.headers.length
    if (!w) return t

    const template = rowsForProposition(pi)
    const rows = template.map((r) => fitRow(r, w))

    return {
      ...t,
      sectionHeading: 'Dead burned Magnesia Buyers',
      sectionDescription:
        'Illustrative buyer snapshot — product requirement, application, end-use industry, technical requirements, and fit (synthetic sample data).',
      rows,
    }
  })
}
