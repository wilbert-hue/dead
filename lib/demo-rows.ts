import type { PropositionTable } from '@/lib/excel-propositions-types'

function fitRow(row: string[], width: number): string[] {
  const next = [...row]
  while (next.length < width) next.push('—')
  return next.slice(0, width)
}

/** Synthetic buyer rows per proposition; trimmed/padded to match live column count. */
const DEMO_ROWS: string[][][] = [
  [
    ['1', 'Northpeak Refractory Works', '1994', 'Pittsburgh, PA', '$312', 'Iron & Steel', 'Ladle Linings', 'Operating'],
    ['2', 'EuroKiln Refractories GmbH', '2001', 'Düsseldorf, Germany', '$285', 'Cement & Lime', 'Cement Kiln Linings', 'Expansion'],
    ['3', 'Pacific Flux Industries', '1988', 'Busan, South Korea', '$410', 'Refractory Manufacturing', 'Basic Brick Mfg', 'Operating'],
    ['4', 'Harbor Steel Refractory Co.', '2009', 'Rotterdam, NL', '$156', 'Iron & Steel', 'Others', 'New Project'],
    ['5', 'Andes Magnesia Partners', '1999', 'São Paulo, Brazil', '$228', 'Others', 'Basic Brick Mfg', 'Operating'],
    ['6', 'Arctic Linings AS', '2012', 'Bergen, Norway', '$142', 'Iron & Steel', 'Ladle Linings', 'Operating'],
    ['7', 'Desert Refractory Holdings', '2005', 'Dubai, UAE', '$198', 'Cement & Lime', 'Others', 'Operating'],
    ['8', 'Crescent Technical Ceramics', '2016', 'İzmir, Turkey', '$267', 'Refractory Manufacturing', 'Cement Kiln Linings', 'Expansion'],
  ],
  [
    ['1', 'Northpeak Refractory Works', '1994', 'Pittsburgh, PA', '$312', 'Iron & Steel', 'High purity DBM', 'Slip casting', 'Technical sales lead', 'Qualified', 'Large'],
    ['2', 'EuroKiln Refractories GmbH', '2001', 'Düsseldorf, Germany', '$285', 'Cement & Lime', 'Standard DBM grades', 'Dry pressing', 'Procurement manager', 'Pilot trial', 'Medium'],
    ['3', 'Pacific Flux Industries', '1988', 'Busan, South Korea', '$410', 'Iron & Steel', 'MgO-rich blends', 'Iso-static', 'Plant engineer', 'Qualified', 'Large'],
    ['4', 'Harbor Steel Refractory Co.', '2009', 'Rotterdam, NL', '$156', 'Iron & Steel', 'Basic mixes', 'Casting', 'Buyer', 'Discovery', 'Small'],
    ['5', 'Andes Magnesia Partners', '1999', 'São Paulo, Brazil', '$228', 'Others', 'Custom sizing', 'Extrusion', 'Category manager', 'Qualified', 'Medium'],
    ['6', 'Arctic Linings AS', '2012', 'Bergen, Norway', '$142', 'Iron & Steel', 'Dense brick grades', 'Dry pressing', 'Maintenance head', 'Pilot trial', 'Small'],
    ['7', 'Desert Refractory Holdings', '2005', 'Dubai, UAE', '$198', 'Cement & Lime', 'Kiln lining specs', 'Casting', 'Procurement', 'Qualified', 'Medium'],
    ['8', 'Crescent Technical Ceramics', '2016', 'İzmir, Turkey', '$267', 'Refractory Manufacturing', 'Fine MgO powders', 'Spray drying', 'R&D contact', 'Pilot trial', 'Large'],
  ],
  [
    ['1', 'Northpeak Refractory Works', '1994', 'Pittsburgh, PA', '$312', 'Iron & Steel', 'DBM 97%', '>3.10', 'Brick', 'Direct', 'Annual tender', 'Qualified'],
    ['2', 'EuroKiln Refractories GmbH', '2001', 'Düsseldorf, Germany', '$285', 'Cement & Lime', 'DBM 95%', '>3.05', 'Castable', 'Distributor', 'Spot buys', 'Pilot'],
    ['3', 'Pacific Flux Industries', '1988', 'Busan, South Korea', '$410', 'Iron & Steel', 'DBM HT', '>3.12', 'Gunning', 'Direct', 'Framework', 'Qualified'],
    ['4', 'Harbor Steel Refractory Co.', '2009', 'Rotterdam, NL', '$156', 'Iron & Steel', 'DBM standard', '>3.00', 'Brick', 'Multi-vendor', 'RFQ stage', 'Discovery'],
    ['5', 'Andes Magnesia Partners', '1999', 'São Paulo, Brazil', '$228', 'Others', 'Custom MgO', '>3.08', 'Powder', 'Agent', 'Project-based', 'Pilot'],
    ['6', 'Arctic Linings AS', '2012', 'Bergen, Norway', '$142', 'Iron & Steel', 'DBM 96%', '>3.06', 'Brick', 'Direct', 'Maintenance budget', 'Qualified'],
    ['7', 'Desert Refractory Holdings', '2005', 'Dubai, UAE', '$198', 'Cement & Lime', 'DBM 94%', '>3.04', 'Castable', 'EPC', 'Annual', 'Qualified'],
    ['8', 'Crescent Technical Ceramics', '2016', 'İzmir, Turkey', '$267', 'Refractory Manufacturing', 'Fine DBM', '>3.11', 'Powder', 'Direct', 'R&D samples', 'Pilot'],
  ],
]

/** Replace body rows with illustrative demo values; keeps Excel headers as-is. */
export function applyDemoRows(tables: PropositionTable[]): PropositionTable[] {
  return tables.map((t, pi) => {
    const w = t.headers.length
    if (!w) return t

    const template = DEMO_ROWS[pi] ?? DEMO_ROWS[0]
    const rows = template.map((r) => fitRow(r, w))

    return {
      ...t,
      sectionHeading: 'Dead burned Magnesia Buyers',
      sectionDescription:
        'Illustrative buyer snapshot — product requirement, application, end-use industry, technical requirements, and fit (demo only).',
      rows,
    }
  })
}
