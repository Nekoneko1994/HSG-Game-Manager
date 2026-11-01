import { useEffect, useState } from 'react'
import ExcelJS from 'exceljs'
import './App.css'

// ============================================================
// HSG Game Manager
// Version 5
//
// Economic rules:
// - Blue duiker: 10 points = 10,000 FCFA
// - Red duiker: 20 points = 20,000 FCFA
// - Red river hog: 30 points = 30,000 FCFA
// - 1 food point = 1,000 FCFA
// - Annual food requirement: 50 points
// - No Hunting: 0 FCFA
// - Initial cash: 30,000 FCFA
// - No bank loans
// - Negative cash allowed
// - One gift penalty if cash ever becomes negative
//
// Wildlife rules:
// - Nearby: Blue 20, Red 10, Failure 20
// - Distant: Red 20, Hog 10, Failure 20
// - Annual regeneration: 50%
// - Three draws per hunting action
// ============================================================

// ============================================================
// Types
// ============================================================

type Choice = '' | 'nearby' | 'distant' | 'noHunting'

type Scenario = 'OA' | 'PAM' | 'CBM'

type Species = 'blue' | 'red' | 'hog'

type Stage =
  | 'round'
  | 'settlement'
  | 'regeneration'
  | 'finished'

type Population = {
  blue: number
  nearRed: number
  farRed: number
  hog: number
}

type FailureTokens = {
  nearby: number
  distant: number
}

type Action = {
  choice: Choice
  blue: number
  red: number
  hog: number
  arrested: boolean
}

type RoundRecord = {
  round: number
  actions: Action[]
}

type PopulationRecord = {
  year: number
  population: Population
}

type FinancialRecord = {
  year: number
  player: number
  food: number
  foodTransaction: number
  cash: number
}

type Game = {
  gameId: string
  scenario: Scenario

  players: number
  rounds: number
  round: number

  initialMoney: number
  fine: number
  travelCost: number

  regenerationRate: number
  foodRequirement: number
  drawsPerHunt: number

  failureTokens: FailureTokens

  actions: Action[]
  records: RoundRecord[]

  population: Population
  history: PopulationRecord[]

  cash: number[]
  financialHistory: FinancialRecord[]

  stage: Stage
}

// ============================================================
// Constants
// ============================================================

// Separate storage from the previous economic rules.
const STORAGE = 'hsg-game-manager-v5'

// New recommended initial settings.

const DEFAULT_POPULATION: Population = {
  blue: 20,
  nearRed: 10,
  farRed: 20,
  hog: 10,
}

const DEFAULT_FAILURE: FailureTokens = {
  nearby: 20,
  distant: 20,
}

const DEFAULT_INITIAL_CASH = 30000

const DEFAULT_REGENERATION_RATE = 50

const DEFAULT_FOOD_REQUIREMENT = 50

const DEFAULT_DRAWS_PER_HUNT = 3

const DEFAULT_TRAVEL_COST = 0

const DEFAULT_FINE = 20000

// Food prices and values.

const FOOD_PRICE = 1000

const BLUE_POINTS = 10
const RED_POINTS = 20
const HOG_POINTS = 30

const NO_HUNTING_INCOME = 0

// Settings dropdowns.

const TOKEN_OPTIONS = Array.from(
  { length: 7 },
  (_, i) => i * 5
)

const TENS = Array.from(
  { length: 11 },
  (_, i) => i * 10
)

const CASH_OPTIONS = Array.from(
  { length: 11 },
  (_, i) => i * 10000
)

const TRAVEL_OPTIONS = [
  0,
  10000,
  20000,
  30000,
  40000,
  50000,
]

const FINE_OPTIONS = [
  10000,
  20000,
  30000,
  50000,
  100000,
]

const ROUND_OPTIONS = Array.from(
  { length: 10 },
  (_, i) => (i + 1) * 4
)

const DRAW_OPTIONS = [1, 2, 3, 4, 5]

// ============================================================
// Formatting helpers
// ============================================================

const fmt = (n: number) =>
  `${n.toLocaleString('en-US')} FCFA`

function blank(): Action {
  return {
    choice: '',
    blue: 0,
    red: 0,
    hog: 0,
    arrested: false,
  }
}

function blanks(n: number): Action[] {
  return Array.from(
    { length: n },
    () => blank()
  )
}

function caught(a: Action): number {
  return a.blue + a.red + a.hog
}

function foodPoints(a: Action): number {
  return (
    BLUE_POINTS * a.blue +
    RED_POINTS * a.red +
    HOG_POINTS * a.hog
  )
}

function choiceLabel(choice: Choice): string {
  switch (choice) {
    case 'nearby':
      return 'Nearby Hunting'

    case 'distant':
      return 'Distant Hunting'

    case 'noHunting':
      return 'No Hunting'

    default:
      return ''
  }
}

// ============================================================
// Financial calculations
// ============================================================

function actionCash(
  a: Action,
  g: Game
): number {
  if (a.choice === 'noHunting') {
    return NO_HUNTING_INCOME
  }

  if (a.choice === 'distant') {
    const travel = g.travelCost

    const fine =
      a.arrested ? g.fine : 0

    return -(travel + fine)
  }

  return 0
}

// Distant hunting is restricted only if
// its preparation cost is positive.
//
// If preparation cost = 0 FCFA,
// negative cash does not prevent hunting.

function cannotAffordDistantHunting(
  cash: number,
  travelCost: number
): boolean {
  return (
    travelCost > 0 &&
    cash < travelCost
  )
}

// ============================================================
// Wildlife calculations
// ============================================================

const wildlifeRows = [
  {
    key: 'blue',
    name: 'Nearby — Blue duiker',
  },
  {
    key: 'nearRed',
    name: 'Nearby — Red duiker',
  },
  {
    key: 'farRed',
    name: 'Distant — Red duiker',
  },
  {
    key: 'hog',
    name: 'Distant — Red river hog',
  },
] as const

function subtract(
  population: Population,
  action: Action
): Population {
  const p = { ...population }

  if (
    action.choice === 'nearby' &&
    !action.arrested
  ) {
    p.blue -= action.blue
    p.nearRed -= action.red
  }

  if (
    action.choice === 'distant' &&
    !action.arrested
  ) {
    p.farRed -= action.red
    p.hog -= action.hog
  }

  return p
}

function growth(
  population: Population,
  rate: number
): Population {
  const r = rate / 100

  return {
    blue: Math.floor(population.blue * r),

    nearRed: Math.floor(
      population.nearRed * r
    ),

    farRed: Math.floor(
      population.farRed * r
    ),

    hog: Math.floor(population.hog * r),
  }
}

function addPopulations(
  a: Population,
  b: Population
): Population {
  return {
    blue: a.blue + b.blue,

    nearRed: a.nearRed + b.nearRed,

    farRed: a.farRed + b.farRed,

    hog: a.hog + b.hog,
  }
}

// ============================================================
// Annual food settlement
// ============================================================

function settlement(
  g: Game
): FinancialRecord[] {
  const year = g.round / 4

  return Array.from(
    { length: g.players },
    (_, i) => {
      const food = g.records
        .filter(
          record =>
            Math.ceil(record.round / 4) ===
            year
        )
        .reduce(
          (sum, record) =>
            sum +
            foodPoints(record.actions[i]),
          0
        )

      const foodTransaction =
        (food - g.foodRequirement) *
        FOOD_PRICE

      return {
        year,
        player: i + 1,
        food,
        foodTransaction,
        cash:
          g.cash[i] + foodTransaction,
      }
    }
  )
}

// ============================================================
// Reconstruct financial history
//
// This is also used for gift penalties.
//
// A player receives only one negative-cash
// penalty even if the balance becomes
// negative multiple times.
// ============================================================

function replay(
  g: Game,
  records: RoundRecord[],
  finances: FinancialRecord[]
) {
  const cash = Array(g.players).fill(
    g.initialMoney
  ) as number[]

  const everNegative = Array(g.players).fill(
    g.initialMoney < 0
  ) as boolean[]

  const settled = new Map<
    number,
    FinancialRecord[]
  >()

  for (const f of finances) {
    const group =
      settled.get(f.year) ?? []

    group.push(f)

    settled.set(f.year, group)
  }

  const sortedRecords = [...records].sort(
    (a, b) => a.round - b.round
  )

  for (const record of sortedRecords) {
    for (
      let i = 0;
      i < g.players;
      i++
    ) {
      cash[i] += actionCash(
        record.actions[i],
        g
      )

      if (cash[i] < 0) {
        everNegative[i] = true
      }
    }

    if (record.round % 4 === 0) {
      const year = record.round / 4

      const rows =
        settled.get(year) ?? []

      for (const f of rows) {
        const i = f.player - 1

        if (
          i >= 0 &&
          i < g.players
        ) {
          cash[i] +=
            f.foodTransaction

          if (cash[i] < 0) {
            everNegative[i] = true
          }
        }
      }
    }
  }

  return {
    cash,
    everNegative,
  }
}

// ============================================================
// Final population before regeneration
// ============================================================

function finalBeforeGrowth(
  g: Game
): Population {
  const finalYear = g.rounds / 4

  const start = g.history.find(
    h => h.year === finalYear
  )

  if (!start) {
    return { ...g.population }
  }

  let population = {
    ...start.population,
  }

  const finalRecords = g.records.filter(
    r =>
      Math.ceil(r.round / 4) ===
      finalYear
  )

  for (const record of finalRecords) {
    for (const action of record.actions) {
      population = subtract(
        population,
        action
      )
    }
  }

  return population
}

// ============================================================
// Gift rewards
// ============================================================

function rewards(g: Game) {
  const reconstructed = replay(
    g,
    g.records,
    g.financialHistory
  )

  const negative =
    reconstructed.everNegative

  const finalPopulation =
    finalBeforeGrowth(g)

  const wildlife = Object.values(
    finalPopulation
  ).every(n => n > 0)

  const highest = Math.max(...g.cash)

  const rows = g.cash.map(
    (cash, i) => {
      const participation = 1

      const wildlifeBonus =
        wildlife ? 1 : 0

      const negativePenalty =
        negative[i] ? -1 : 0

      const highestBonus =
        cash === highest ? 1 : 0

      const total = Math.max(
        0,
        participation +
          wildlifeBonus +
          negativePenalty +
          highestBonus
      )

      return {
        player: i + 1,

        participation,

        wildlifeBonus,

        negativePenalty,

        highestBonus,

        everNegative: negative[i],

        finalCash: cash,

        total,
      }
    }
  )

  const total = rows.reduce(
    (sum, row) => sum + row.total,
    0
  )

  return {
    rows,
    wildlife,
    finalPopulation,
    total,
  }
}

// ============================================================
// Excel export
// Five worksheets
// ============================================================

async function exportExcel(g: Game) {
  try {
    const workbook =
      new ExcelJS.Workbook()

    workbook.creator =
      'HSG Game Manager'

    // --------------------------------------------------------
    // Sheet 1: Round Decisions
    // --------------------------------------------------------

    const decisions =
      workbook.addWorksheet(
        'Round Decisions'
      )

    decisions.addRow([
      'Game ID',
      'Scenario',
      'Round',
      'Year',
      'Player',
      'Choice',
      'Arrested',
      'Blue duiker',
      'Red duiker',
      'Red river hog',
      'Food points',
      'Cash change (FCFA)',
    ])

    g.records.forEach(record => {
      record.actions.forEach(
        (action, i) => {
          decisions.addRow([
            g.gameId,
            g.scenario,
            record.round,
            Math.ceil(record.round / 4),
            i + 1,
            choiceLabel(action.choice),
            action.arrested,
            action.blue,
            action.red,
            action.hog,
            foodPoints(action),
            actionCash(action, g),
          ])
        }
      )
    })

    // --------------------------------------------------------
    // Sheet 2: Annual Finances
    // --------------------------------------------------------

    const finances =
      workbook.addWorksheet(
        'Annual Finances'
      )

    finances.addRow([
      'Game ID',
      'Year',
      'Player',
      'Food points',
      'Required food points',
      'Food transaction (FCFA)',
      'Year-end cash (FCFA)',
      'Negative year-end balance',
    ])

    g.financialHistory.forEach(
      record => {
        finances.addRow([
          g.gameId,
          record.year,
          record.player,
          record.food,
          g.foodRequirement,
          record.foodTransaction,
          record.cash,
          record.cash < 0,
        ])
      }
    )

    // --------------------------------------------------------
    // Sheet 3: Wildlife Populations
    // --------------------------------------------------------

    const populations =
      workbook.addWorksheet(
        'Wildlife Populations'
      )

    populations.addRow([
      'Game ID',
      'Beginning of year',
      'Nearby blue duiker',
      'Nearby red duiker',
      'Distant red duiker',
      'Distant red river hog',
    ])

    g.history.forEach(record => {
      populations.addRow([
        g.gameId,
        record.year,
        record.population.blue,
        record.population.nearRed,
        record.population.farRed,
        record.population.hog,
      ])
    })

    // --------------------------------------------------------
    // Sheet 4: Game Settings
    // --------------------------------------------------------

    const settings =
      workbook.addWorksheet(
        'Game Settings'
      )

    settings.addRow([
      'Parameter',
      'Value',
    ])

    const initial =
      g.history[0].population

    const settingRows: [
      string,
      string | number
    ][] = [
      [
        'Game ID',
        g.gameId,
      ],

      [
        'Scenario',
        g.scenario,
      ],

      [
        'Number of players',
        g.players,
      ],

      [
        'Number of rounds',
        g.rounds,
      ],

      [
        'Initial cash (FCFA)',
        g.initialMoney,
      ],

      [
        'Distant hunting cost (FCFA)',
        g.travelCost,
      ],

      [
        'PAM fine (FCFA)',
        g.scenario === 'PAM'
          ? g.fine
          : 'N/A',
      ],

      [
        'Annual regeneration rate (%)',
        g.regenerationRate,
      ],

      [
        'Annual food requirement',
        g.foodRequirement,
      ],

      [
        'Food price per point (FCFA)',
        FOOD_PRICE,
      ],

      [
        'Draws per hunt',
        g.drawsPerHunt,
      ],

      [
        'Blue duiker food points',
        BLUE_POINTS,
      ],

      [
        'Red duiker food points',
        RED_POINTS,
      ],

      [
        'Red river hog food points',
        HOG_POINTS,
      ],

      [
        'Nearby blue duiker',
        initial.blue,
      ],

      [
        'Nearby red duiker',
        initial.nearRed,
      ],

      [
        'Nearby failure tokens',
        g.failureTokens.nearby,
      ],

      [
        'Distant red duiker',
        initial.farRed,
      ],

      [
        'Distant red river hog',
        initial.hog,
      ],

      [
        'Distant failure tokens',
        g.failureTokens.distant,
      ],

      [
        'No Hunting income (FCFA)',
        NO_HUNTING_INCOME,
      ],

      [
        'Negative balance rule',
        'Negative cash allowed; no bank loans',
      ],

      [
        'Gift penalty',
        'Minus one if cash was negative at least once',
      ],

      [
        'Distant hunting restriction',
        'Positive travel cost requires sufficient cash',
      ],
    ]

    settingRows.forEach(
      row => settings.addRow(row)
    )

    // --------------------------------------------------------
    // Sheet 5: Gift Rewards
    // --------------------------------------------------------

    const gifts =
      workbook.addWorksheet(
        'Gift Rewards'
      )

    gifts.addRow([
      'Game ID',
      'Player',
      'Participation',
      'Wildlife survival bonus',
      'Negative balance penalty',
      'Highest cash bonus',
      'Total gifts',
      'Ever negative',
      'Final cash (FCFA)',
      'All populations survived',
    ])

    const summary = rewards(g)

    summary.rows.forEach(row => {
      gifts.addRow([
        g.gameId,
        row.player,
        row.participation,
        row.wildlifeBonus,
        row.negativePenalty,
        row.highestBonus,
        row.total,
        row.everNegative,
        row.finalCash,
        summary.wildlife,
      ])
    })

    gifts.addRow([])

    gifts.addRow([
      'Final population before regeneration',
    ])

    wildlifeRows.forEach(row => {
      gifts.addRow([
        row.name,
        summary.finalPopulation[row.key],
      ])
    })

    gifts.addRow([
      'Total gifts required',
      summary.total,
    ])

    // --------------------------------------------------------
    // Excel formatting
    // --------------------------------------------------------

    workbook.worksheets.forEach(sheet => {
      const header = sheet.getRow(1)

      header.font = {
        bold: true,
        color: {
          argb: 'FFFFFFFF',
        },
      }

      header.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF176447',
        },
      }

      header.alignment = {
        vertical: 'middle',
        wrapText: true,
      }

      header.height = 40

      sheet.views = [
        {
          state: 'frozen',
          ySplit: 1,
        },
      ]

      sheet.columns.forEach(column => {
        column.width = 23
      })

      if (
        sheet.name === 'Game Settings'
      ) {
        sheet.getColumn(1).width = 46
      }

      sheet.eachRow(
        (row, number) => {
          if (
            number > 1 &&
            number % 2 === 0
          ) {
            row.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: {
                argb: 'FFF2F7F3',
              },
            }
          }
        }
      )

      if (
        sheet.name !== 'Game Settings'
      ) {
        sheet.autoFilter = {
          from: {
            row: 1,
            column: 1,
          },

          to: {
            row: 1,
            column: sheet.columnCount,
          },
        }
      }
    })

    // --------------------------------------------------------
    // Create and download Excel
    // --------------------------------------------------------

    const data =
      await workbook.xlsx.writeBuffer()

    const bytes = new Uint8Array(
      data
    )

    const buffer = new ArrayBuffer(
      bytes.byteLength
    )

    new Uint8Array(buffer).set(
      bytes
    )

    const blob = new Blob(
      [buffer],
      {
        type:
          'application/vnd.openxmlformats-officedocument.' +
          'spreadsheetml.sheet',
      }
    )

    const url =
      URL.createObjectURL(blob)

    const link =
      document.createElement('a')

    link.href = url

    const safeGameId =
      g.gameId.replace(
        /[^a-zA-Z0-9_-]/g,
        '_'
      ) || 'HSG'

    link.download =
      `${safeGameId}.xlsx`

    document.body.appendChild(
      link
    )

    link.click()

    link.remove()

    setTimeout(
      () => URL.revokeObjectURL(url),
      1000
    )
  } catch (error) {
    console.error(error)

    window.alert(
      'Excel export failed.'
    )
  }
}

// ============================================================
// Population table
// ============================================================

function PopulationTable({
  population,
  increase,
}: {
  population: Population
  increase?: Population
}) {
  return (
    <div className="table-container">
      <table className="population-table">
        <thead>
          <tr>
            <th>Area / Species</th>

            <th>Animals</th>

            {increase && (
              <th>To add</th>
            )}
          </tr>
        </thead>

        <tbody>
          {wildlifeRows.map(row => (
            <tr key={row.key}>
              <td>
                {row.name}
              </td>

              <td>
                {population[row.key]}
              </td>

              {increase && (
                <td>
                  +{increase[row.key]}
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ============================================================
// Population graph
// ============================================================

function Graph({
  history,
  area,
  years,
}: {
  history: PopulationRecord[]
  area: 'nearby' | 'distant'
  years: number
}) {
  const series =
    area === 'nearby'
      ? [
          {
            key: 'blue' as const,
            name: 'Blue duiker',
            color: '#2563eb',
          },
          {
            key: 'nearRed' as const,
            name: 'Red duiker',
            color: '#dc2626',
          },
        ]
      : [
          {
            key: 'farRed' as const,
            name: 'Red duiker',
            color: '#dc2626',
          },
          {
            key: 'hog' as const,
            name: 'Red river hog',
            color: '#d97706',
          },
        ]

  const maximum = Math.max(
    10,
    ...history.flatMap(record =>
      series.map(
        s => record.population[s.key]
      )
    )
  )

  const top =
    Math.ceil(maximum / 10) * 10

  const count =
    Math.max(2, years + 1)

  const x = (year: number) =>
    42 +
    ((year - 1) / (count - 1)) * 290

  const y = (n: number) =>
    190 - (n / top) * 165

  return (
    <div className="chart-box">
      <h3>
        {area === 'nearby'
          ? 'Nearby Area'
          : 'Distant Area'}
      </h3>

      <svg
        className="population-chart"
        viewBox="0 0 350 240"
        role="img"
        aria-label={`${area} population trend`}
      >
        {[0, 0.25, 0.5, 0.75, 1].map(
          value => (
            <g key={value}>
              <line
                x1="42"
                x2="332"
                y1={y(value * top)}
                y2={y(value * top)}
                stroke="#e5e7eb"
              />

              <text
                x="36"
                y={y(value * top) + 4}
                textAnchor="end"
                fontSize="11"
              >
                {Math.round(value * top)}
              </text>
            </g>
          )
        )}

        {Array.from(
          { length: count },
          (_, i) => i + 1
        ).map(year => (
          <text
            key={year}
            x={x(year)}
            y="210"
            textAnchor="middle"
            fontSize="11"
          >
            {year}
          </text>
        ))}

        {series.map(s => (
          <g key={s.key}>
            <polyline
              fill="none"
              stroke={s.color}
              strokeWidth="3"
              points={history
                .map(
                  record =>
                    `${x(record.year)},${y(
                      record.population[s.key]
                    )}`
                )
                .join(' ')}
            />

            {history.map(record => (
              <circle
                key={record.year}
                cx={x(record.year)}
                cy={y(
                  record.population[s.key]
                )}
                r="4"
                fill={s.color}
              >
                <title>
                  {s.name}: Year {record.year}
                  {' — '}
                  {record.population[s.key]}
                </title>
              </circle>
            ))}
          </g>
        ))}

        <text
          x="185"
          y="234"
          textAnchor="middle"
          fontSize="12"
        >
          Year (beginning)
        </text>
      </svg>

      <div className="chart-legend">
        {series.map(s => (
          <span key={s.key}>
            <i
              className="legend-dot"
              style={{
                backgroundColor: s.color,
              }}
            />

            {s.name}
          </span>
        ))}
      </div>
    </div>
  )
}

// ============================================================
// Wildlife monitoring
// ============================================================

function Monitoring({
  history,
  years,
}: {
  history: PopulationRecord[]
  years: number
}) {
  return (
    <section className="monitoring-panel">
      <h2>
        Wildlife Population Monitoring
      </h2>

      <p className="muted">
        Populations at the beginning
        of each year.
      </p>

      <Graph
        history={history}
        area="nearby"
        years={years}
      />

      <Graph
        history={history}
        area="distant"
        years={years}
      />
    </section>
  )
}

// ============================================================
// Excel export panel
// ============================================================

function ExportPanel({
  game,
}: {
  game: Game
}) {
  const [busy, setBusy] =
    useState(false)

  return (
    <section className="card">
      <h2>
        Export Research Data
      </h2>

      <p className="muted">
        Five worksheets in one Excel workbook.
      </p>

      <button
        className="primary-button"
        disabled={busy}
        onClick={async () => {
          setBusy(true)

          try {
            await exportExcel(game)
          } finally {
            setBusy(false)
          }
        }}
      >
        {busy
          ? 'Creating Excel File...'
          : 'Export Excel Workbook (.xlsx)'}
      </button>
    </section>
  )
}

// ============================================================
// Gift rewards panel
// ============================================================

function GiftPanel({
  game,
}: {
  game: Game
}) {
  const summary = rewards(game)

  return (
    <section className="card">
      <h2>Gift Rewards</h2>

      <p>
        <strong>
          {summary.wildlife
            ? 'All four populations survived.'
            : 'One or more populations became extinct.'}
        </strong>
      </p>

      <p className="muted">
        Wildlife is evaluated before
        final regeneration.
        Negative cash incurs a one-time
        penalty, even if subsequently
        recovered.
      </p>

      <PopulationTable
        population={summary.finalPopulation}
      />

      <h3>
        Rewards by Player
      </h3>

      <div className="table-container">
        <table className="population-table">
          <thead>
            <tr>
              <th>Player</th>
              <th>Participation</th>
              <th>Wildlife</th>
              <th>Negative Cash</th>
              <th>Highest Cash</th>
              <th>Total Gifts</th>
            </tr>
          </thead>

          <tbody>
            {summary.rows.map(row => (
              <tr key={row.player}>
                <td>
                  Player {row.player}
                </td>

                <td>
                  +{row.participation}
                </td>

                <td>
                  +{row.wildlifeBonus}
                </td>

                <td>
                  {row.negativePenalty}
                </td>

                <td>
                  +{row.highestBonus}
                </td>

                <td>
                  <strong>
                    {row.total}
                  </strong>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h3>
        Total Gifts Required:
        {' '}{summary.total}
      </h3>
    </section>
  )
}

// ============================================================
// Main application
// ============================================================

function App() {

  // ----------------------------------------------------------
  // New game settings
  // ----------------------------------------------------------

  const [id, setId] =
    useState('HSG-001')

  const [scenario, setScenario] =
    useState<Scenario>('OA')

  const [players, setPlayers] =
    useState(4)

  const [initialMoney, setInitialMoney] =
    useState(DEFAULT_INITIAL_CASH)

  const [rounds, setRounds] =
    useState(20)

  const [fine, setFine] =
    useState(DEFAULT_FINE)

  const [travelCost, setTravelCost] =
    useState(DEFAULT_TRAVEL_COST)

  const [
    regenerationRate,
    setRegenerationRate,
  ] = useState(
    DEFAULT_REGENERATION_RATE
  )

  const [
    foodRequirement,
    setFoodRequirement,
  ] = useState(
    DEFAULT_FOOD_REQUIREMENT
  )

  const [
    drawsPerHunt,
    setDrawsPerHunt,
  ] = useState(
    DEFAULT_DRAWS_PER_HUNT
  )

  const [initialPop, setInitialPop] =
    useState<Population>({
      ...DEFAULT_POPULATION,
    })

  const [
    failureTokens,
    setFailureTokens,
  ] = useState<FailureTokens>({
    ...DEFAULT_FAILURE,
  })

  const [showGraphs, setShowGraphs] =
    useState(false)

  // ----------------------------------------------------------
  // Load saved game
  // ----------------------------------------------------------

  const [game, setGame] =
    useState<Game | null>(() => {
      try {
        const text =
          localStorage.getItem(STORAGE)

        if (!text) {
          return null
        }

        return JSON.parse(text) as Game
      } catch {
        return null
      }
    })

  // ----------------------------------------------------------
  // Autosave
  // ----------------------------------------------------------

  useEffect(() => {
    if (game) {
      localStorage.setItem(
        STORAGE,
        JSON.stringify(game)
      )
    } else {
      localStorage.removeItem(
        STORAGE
      )
    }
  }, [game])

  // ----------------------------------------------------------
  // Reset settings
  // ----------------------------------------------------------

  function resetSettings() {
    setId('HSG-001')

    setScenario('OA')

    setPlayers(4)

    setInitialMoney(
      DEFAULT_INITIAL_CASH
    )

    setRounds(20)

    setFine(DEFAULT_FINE)

    setTravelCost(
      DEFAULT_TRAVEL_COST
    )

    setRegenerationRate(
      DEFAULT_REGENERATION_RATE
    )

    setFoodRequirement(
      DEFAULT_FOOD_REQUIREMENT
    )

    setDrawsPerHunt(
      DEFAULT_DRAWS_PER_HUNT
    )

    setInitialPop({
      ...DEFAULT_POPULATION,
    })

    setFailureTokens({
      ...DEFAULT_FAILURE,
    })

    setShowGraphs(false)
  }

  // ----------------------------------------------------------
  // Start game
  // ----------------------------------------------------------

  function start() {
    const valid =
      id.trim().length > 0 &&
      Number.isInteger(players) &&
      players >= 2 &&
      players <= 10 &&
      CASH_OPTIONS.includes(initialMoney) &&
      TRAVEL_OPTIONS.includes(travelCost) &&
      ROUND_OPTIONS.includes(rounds) &&
      FINE_OPTIONS.includes(fine) &&
      TENS.includes(regenerationRate) &&
      TENS.includes(foodRequirement) &&
      DRAW_OPTIONS.includes(drawsPerHunt) &&
      Object.values(initialPop).every(
        n => TOKEN_OPTIONS.includes(n)
      ) &&
      Object.values(failureTokens).every(
        n => TOKEN_OPTIONS.includes(n)
      )

    if (!valid) {
      window.alert(
        'Please check the game settings.'
      )

      return
    }

    setGame({
      gameId: id.trim(),

      scenario,

      players,
      rounds,
      round: 1,

      initialMoney,
      fine,
      travelCost,

      regenerationRate,
      foodRequirement,
      drawsPerHunt,

      failureTokens: {
        ...failureTokens,
      },

      actions: blanks(players),

      records: [],

      population: {
        ...initialPop,
      },

      history: [
        {
          year: 1,

          population: {
            ...initialPop,
          },
        },
      ],

      cash: Array(players).fill(
        initialMoney
      ),

      financialHistory: [],

      stage: 'round',
    })
  }

  // ----------------------------------------------------------
  // Choose action
  // ----------------------------------------------------------

  function choose(
    i: number,
    choice: Choice
  ) {
    if (
      !game ||
      game.stage !== 'round'
    ) {
      return
    }

    if (
      choice === 'distant' &&
      cannotAffordDistantHunting(
        game.cash[i],
        game.travelCost
      )
    ) {
      return
    }

    setGame(current => {
      if (
        !current ||
        current.stage !== 'round'
      ) {
        return current
      }

      return {
        ...current,

        actions: current.actions.map(
          (action, j) =>
            i === j
              ? {
                  ...blank(),
                  choice,
                }
              : action
        ),
      }
    })
  }

  // ----------------------------------------------------------
  // Update catch
  // ----------------------------------------------------------

  function catchChange(
    i: number,
    species: Species,
    delta: number
  ) {
    setGame(current => {
      if (
        !current ||
        current.stage !== 'round'
      ) {
        return current
      }

      const action =
        current.actions[i]

      if (
        action.arrested ||
        (
          action.choice !== 'nearby' &&
          action.choice !== 'distant'
        ) ||
        (
          action.choice === 'nearby' &&
          species === 'hog'
        ) ||
        (
          action.choice === 'distant' &&
          species === 'blue'
        ) ||
        (
          delta > 0 &&
          caught(action) >=
            current.drawsPerHunt
        ) ||
        (
          delta < 0 &&
          action[species] <= 0
        )
      ) {
        return current
      }

      const actions =
        current.actions.map(
          (item, j) =>
            j === i
              ? {
                  ...item,
                  [species]:
                    item[species] + delta,
                }
              : item
        )

      const remaining =
        actions.reduce(
          subtract,
          { ...current.population }
        )

      if (
        Object.values(remaining).some(
          n => n < 0
        )
      ) {
        return current
      }

      return {
        ...current,
        actions,
      }
    })
  }

  // ----------------------------------------------------------
  // PAM arrest
  // ----------------------------------------------------------

  function toggleArrest(i: number) {
    setGame(current => {
      if (
        !current ||
        current.stage !== 'round' ||
        current.scenario !== 'PAM'
      ) {
        return current
      }

      return {
        ...current,

        actions: current.actions.map(
          (action, j) =>
            i === j &&
            action.choice === 'distant'
              ? {
                  ...action,

                  arrested:
                    !action.arrested,

                  blue: 0,
                  red: 0,
                  hog: 0,
                }
              : action
        ),
      }
    })
  }

  // ----------------------------------------------------------
  // Confirm round
  // ----------------------------------------------------------

  function confirmRound() {
    if (
      !game ||
      game.stage !== 'round' ||
      game.actions.some(
        action => !action.choice
      )
    ) {
      return
    }

    for (
      let i = 0;
      i < game.players;
      i++
    ) {
      const action =
        game.actions[i]

      if (
        action.choice === 'distant' &&
        cannotAffordDistantHunting(
          game.cash[i],
          game.travelCost
        )
      ) {
        window.alert(
          `Player ${i + 1} cannot afford distant hunting.`
        )

        return
      }

      if (
        caught(action) >
        game.drawsPerHunt
      ) {
        window.alert(
          'Too many animals caught.'
        )

        return
      }
    }

    const remaining =
      game.actions.reduce(
        subtract,
        { ...game.population }
      )

    if (
      Object.values(remaining).some(
        n => n < 0
      )
    ) {
      window.alert(
        'Catches exceed remaining populations.'
      )

      return
    }

    const cash = game.cash.map(
      (balance, i) =>
        balance +
        actionCash(game.actions[i], game)
    )

    const records = [
      ...game.records,

      {
        round: game.round,

        actions: game.actions.map(
          action => ({ ...action })
        ),
      },
    ]

    const yearEnd =
      game.round % 4 === 0

    setGame({
      ...game,

      records,

      cash,

      population: remaining,

      stage: yearEnd
        ? 'settlement'
        : 'round',

      round: yearEnd
        ? game.round
        : game.round + 1,

      actions: blanks(game.players),
    })

    window.scrollTo(0, 0)
  }

  // ----------------------------------------------------------
  // Confirm annual food settlement
  // ----------------------------------------------------------

  function confirmSettlement() {
    if (
      !game ||
      game.stage !== 'settlement'
    ) {
      return
    }

    const rows = settlement(game)

    setGame({
      ...game,

      cash: rows.map(
        row => row.cash
      ),

      financialHistory: [
        ...game.financialHistory,
        ...rows,
      ],

      stage: 'regeneration',
    })

    window.scrollTo(0, 0)
  }

  // ----------------------------------------------------------
  // Confirm annual regeneration
  // ----------------------------------------------------------

  function confirmRegeneration() {
    if (
      !game ||
      game.stage !== 'regeneration'
    ) {
      return
    }

    const increase = growth(
      game.population,
      game.regenerationRate
    )

    const next = addPopulations(
      game.population,
      increase
    )

    const finished =
      game.round === game.rounds

    setGame({
      ...game,

      population: next,

      history: [
        ...game.history,

        {
          year:
            game.round / 4 + 1,

          population: next,
        },
      ],

      stage: finished
        ? 'finished'
        : 'round',

      round: finished
        ? game.round
        : game.round + 1,

      actions: blanks(game.players),
    })

    window.scrollTo(0, 0)
  }

  // ----------------------------------------------------------
  // Edit previous round
  // ----------------------------------------------------------

  function previousRound() {
    if (
      !game ||
      !game.records.length
    ) {
      return
    }

    if (
      !window.confirm(
        'Edit the last confirmed round?'
      )
    ) {
      return
    }

    const last =
      game.records[
        game.records.length - 1
      ]

    const year =
      Math.ceil(last.round / 4)

    const history =
      game.history.filter(
        record => record.year <= year
      )

    const yearStart =
      history.find(
        record => record.year === year
      )?.population

    if (!yearStart) {
      return
    }

    let population = {
      ...yearStart,
    }

    const records =
      game.records.slice(0, -1)

    const recordsThisYear =
      records.filter(
        record =>
          Math.ceil(record.round / 4) ===
          year
      )

    for (
      const record of recordsThisYear
    ) {
      for (
        const action of record.actions
      ) {
        population = subtract(
          population,
          action
        )
      }
    }

    const financialHistory =
      game.financialHistory.filter(
        record => record.year < year
      )

    const reconstructed = replay(
      game,
      records,
      financialHistory
    )

    setGame({
      ...game,

      round: last.round,

      actions: last.actions.map(
        action => ({ ...action })
      ),

      records,

      population,

      history,

      cash: reconstructed.cash,

      financialHistory,

      stage: 'round',
    })

    window.scrollTo(0, 0)
  }

  // ----------------------------------------------------------
  // Delete game
  // ----------------------------------------------------------

  function reset() {
    if (
      !window.confirm(
        'Delete this game? All recorded data will be lost.'
      )
    ) {
      return
    }

    resetSettings()

    setGame(null)
  }

  // ----------------------------------------------------------
  // Dropdown helper
  // ----------------------------------------------------------

  const select = (
    id: string,
    label: string,
    value: number,
    change: (n: number) => void,
    values: number[],
    suffix = ''
  ) => (
    <>
      <label htmlFor={id}>
        {label}
      </label>

      <select
        id={id}
        value={value}
        onChange={event =>
          change(
            Number(event.target.value)
          )
        }
      >
        {values.map(n => (
          <option
            key={n}
            value={n}
          >
            {n.toLocaleString('en-US')}
            {suffix}
          </option>
        ))}
      </select>
    </>
  )

  // ==========================================================
  // NEW GAME SCREEN
  // ==========================================================

  if (!game) {
    return (
      <main className="app">
        <header className="header">
          <h1>
            Hunting Serious Game
          </h1>

          <p>
            Game management and wildlife monitoring
          </p>
        </header>

        <section className="card">
          <h2>New Game</h2>

          <label htmlFor="game-id">
            Game ID
          </label>

          <input
            id="game-id"
            value={id}
            onChange={event =>
              setId(
                event.target.value
              )
            }
          />

          <label htmlFor="scenario">
            Scenario
          </label>

          <select
            id="scenario"
            value={scenario}
            onChange={event =>
              setScenario(
                event.target.value as Scenario
              )
            }
          >
            <option value="OA">
              Open Access (OA)
            </option>

            <option value="PAM">
              Protected Area Management (PAM)
            </option>

            <option value="CBM">
              Community-Based Management (CBM)
            </option>
          </select>

          {scenario === 'PAM' &&
            select(
              'fine',
              'Fine amount (FCFA)',
              fine,
              setFine,
              FINE_OPTIONS,
              ' FCFA'
            )}

          <label htmlFor="players">
            Number of players
          </label>

          <input
            id="players"
            type="number"
            min="2"
            max="10"
            value={players}
            onChange={event =>
              setPlayers(
                Number(event.target.value)
              )
            }
          />

          {select(
            'cash',
            'Initial cash (FCFA)',
            initialMoney,
            setInitialMoney,
            CASH_OPTIONS,
            ' FCFA'
          )}

          {select(
            'travel',
            'Distant hunting cost (FCFA)',
            travelCost,
            setTravelCost,
            TRAVEL_OPTIONS,
            ' FCFA'
          )}

          {select(
            'rounds',
            'Number of rounds',
            rounds,
            setRounds,
            ROUND_OPTIONS,
            ' rounds'
          )}

          <h3>
            Game Rules
          </h3>

          {select(
            'regen',
            'Annual regeneration',
            regenerationRate,
            setRegenerationRate,
            TENS,
            '% (rounded down)'
          )}

          {select(
            'food',
            'Food requirement (points/player/year)',
            foodRequirement,
            setFoodRequirement,
            TENS,
            ' points'
          )}

          {select(
            'draws',
            'Draws per hunt',
            drawsPerHunt,
            setDrawsPerHunt,
            DRAW_OPTIONS,
            ' draws'
          )}

          <h3>
            Initial Wildlife Populations
            &amp; Failure Tokens
          </h3>

          <p className="muted">
            Select the initial number
            of each token type.
          </p>

          <h4>
            Nearby Area
          </h4>

          {select(
            'blue',
            'Blue duiker',
            initialPop.blue,
            n =>
              setInitialPop(p => ({
                ...p,
                blue: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          {select(
            'nearred',
            'Red duiker',
            initialPop.nearRed,
            n =>
              setInitialPop(p => ({
                ...p,
                nearRed: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          {select(
            'nearfail',
            'Failure tokens',
            failureTokens.nearby,
            n =>
              setFailureTokens(p => ({
                ...p,
                nearby: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          <h4>
            Distant Area
          </h4>

          {select(
            'farred',
            'Red duiker',
            initialPop.farRed,
            n =>
              setInitialPop(p => ({
                ...p,
                farRed: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          {select(
            'hog',
            'Red river hog',
            initialPop.hog,
            n =>
              setInitialPop(p => ({
                ...p,
                hog: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          {select(
            'farfail',
            'Failure tokens',
            failureTokens.distant,
            n =>
              setFailureTokens(p => ({
                ...p,
                distant: n,
              })),
            TOKEN_OPTIONS,
            ' tokens'
          )}

          <div className="settings">
            <p>
              <strong>
                Blue duiker:
              </strong>
              {' '}
              10 points = 10,000 FCFA
            </p>

            <p>
              <strong>
                Red duiker:
              </strong>
              {' '}
              20 points = 20,000 FCFA
            </p>

            <p>
              <strong>
                Red river hog:
              </strong>
              {' '}
              30 points = 30,000 FCFA
            </p>

            <p>
              <strong>
                Food price:
              </strong>
              {' '}
              1,000 FCFA per point
            </p>

            <p>
              <strong>
                No Hunting:
              </strong>
              {' '}
              0 FCFA income
            </p>

            <p>
              <strong>
                Negative balance:
              </strong>
              {' '}
              Allowed; no bank loans
            </p>

            <p>
              <strong>
                Distant hunting:
              </strong>
              {' '}
              No cash restriction when
              travel cost is 0 FCFA
            </p>
          </div>

          <button
            className="primary-button"
            onClick={start}
          >
            Start Game
          </button>
        </section>
      </main>
    )
  }

  // ==========================================================
  // Shared UI
  // ==========================================================

  const header = (
    <header className="header">
      <h1>
        HSG Game Manager
      </h1>

      <p>
        Game: {game.gameId}
        {' — '}
        {game.scenario}
      </p>
    </header>
  )

  const monitor = (
    history: PopulationRecord[]
  ) => (
    <Monitoring
      history={history}
      years={game.rounds / 4}
    />
  )

  const failureInfo = (
    <div className="settings">
      <p>
        <strong>
          Failure tokens — Nearby:
        </strong>
        {' '}
        {game.failureTokens.nearby}
      </p>

      <p>
        <strong>
          Failure tokens — Distant:
        </strong>
        {' '}
        {game.failureTokens.distant}
      </p>
    </div>
  )

  // ==========================================================
  // ANNUAL FOOD SETTLEMENT
  // ==========================================================

  if (
    game.stage === 'settlement'
  ) {
    const rows = settlement(game)

    return (
      <main className="app">
        {header}

        <section className="card">
          <h2>
            Year {game.round / 4}
            {' — '}
            Food Settlement
          </h2>

          <p>
            Each player requires
            {' '}
            {game.foodRequirement}
            {' '}
            food points.
            Each food point is worth
            {' '}
            {fmt(FOOD_PRICE)}.
          </p>

          <p className="muted">
            Negative cash balances are
            allowed. No bank loans
            are issued.
          </p>

          <div className="table-container">
            <table className="population-table">
              <thead>
                <tr>
                  <th>Player</th>

                  <th>Food</th>

                  <th>
                    Food transaction
                  </th>

                  <th>
                    Cash before
                  </th>

                  <th>
                    Cash after
                  </th>
                </tr>
              </thead>

              <tbody>
                {rows.map(row => (
                  <tr key={row.player}>
                    <td>
                      Player {row.player}
                    </td>

                    <td>
                      {row.food}
                    </td>

                    <td>
                      {fmt(
                        row.foodTransaction
                      )}
                    </td>

                    <td>
                      {fmt(
                        game.cash[
                          row.player - 1
                        ]
                      )}
                    </td>

                    <td>
                      {fmt(row.cash)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <button
            className="primary-button"
            onClick={confirmSettlement}
          >
            Confirm Food Settlement
          </button>

          <button
            className="secondary-button"
            onClick={previousRound}
          >
            Edit Last Round
          </button>
        </section>
      </main>
    )
  }

  // ==========================================================
  // ANNUAL REGENERATION
  // ==========================================================

  if (
    game.stage === 'regeneration'
  ) {
    const increase = growth(
      game.population,
      game.regenerationRate
    )

    const next = addPopulations(
      game.population,
      increase
    )

    const year =
      game.round / 4 + 1

    return (
      <main className="app">
        {header}

        <section className="card">
          <h2>
            End of Year {year - 1}
          </h2>

          <p>
            Annual regeneration:
            {' '}
            {game.regenerationRate}%
            {' '}
            (rounded down)
          </p>

          <h3>
            Animals to Add
          </h3>

          <PopulationTable
            population={game.population}
            increase={increase}
          />

          <h3>
            Beginning of Year {year}
          </h3>

          <PopulationTable
            population={next}
          />

          {failureInfo}

          <p className="muted">
            Failure tokens do not regenerate.
          </p>

          {monitor([
            ...game.history,
            {
              year,
              population: next,
            },
          ])}

          <button
            className="primary-button"
            onClick={confirmRegeneration}
          >
            {game.round === game.rounds
              ? 'Confirm and Finish Game'
              : 'Tokens Added — Next Year'}
          </button>

          <button
            className="secondary-button"
            onClick={previousRound}
          >
            Edit Last Round
          </button>
        </section>
      </main>
    )
  }

  // ==========================================================
  // FINISHED GAME SCREEN
  // ==========================================================

  if (
    game.stage === 'finished'
  ) {
    const negatives = replay(
      game,
      game.records,
      game.financialHistory
    ).everNegative

    return (
      <main className="app">
        {header}

        <section className="card">
          <h2>
            Game Completed
          </h2>

          {failureInfo}

          {monitor(game.history)}

          <h3>
            Final Finances
          </h3>

          <div className="table-container">
            <table className="population-table">
              <thead>
                <tr>
                  <th>Player</th>

                  <th>
                    Final cash
                  </th>

                  <th>
                    Ever negative?
                  </th>
                </tr>
              </thead>

              <tbody>
                {game.cash.map(
                  (balance, i) => (
                    <tr key={i}>
                      <td>
                        Player {i + 1}
                      </td>

                      <td>
                        {fmt(balance)}
                      </td>

                      <td>
                        {negatives[i]
                          ? 'Yes'
                          : 'No'}
                      </td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
          </div>
        </section>

        <GiftPanel
          game={game}
        />

        <section className="card">
          <h2>
            Round History
          </h2>

          {game.records.map(record => (
            <div
              className="player-card"
              key={record.round}
            >
              <h4>
                Round {record.round}
              </h4>

              {record.actions.map(
                (action, i) => (
                  <p key={i}>
                    Player {i + 1}:
                    {' '}
                    {choiceLabel(
                      action.choice
                    )}

                    {action.arrested
                      ? ' — Arrested'
                      : (
                        ` — Blue: ${action.blue}, ` +
                        `Red: ${action.red}, ` +
                        `Hog: ${action.hog}`
                      )}
                  </p>
                )
              )}
            </div>
          ))}

          <button
            className="secondary-button"
            onClick={previousRound}
          >
            Edit Last Round
          </button>
        </section>

        <ExportPanel
          game={game}
        />

        <section className="card">
          <button
            className="secondary-button"
            onClick={reset}
          >
            New Game
          </button>
        </section>
      </main>
    )
  }

  // ==========================================================
  // ROUND INPUT SCREEN
  // ==========================================================

  const remaining =
    game.actions.reduce(
      subtract,
      { ...game.population }
    )

  const ready =
    game.actions.every(
      action => !!action.choice
    )

  return (
    <main className="app">
      {header}

      <section className="card">
        <h2>
          Year {Math.ceil(game.round / 4)}
          {' — '}
          Round {game.round}/{game.rounds}
        </h2>

        <div className="settings">
          <p>
            <strong>
              Round in year:
            </strong>
            {' '}
            {((game.round - 1) % 4) + 1}/4
          </p>

          <p>
            <strong>
              Players:
            </strong>
            {' '}
            {game.players}
          </p>

          <p>
            <strong>
              Distant hunting cost:
            </strong>
            {' '}
            {fmt(game.travelCost)}
          </p>

          <p>
            <strong>
              Draws per hunt:
            </strong>
            {' '}
            {game.drawsPerHunt}
          </p>

          <p>
            <strong>
              No Hunting income:
            </strong>
            {' '}
            {fmt(NO_HUNTING_INCOME)}
          </p>

          {game.scenario === 'PAM' && (
            <p>
              <strong>
                PAM Fine:
              </strong>
              {' '}
              {fmt(game.fine)}
            </p>
          )}
        </div>

        {game.actions.map(
          (action, i) => (
            <div
              className="player-card"
              key={i}
            >
              <h3>
                Player {i + 1}
              </h3>

              <p className="muted">
                Current cash:
                {' '}
                {fmt(game.cash[i])}
              </p>

              <div className="choice-buttons">
                {(
                  [
                    [
                      'nearby',
                      'Nearby Hunting',
                    ],
                    [
                      'distant',
                      'Distant Hunting',
                    ],
                    [
                      'noHunting',
                      'No Hunting',
                    ],
                  ] as const
                ).map(
                  ([choice, label]) => (
                    <button
                      key={choice}
                      className={
                        action.choice ===
                        choice
                          ? 'selected'
                          : ''
                      }

                      disabled={
                        choice ===
                          'distant' &&
                        cannotAffordDistantHunting(
                          game.cash[i],
                          game.travelCost
                        )
                      }

                      onClick={() =>
                        choose(
                          i,
                          choice
                        )
                      }
                    >
                      {label}
                    </button>
                  )
                )}
              </div>

              {cannotAffordDistantHunting(
                game.cash[i],
                game.travelCost
              ) && (
                <p
                  className="muted"
                  style={{
                    color: '#b45309',
                    fontWeight: 600,
                  }}
                >
                  Insufficient cash for distant hunting.
                  <br />

                  Required:
                  {' '}
                  {fmt(game.travelCost)}
                  <br />

                  Available:
                  {' '}
                  {fmt(game.cash[i])}
                </p>
              )}

              {game.scenario === 'PAM' &&
                action.choice ===
                  'distant' && (
                  <div className="arrest-section">
                    <button
                      className={
                        `arrest-button ${
                          action.arrested
                            ? 'arrested'
                            : ''
                        }`
                      }

                      onClick={() =>
                        toggleArrest(i)
                      }
                    >
                      {action.arrested
                        ? '✓ Arrested'
                        : 'Arrested by Authorities?'}
                    </button>

                    {action.arrested && (
                      <p className="arrest-message">
                        Fine:
                        {' '}
                        {fmt(game.fine)}
                        <br />

                        No hunting — No catch
                      </p>
                    )}
                  </div>
                )}

              {(
                action.choice ===
                  'nearby' ||
                action.choice ===
                  'distant'
              ) &&
                !action.arrested && (
                  <div className="catch-section">
                    <h4>
                      Animals Caught
                    </h4>

                    {(
                      [
                        [
                          'blue',
                          'Blue duiker',
                        ],
                        [
                          'red',
                          'Red duiker',
                        ],
                        [
                          'hog',
                          'Red river hog',
                        ],
                      ] as const
                    )
                      .filter(([species]) =>
                        action.choice ===
                        'nearby'
                          ? species !== 'hog'
                          : species !== 'blue'
                      )
                      .map(
                        ([species, label]) => (
                          <div
                            className="animal-row"
                            key={species}
                          >
                            <span>
                              {label}
                            </span>

                            <div className="counter">
                              <button
                                onClick={() =>
                                  catchChange(
                                    i,
                                    species,
                                    -1
                                  )
                                }
                              >
                                −
                              </button>

                              <strong>
                                {action[species]}
                              </strong>

                              <button
                                disabled={
                                  caught(
                                    action
                                  ) >=
                                  game.drawsPerHunt
                                }

                                onClick={() =>
                                  catchChange(
                                    i,
                                    species,
                                    1
                                  )
                                }
                              >
                                +
                              </button>
                            </div>
                          </div>
                        )
                      )}

                    <p className="muted">
                      Total:
                      {' '}
                      {caught(action)}
                      /
                      {game.drawsPerHunt}
                    </p>
                  </div>
                )}

              <p className="muted">
                Cash change:
                {' '}
                {fmt(
                  actionCash(
                    action,
                    game
                  )
                )}
                <br />

                Food points:
                {' '}
                {foodPoints(action)}
              </p>
            </div>
          )
        )}

        <button
          className="primary-button"
          disabled={!ready}
          onClick={confirmRound}
        >
          {game.round % 4 === 0
            ? 'Confirm Round and Settle Year'
            : 'Confirm Round'}
        </button>

        {game.records.length > 0 && (
          <button
            className="secondary-button"
            onClick={previousRound}
          >
            Previous Round
          </button>
        )}

        <h3>
          Remaining Wildlife Populations
        </h3>

        <PopulationTable
          population={remaining}
        />

        {failureInfo}

        <h3>
          Current Cash
        </h3>

        {game.cash.map(
          (balance, i) => (
            <p key={i}>
              Player {i + 1}:
              {' '}
              {fmt(balance)}
            </p>
          )
        )}

        <button
          className="secondary-button"
          onClick={() =>
            setShowGraphs(
              current => !current
            )
          }
        >
          {showGraphs
            ? 'Hide Population Graphs'
            : 'View Population Graphs'}
        </button>

        {showGraphs &&
          monitor(game.history)}

        <p className="muted">
          Changes are saved automatically.
        </p>
      </section>

      <section className="card">
        <button
          className="secondary-button"
          onClick={reset}
        >
          Delete Game
        </button>
      </section>
    </main>
  )
}

export default App