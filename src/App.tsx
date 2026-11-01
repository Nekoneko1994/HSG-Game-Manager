import { useEffect, useState } from 'react'
import ExcelJS from 'exceljs'
import './App.css'

// ========================================
// Types
// ========================================

type Choice = '' | 'nearby' | 'distant' | 'agriculture'
type Scenario = 'OA' | 'PAM' | 'CBM'
type Species = 'blue' | 'red' | 'hog'
type Stage = 'round' | 'settlement' | 'regeneration' | 'finished'

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
  initialMoney: number
  rounds: number
  fine: number
  travelCost: number
  regenerationRate: number
  foodRequirement: number
  drawsPerHunt: number
  failureTokens: FailureTokens
  round: number
  actions: Action[]
  records: RoundRecord[]
  population: Population
  history: PopulationRecord[]
  cash: number[]
  financialHistory: FinancialRecord[]
  stage: Stage
}

// ========================================
// Constants
// ========================================

const STORAGE = 'hsg-game-manager-v4'

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

const TOKEN_OPTIONS = Array.from(
  { length: 7 },
  (_, i) => i * 5
)

const TENS = Array.from(
  { length: 11 },
  (_, i) => i * 10
)

const fmt = (n: number) =>
  `${n.toLocaleString('en-US')} FCFA`

const blank = (): Action => ({
  choice: '',
  blue: 0,
  red: 0,
  hog: 0,
  arrested: false,
})

const blanks = (n: number) =>
  Array.from({ length: n }, blank)

const caught = (a: Action) =>
  a.blue + a.red + a.hog

const foodPoints = (a: Action) =>
  10 * a.blue + 20 * a.red + 30 * a.hog

const actionCash = (a: Action, g: Game) =>
  a.choice === 'agriculture'
    ? 30000
    : a.choice === 'distant'
      ? -g.travelCost - (a.arrested ? g.fine : 0)
      : 0

const wildlifeRows = [
  { key: 'blue', name: 'Nearby — Blue duiker' },
  { key: 'nearRed', name: 'Nearby — Red duiker' },
  { key: 'farRed', name: 'Distant — Red duiker' },
  { key: 'hog', name: 'Distant — Red river hog' },
] as const

// ========================================
// Wildlife population calculations
// ========================================

function subtract(
  p: Population,
  a: Action
): Population {
  const q = { ...p }

  if (!a.arrested && a.choice === 'nearby') {
    q.blue -= a.blue
    q.nearRed -= a.red
  }

  if (!a.arrested && a.choice === 'distant') {
    q.farRed -= a.red
    q.hog -= a.hog
  }

  return q
}

function grow(
  p: Population,
  rate: number
): Population {
  const f = rate / 100

  return {
    blue: Math.floor(p.blue * f),
    nearRed: Math.floor(p.nearRed * f),
    farRed: Math.floor(p.farRed * f),
    hog: Math.floor(p.hog * f),
  }
}

function add(
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

// ========================================
// Annual food settlement
// No bank loans or repayments
// ========================================

function settlement(
  g: Game
): FinancialRecord[] {
  const year = g.round / 4

  return Array.from(
    { length: g.players },
    (_, i) => {
      const food = g.records
        .filter(
          r => Math.ceil(r.round / 4) === year
        )
        .reduce(
          (sum, r) =>
            sum + foodPoints(r.actions[i]),
          0
        )

      const foodTransaction =
        (food - g.foodRequirement) * 3000

      return {
        year,
        player: i + 1,
        food,
        foodTransaction,
        cash: g.cash[i] + foodTransaction,
      }
    }
  )
}

// ========================================
// Financial history reconstruction
//
// Negative balance at least once:
// ONE penalty only, regardless of frequency.
//
// The penalty is reconstructed from
// confirmed records, so editing the last
// round also updates the result.
// ========================================

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
    const group = settled.get(f.year) ?? []

    group.push(f)
    settled.set(f.year, group)
  }

  for (
    const record of [...records].sort(
      (a, b) => a.round - b.round
    )
  ) {
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
      for (
        const f of
        settled.get(record.round / 4) ?? []
      ) {
        const i = f.player - 1

        if (
          i >= 0 &&
          i < g.players
        ) {
          cash[i] += f.foodTransaction

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

// ========================================
// Final population before regeneration
// ========================================

function finalBeforeGrowth(
  g: Game
): Population {
  const year = g.rounds / 4

  const start = g.history.find(
    h => h.year === year
  )

  if (!start) {
    return { ...g.population }
  }

  let p = { ...start.population }

  for (
    const r of g.records.filter(
      r => Math.ceil(r.round / 4) === year
    )
  ) {
    for (const a of r.actions) {
      p = subtract(p, a)
    }
  }

  return p
}

// ========================================
// Gift rewards
// ========================================

function rewards(g: Game) {
  const negative = replay(
    g,
    g.records,
    g.financialHistory
  ).everNegative

  const finalPopulation = finalBeforeGrowth(g)

  const wildlife = Object.values(
    finalPopulation
  ).every(x => x > 0)

  const highest = Math.max(...g.cash)

  const rows = g.cash.map((cash, i) => {
    const participation = 1

    const wildlifeBonus =
      wildlife ? 1 : 0

    // Applied once, even if the player
    // had negative cash multiple times.
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
  })

  return {
    rows,
    wildlife,
    finalPopulation,
    total: rows.reduce(
      (s, r) => s + r.total,
      0
    ),
  }
}

// ========================================
// Excel export
// ========================================

async function exportExcel(g: Game) {
  try {
    const w = new ExcelJS.Workbook()

    w.creator = 'HSG Game Manager'

    // Sheet 1: Round Decisions

    const rounds = w.addWorksheet(
      'Round Decisions'
    )

    rounds.addRow([
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

    g.records.forEach(r =>
      r.actions.forEach((a, i) =>
        rounds.addRow([
          g.gameId,
          g.scenario,
          r.round,
          Math.ceil(r.round / 4),
          i + 1,
          a.choice,
          a.arrested,
          a.blue,
          a.red,
          a.hog,
          foodPoints(a),
          actionCash(a, g),
        ])
      )
    )

    // Sheet 2: Annual Finances

    const finances = w.addWorksheet(
      'Annual Finances'
    )

    finances.addRow([
      'Game ID',
      'Year',
      'Player',
      'Food points',
      'Food sale/purchase (FCFA)',
      'Year-end cash (FCFA)',
      'Negative year-end balance',
    ])

    g.financialHistory.forEach(r =>
      finances.addRow([
        g.gameId,
        r.year,
        r.player,
        r.food,
        r.foodTransaction,
        r.cash,
        r.cash < 0,
      ])
    )

    // Sheet 3: Wildlife Populations

    const populations = w.addWorksheet(
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

    g.history.forEach(r =>
      populations.addRow([
        g.gameId,
        r.year,
        r.population.blue,
        r.population.nearRed,
        r.population.farRed,
        r.population.hog,
      ])
    )

    // Sheet 4: Game Settings

    const settings = w.addWorksheet(
      'Game Settings'
    )

    settings.addRow([
      'Parameter',
      'Value',
    ])

    const p = g.history[0].population

    const settingRows: [
      string,
      string | number
    ][] = [
      ['Game ID', g.gameId],
      ['Scenario', g.scenario],
      ['Number of players', g.players],
      ['Initial cash (FCFA)', g.initialMoney],
      ['Number of rounds', g.rounds],
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
        'Food requirement (points/player/year)',
        g.foodRequirement,
      ],
      ['Draws per hunt', g.drawsPerHunt],
      ['Nearby blue duiker', p.blue],
      ['Nearby red duiker', p.nearRed],
      [
        'Nearby failure tokens',
        g.failureTokens.nearby,
      ],
      ['Distant red duiker', p.farRed],
      ['Distant red river hog', p.hog],
      [
        'Distant failure tokens',
        g.failureTokens.distant,
      ],
      [
        'Part-time job income (FCFA)',
        30000,
      ],
      [
        'Food price per point (FCFA)',
        3000,
      ],
      [
        'Negative balance rule',
        'No bank loans; negative cash allowed',
      ],
      [
        'Gift penalty',
        'Minus one if cash was negative at least once',
      ],
    ]

    settingRows.forEach(row =>
      settings.addRow(row)
    )

    // Sheet 5: Gift Rewards

    const gifts = w.addWorksheet(
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

    summary.rows.forEach(r =>
      gifts.addRow([
        g.gameId,
        r.player,
        r.participation,
        r.wildlifeBonus,
        r.negativePenalty,
        r.highestBonus,
        r.total,
        r.everNegative,
        r.finalCash,
        summary.wildlife,
      ])
    )

    gifts.addRow([])

    gifts.addRow([
      'Final population before regeneration',
    ])

    wildlifeRows.forEach(r =>
      gifts.addRow([
        r.name,
        summary.finalPopulation[r.key],
      ])
    )

    gifts.addRow([
      'Total gifts required',
      summary.total,
    ])

    // Formatting

    w.worksheets.forEach(s => {
      const h = s.getRow(1)

      h.font = {
        bold: true,
        color: {
          argb: 'FFFFFFFF',
        },
      }

      h.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: 'FF176447',
        },
      }

      h.alignment = {
        vertical: 'middle',
        wrapText: true,
      }

      h.height = 40

      s.views = [
        {
          state: 'frozen',
          ySplit: 1,
        },
      ]

      s.columns.forEach(c => {
        c.width = 23
      })

      if (s.name === 'Game Settings') {
        s.getColumn(1).width = 46
      }

      s.eachRow((r, number) => {
        if (
          number > 1 &&
          number % 2 === 0
        ) {
          r.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'FFF2F7F3',
            },
          }
        }
      })

      if (s.name !== 'Game Settings') {
        s.autoFilter = {
          from: {
            row: 1,
            column: 1,
          },
          to: {
            row: 1,
            column: s.columnCount,
          },
        }
      }
    })

    const data = await w.xlsx.writeBuffer()

    const copy = new Uint8Array(data)

    const buffer = new ArrayBuffer(
      copy.byteLength
    )

    new Uint8Array(buffer).set(copy)

    const blob = new Blob(
      [buffer],
      {
        type:
          'application/vnd.openxmlformats-officedocument.' +
          'spreadsheetml.sheet',
      }
    )

    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url

    link.download = `${
      g.gameId.replace(
        /[^a-zA-Z0-9_-]/g,
        '_'
      ) || 'HSG'
    }.xlsx`

    document.body.appendChild(link)
    link.click()
    link.remove()

    setTimeout(
      () => URL.revokeObjectURL(url),
      1000
    )
  } catch (e) {
    console.error(e)

    window.alert(
      'Excel export failed.'
    )
  }
}

// ========================================
// Population table
// ========================================

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
            {increase && <th>To add</th>}
          </tr>
        </thead>

        <tbody>
          {wildlifeRows.map(a => (
            <tr key={a.key}>
              <td>{a.name}</td>

              <td>
                {population[a.key]}
              </td>

              {increase && (
                <td>
                  +{increase[a.key]}
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ========================================
// Population graph
// ========================================

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
    ...history.flatMap(r =>
      series.map(s =>
        r.population[s.key]
      )
    )
  )

  const top =
    Math.ceil(maximum / 10) * 10

  const count = Math.max(
    2,
    years + 1
  )

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
        {[0, 0.25, 0.5, 0.75, 1].map(v => (
          <g key={v}>
            <line
              x1="42"
              x2="332"
              y1={y(v * top)}
              y2={y(v * top)}
              stroke="#e5e7eb"
            />

            <text
              x="36"
              y={y(v * top) + 4}
              textAnchor="end"
              fontSize="11"
            >
              {Math.round(v * top)}
            </text>
          </g>
        ))}

        {Array.from(
          { length: count },
          (_, i) => i + 1
        ).map(yr => (
          <text
            key={yr}
            x={x(yr)}
            y="210"
            textAnchor="middle"
            fontSize="11"
          >
            {yr}
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
                  r =>
                    `${x(r.year)},${y(
                      r.population[s.key]
                    )}`
                )
                .join(' ')}
            />

            {history.map(r => (
              <circle
                key={r.year}
                cx={x(r.year)}
                cy={y(
                  r.population[s.key]
                )}
                r="4"
                fill={s.color}
              >
                <title>
                  {s.name}: Year {r.year}
                  {' — '}
                  {r.population[s.key]}
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
        Populations at the beginning of each year.
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

// ========================================
// Excel export panel
// ========================================

function ExportPanel({
  game,
}: {
  game: Game
}) {
  const [busy, setBusy] = useState(false)

  return (
    <section className="card">
      <h2>Export Research Data</h2>

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

// ========================================
// Gift rewards panel
// ========================================

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
        Wildlife evaluated before final regeneration.
        Negative cash incurs a one-time penalty,
        even if subsequently recovered.
      </p>

      <PopulationTable
        population={summary.finalPopulation}
      />

      <h3>Rewards by Player</h3>

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
            {summary.rows.map(r => (
              <tr key={r.player}>
                <td>
                  Player {r.player}
                </td>

                <td>
                  +{r.participation}
                </td>

                <td>
                  +{r.wildlifeBonus}
                </td>

                <td>
                  {r.negativePenalty}
                </td>

                <td>
                  +{r.highestBonus}
                </td>

                <td>
                  <strong>
                    {r.total}
                  </strong>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h3>
        Total Gifts Required: {summary.total}
      </h3>
    </section>
  )
}

// ========================================
// Main application
// ========================================

function App() {
  // New-game settings

  const [id, setId] = useState('HSG-001')

  const [scenario, setScenario] =
    useState<Scenario>('OA')

  const [players, setPlayers] = useState(4)

  const [initialMoney, setInitialMoney] =
    useState(0)

  const [rounds, setRounds] = useState(20)

  const [fine, setFine] = useState(50000)

  const [travelCost, setTravelCost] =
    useState(0)

  const [regenerationRate, setRegenerationRate] =
    useState(50)

  const [foodRequirement, setFoodRequirement] =
    useState(50)

  const [drawsPerHunt, setDrawsPerHunt] =
    useState(3)

  const [initialPop, setInitialPop] =
    useState<Population>({
      ...DEFAULT_POPULATION,
    })

  const [failureTokens, setFailureTokens] =
    useState<FailureTokens>({
      ...DEFAULT_FAILURE,
    })

  const [showGraphs, setShowGraphs] =
    useState(false)

  // v4 is kept separate from old v3 data.

  const [game, setGame] =
    useState<Game | null>(() => {
      try {
        const text =
          localStorage.getItem(STORAGE)

        if (!text) return null

        return JSON.parse(text) as Game
      } catch {
        return null
      }
    })

  // Autosave

  useEffect(() => {
    if (game) {
      localStorage.setItem(
        STORAGE,
        JSON.stringify(game)
      )
    } else {
      localStorage.removeItem(STORAGE)
    }
  }, [game])

  // Reset new-game settings

  function resetSettings() {
    setId('HSG-001')
    setScenario('OA')
    setPlayers(4)
    setInitialMoney(0)
    setRounds(20)
    setFine(50000)
    setTravelCost(0)
    setRegenerationRate(50)
    setFoodRequirement(50)
    setDrawsPerHunt(3)

    setInitialPop({
      ...DEFAULT_POPULATION,
    })

    setFailureTokens({
      ...DEFAULT_FAILURE,
    })

    setShowGraphs(false)
  }

  // Start game

  function start() {
    if (
      !id.trim() ||
      !Number.isInteger(players) ||
      players < 2 ||
      players > 10 ||
      initialMoney < 0 ||
      initialMoney > 100000 ||
      initialMoney % 10000 !== 0 ||
      travelCost < 0 ||
      travelCost > 50000 ||
      travelCost % 10000 !== 0 ||
      rounds < 4 ||
      rounds > 40 ||
      rounds % 4 !== 0 ||
      !TENS.includes(regenerationRate) ||
      !TENS.includes(foodRequirement) ||
      ![1, 2, 3, 4, 5].includes(drawsPerHunt) ||
      Object.values(initialPop).some(
        n => !TOKEN_OPTIONS.includes(n)
      ) ||
      Object.values(failureTokens).some(
        n => !TOKEN_OPTIONS.includes(n)
      )
    ) {
      window.alert(
        'Please check the game settings.'
      )
      return
    }

    setGame({
      gameId: id.trim(),
      scenario,
      players,
      initialMoney,
      rounds,
      fine,
      travelCost,
      regenerationRate,
      foodRequirement,
      drawsPerHunt,

      failureTokens: {
        ...failureTokens,
      },

      round: 1,
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

      cash: Array(players).fill(initialMoney),
      financialHistory: [],
      stage: 'round',
    })
  }

  // Change player choice

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
      game.cash[i] < game.travelCost
    ) {
      return
    }

    setGame(g =>
      g && g.stage === 'round'
        ? {
            ...g,

            actions: g.actions.map(
              (a, j) =>
                i === j
                  ? {
                      ...blank(),
                      choice,
                    }
                  : a
            ),
          }
        : g
    )
  }

  // Update catch

  function catchChange(
    i: number,
    species: Species,
    delta: number
  ) {
    setGame(g => {
      if (
        !g ||
        g.stage !== 'round'
      ) {
        return g
      }

      const a = g.actions[i]

      if (
        a.arrested ||
        (
          a.choice !== 'nearby' &&
          a.choice !== 'distant'
        ) ||
        (
          a.choice === 'nearby' &&
          species === 'hog'
        ) ||
        (
          a.choice === 'distant' &&
          species === 'blue'
        ) ||
        (
          delta > 0 &&
          caught(a) >= g.drawsPerHunt
        ) ||
        (
          delta < 0 &&
          a[species] <= 0
        )
      ) {
        return g
      }

      const actions = g.actions.map(
        (item, j) =>
          j === i
            ? {
                ...item,
                [species]:
                  item[species] + delta,
              }
            : item
      )

      const remain = actions.reduce(
        subtract,
        { ...g.population }
      )

      if (
        Object.values(remain).some(
          x => x < 0
        )
      ) {
        return g
      }

      return {
        ...g,
        actions,
      }
    })
  }

  // PAM arrest

  function toggleArrest(i: number) {
    setGame(g =>
      g &&
      g.stage === 'round' &&
      g.scenario === 'PAM'
        ? {
            ...g,

            actions: g.actions.map(
              (a, j) =>
                i === j &&
                a.choice === 'distant'
                  ? {
                      ...a,
                      arrested: !a.arrested,
                      blue: 0,
                      red: 0,
                      hog: 0,
                    }
                  : a
            ),
          }
        : g
    )
  }

  // Confirm round

  function confirmRound() {
    if (
      !game ||
      game.stage !== 'round' ||
      game.actions.some(a => !a.choice)
    ) {
      return
    }

    for (
      let i = 0;
      i < game.players;
      i++
    ) {
      if (
        game.actions[i].choice === 'distant' &&
        game.cash[i] < game.travelCost
      ) {
        window.alert(
          `Player ${i + 1} cannot afford distant hunting.`
        )
        return
      }

      if (
        caught(game.actions[i]) >
        game.drawsPerHunt
      ) {
        return
      }
    }

    const remaining = game.actions.reduce(
      subtract,
      { ...game.population }
    )

    if (
      Object.values(remaining).some(
        x => x < 0
      )
    ) {
      window.alert(
        'Catches exceed the remaining population.'
      )
      return
    }

    const cash = game.cash.map(
      (v, i) =>
        v + actionCash(game.actions[i], game)
    )

    const records = [
      ...game.records,

      {
        round: game.round,
        actions: game.actions.map(
          a => ({ ...a })
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

      stage:
        yearEnd
          ? 'settlement'
          : 'round',

      round:
        yearEnd
          ? game.round
          : game.round + 1,

      actions: blanks(game.players),
    })

    window.scrollTo(0, 0)
  }

  // Confirm food settlement

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

      cash: rows.map(r => r.cash),

      financialHistory: [
        ...game.financialHistory,
        ...rows,
      ],

      stage: 'regeneration',
    })

    window.scrollTo(0, 0)
  }

  // Confirm annual regeneration

  function confirmRegeneration() {
    if (
      !game ||
      game.stage !== 'regeneration'
    ) {
      return
    }

    const p = add(
      game.population,

      grow(
        game.population,
        game.regenerationRate
      )
    )

    const finished =
      game.round === game.rounds

    setGame({
      ...game,

      population: p,

      history: [
        ...game.history,

        {
          year:
            game.round / 4 + 1,
          population: p,
        },
      ],

      stage:
        finished
          ? 'finished'
          : 'round',

      round:
        finished
          ? game.round
          : game.round + 1,

      actions: blanks(game.players),
    })

    window.scrollTo(0, 0)
  }

  // Edit previous round

  function previousRound() {
    if (
      !game ||
      !game.records.length ||
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

    const history = game.history.filter(
      h => h.year <= year
    )

    const initial = history.find(
      h => h.year === year
    )?.population

    if (!initial) {
      return
    }

    let p = { ...initial }

    const records =
      game.records.slice(0, -1)

    for (
      const r of records.filter(
        r =>
          Math.ceil(r.round / 4) === year
      )
    ) {
      for (const a of r.actions) {
        p = subtract(p, a)
      }
    }

    const financialHistory =
      game.financialHistory.filter(
        f => f.year < year
      )

    const { cash } = replay(
      game,
      records,
      financialHistory
    )

    setGame({
      ...game,

      round: last.round,

      actions: last.actions.map(
        a => ({ ...a })
      ),

      records,
      population: p,
      history,
      cash,
      financialHistory,
      stage: 'round',
    })

    window.scrollTo(0, 0)
  }

  // Delete game

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

  // Dropdown helper

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
        onChange={e =>
          change(Number(e.target.value))
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

  // ======================================
  // New Game Screen
  // ======================================

  if (!game) {
    return (
      <main className="app">
        <header className="header">
          <h1>Hunting Serious Game</h1>

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
            onChange={e =>
              setId(e.target.value)
            }
          />

          <label htmlFor="scenario">
            Scenario
          </label>

          <select
            id="scenario"
            value={scenario}
            onChange={e =>
              setScenario(
                e.target.value as Scenario
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
              [
                25000,
                50000,
                100000,
                200000,
                300000,
              ],
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
            onChange={e =>
              setPlayers(
                Number(e.target.value)
              )
            }
          />

          {select(
            'cash',
            'Initial cash (FCFA)',
            initialMoney,
            setInitialMoney,
            Array.from(
              { length: 11 },
              (_, i) => i * 10000
            ),
            ' FCFA'
          )}

          {select(
            'travel',
            'Distant hunting cost (FCFA)',
            travelCost,
            setTravelCost,
            Array.from(
              { length: 6 },
              (_, i) => i * 10000
            ),
            ' FCFA'
          )}

          {select(
            'rounds',
            'Number of rounds',
            rounds,
            setRounds,
            Array.from(
              { length: 10 },
              (_, i) => (i + 1) * 4
            ),
            ' rounds'
          )}

          <h3>Game Rules</h3>

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
            [1, 2, 3, 4, 5],
            ' draws'
          )}

          <h3>
            Initial Wildlife Populations
            &amp; Failure Tokens
          </h3>

          <p className="muted">
            Select the initial number of
            each token type.
          </p>

          <h4>Nearby Area</h4>

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

          <h4>Distant Area</h4>

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
              <strong>Food price:</strong>
              {' '}3,000 FCFA per point
            </p>

            <p>
              <strong>Part-time job:</strong>
              {' '}30,000 FCFA per round
            </p>

            <p>
              <strong>Negative balance:</strong>
              {' '}Allowed; no bank loans
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

  // ======================================
  // Shared UI
  // ======================================

  const header = (
    <header className="header">
      <h1>HSG Game Manager</h1>

      <p>
        Game: {game.gameId} — {game.scenario}
      </p>
    </header>
  )

  const monitor = (
    h: PopulationRecord[]
  ) => (
    <Monitoring
      history={h}
      years={game.rounds / 4}
    />
  )

  const failureInfo = (
    <div className="settings">
      <p>
        <strong>
          Failure tokens — Nearby:
        </strong>
        {' '}{game.failureTokens.nearby}
      </p>

      <p>
        <strong>
          Failure tokens — Distant:
        </strong>
        {' '}{game.failureTokens.distant}
      </p>
    </div>
  )

  // ======================================
  // Annual Food Settlement
  // ======================================

  if (game.stage === 'settlement') {
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
            {' '}{game.foodRequirement}
            {' '}food points.
            Each food point is worth
            {' '}3,000 FCFA.
            Negative cash balances are allowed;
            no loans are issued.
          </p>

          <div className="table-container">
            <table className="population-table">
              <thead>
                <tr>
                  <th>Player</th>
                  <th>Food</th>
                  <th>Food transaction</th>
                  <th>Cash before</th>
                  <th>Cash after</th>
                </tr>
              </thead>

              <tbody>
                {rows.map(r => (
                  <tr key={r.player}>
                    <td>
                      Player {r.player}
                    </td>

                    <td>{r.food}</td>

                    <td>
                      {fmt(r.foodTransaction)}
                    </td>

                    <td>
                      {fmt(
                        game.cash[r.player - 1]
                      )}
                    </td>

                    <td>
                      {fmt(r.cash)}
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

  // ======================================
  // Annual Regeneration
  // ======================================

  if (game.stage === 'regeneration') {
    const increase = grow(
      game.population,
      game.regenerationRate
    )

    const next = add(
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
            {' '}{game.regenerationRate}%
            {' '}(rounded down)
          </p>

          <h3>Animals to Add</h3>

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

  // ======================================
  // Finished Game Screen
  // ======================================

  if (game.stage === 'finished') {
    const negatives = replay(
      game,
      game.records,
      game.financialHistory
    ).everNegative

    return (
      <main className="app">
        {header}

        <section className="card">
          <h2>Game Completed</h2>

          {failureInfo}

          {monitor(game.history)}

          <h3>Final Finances</h3>

          <div className="table-container">
            <table className="population-table">
              <thead>
                <tr>
                  <th>Player</th>
                  <th>Final cash</th>
                  <th>Ever negative?</th>
                </tr>
              </thead>

              <tbody>
                {game.cash.map((n, i) => (
                  <tr key={i}>
                    <td>
                      Player {i + 1}
                    </td>

                    <td>
                      {fmt(n)}
                    </td>

                    <td>
                      {negatives[i]
                        ? 'Yes'
                        : 'No'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <GiftPanel game={game} />

        <section className="card">
          <h2>Round History</h2>

          {game.records.map(r => (
            <div
              className="player-card"
              key={r.round}
            >
              <h4>
                Round {r.round}
              </h4>

              {r.actions.map((a, i) => (
                <p key={i}>
                  Player {i + 1}:
                  {' '}{a.choice}

                  {a.arrested
                    ? ' — Arrested'
                    : (
                      ` — Blue: ${a.blue}, ` +
                      `Red: ${a.red}, ` +
                      `Hog: ${a.hog}`
                    )}
                </p>
              ))}
            </div>
          ))}

          <button
            className="secondary-button"
            onClick={previousRound}
          >
            Edit Last Round
          </button>
        </section>

        <ExportPanel game={game} />

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

  // ======================================
  // Round Input Screen
  // ======================================

  const remaining = game.actions.reduce(
    subtract,
    { ...game.population }
  )

  const ready = game.actions.every(
    a => !!a.choice
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
            <strong>Players:</strong>
            {' '}{game.players}
          </p>

          <p>
            <strong>
              Distant hunting cost:
            </strong>
            {' '}{fmt(game.travelCost)}
          </p>

          <p>
            <strong>
              Draws per hunt:
            </strong>
            {' '}{game.drawsPerHunt}
          </p>

          {game.scenario === 'PAM' && (
            <p>
              <strong>PAM Fine:</strong>
              {' '}{fmt(game.fine)}
            </p>
          )}
        </div>

        {game.actions.map((a, i) => (
          <div
            className="player-card"
            key={i}
          >
            <h3>
              Player {i + 1}
            </h3>

            <p className="muted">
              Current cash:
              {' '}{fmt(game.cash[i])}
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
                    'agriculture',
                    'Part-time Job',
                  ],
                ] as const
              ).map(([choice, label]) => (
                <button
                  key={choice}
                  className={
                    a.choice === choice
                      ? 'selected'
                      : ''
                  }
                  disabled={
                    choice === 'distant' &&
                    game.cash[i] < game.travelCost
                  }
                  onClick={() =>
                    choose(i, choice)
                  }
                >
                  {label}
                </button>
              ))}
            </div>

            {game.cash[i] < game.travelCost && (
              <p
                className="muted"
                style={{
                  color: '#b45309',
                  fontWeight: 600,
                }}
              >
                Insufficient cash for distant hunting.
                <br />
                Required: {fmt(game.travelCost)}
                <br />
                Available: {fmt(game.cash[i])}
              </p>
            )}

            {game.scenario === 'PAM' &&
              a.choice === 'distant' && (
                <div className="arrest-section">
                  <button
                    className={
                      `arrest-button ${
                        a.arrested
                          ? 'arrested'
                          : ''
                      }`
                    }
                    onClick={() =>
                      toggleArrest(i)
                    }
                  >
                    {a.arrested
                      ? '✓ Arrested'
                      : 'Arrested by Authorities?'}
                  </button>

                  {a.arrested && (
                    <p className="arrest-message">
                      Fine: {fmt(game.fine)}
                      <br />
                      No hunting — No catch
                    </p>
                  )}
                </div>
              )}

            {(a.choice === 'nearby' ||
              a.choice === 'distant') &&
              !a.arrested && (
                <div className="catch-section">
                  <h4>Animals Caught</h4>

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
                    .filter(([s]) =>
                      a.choice === 'nearby'
                        ? s !== 'hog'
                        : s !== 'blue'
                    )
                    .map(([s, label]) => (
                      <div
                        className="animal-row"
                        key={s}
                      >
                        <span>
                          {label}
                        </span>

                        <div className="counter">
                          <button
                            onClick={() =>
                              catchChange(
                                i,
                                s,
                                -1
                              )
                            }
                          >
                            −
                          </button>

                          <strong>
                            {a[s]}
                          </strong>

                          <button
                            disabled={
                              caught(a) >=
                              game.drawsPerHunt
                            }
                            onClick={() =>
                              catchChange(
                                i,
                                s,
                                1
                              )
                            }
                          >
                            +
                          </button>
                        </div>
                      </div>
                    ))}

                  <p className="muted">
                    Total:
                    {' '}{caught(a)}
                    /{game.drawsPerHunt}
                  </p>
                </div>
              )}

            <p className="muted">
              Cash change:
              {' '}
              {fmt(actionCash(a, game))}
              <br />
              Food points:
              {' '}{foodPoints(a)}
            </p>
          </div>
        ))}

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

        <h3>Current Cash</h3>

        {game.cash.map((v, i) => (
          <p key={i}>
            Player {i + 1}:
            {' '}{fmt(v)}
          </p>
        ))}

        <button
          className="secondary-button"
          onClick={() =>
            setShowGraphs(v => !v)
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