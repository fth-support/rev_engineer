import { useEffect, useLayoutEffect, useRef, useState, useCallback } from 'react'
import { Link } from 'react-router-dom'
import {
  Search, ShieldCheck, BarChart3, Ruler, Hammer, FlaskConical,
  Brain, BookOpen, UserCheck, Bot, Sparkles, Play, Pause,
  Network, Download,
} from 'lucide-react'
import StatCounter from '../components/primitives/StatCounter'
import './methodology.css'

const prefersReduced =
  typeof window !== 'undefined' &&
  window.matchMedia &&
  window.matchMedia('(prefers-reduced-motion: reduce)').matches

/* ------------------------------------------------------------------ data --- */
// Six phases of Evidence-Driven Modernization, mapped to the AI-platform lifecycle.
const PHASES = [
  {
    key: 'discover', n: '01', phase: 'PLAN', map: 'Planning', name: 'Discover',
    tone: 'var(--role-branch)', toneSoft: 'var(--role-branch-soft)', Icon: Search,
    agent: 'Code Archaeologist', to: '/flows',
    features: [{ t: 'Source reading', s: 'p' }, { t: 'Dependency map', s: 'p' }, { t: 'Structure extraction', s: 'p' }],
    tools: 'Grep · Glob · Read · Git', llm: 'Claude Opus 4.x',
  },
  {
    key: 'verify', n: '02', phase: 'ASSET', map: 'Asset Visualization', name: 'Verify',
    tone: 'var(--role-central)', toneSoft: 'var(--role-central-soft)', Icon: ShieldCheck,
    agent: 'Verification Agent', to: '/docs/data-dictionary',
    features: [{ t: 'Fact-tracing', s: 'p' }, { t: 'Error detection', s: 'p' }, { t: 'No-hallucination', s: 'p' }],
    tools: 'cross-ref · diff · trace', llm: 'Claude Opus 4.x',
  },
  {
    key: 'visualize', n: '03', phase: 'DESIGN', map: 'Requirements / Design', name: 'Visualize',
    tone: 'var(--role-member)', toneSoft: 'var(--role-member-soft)', Icon: BarChart3,
    agent: 'Docs / Frontend Agent', to: '/flows',
    features: [{ t: 'Interactive portal', s: 'p' }, { t: 'Live ER / flow', s: 'p' }, { t: 'Theme / UX', s: 'p' }],
    tools: 'React · Vite · SVG', llm: 'Claude Opus 4.x',
  },
  {
    key: 'specify', n: '04', phase: 'SPEC', map: 'Design hand-off', name: 'Specify',
    tone: 'var(--role-safenet)', toneSoft: 'var(--role-safenet-soft)', Icon: Ruler,
    agent: 'System Analyst', to: '/docs/program-spec',
    features: [{ t: 'Any-language brief', s: 'p' }, { t: 'Pseudocode + rules', s: 'p' }, { t: 'Bilingual docs', s: 'p' }],
    tools: 'Markdown · docx/pptx', llm: 'Claude Opus 4.x',
  },
  {
    key: 'modernize', n: '05', phase: 'BUILD', map: 'Implementation', name: 'Modernize',
    tone: 'var(--role-purple)', toneSoft: 'var(--role-purple-soft)', Icon: Hammer,
    agent: 'Solution Architect', to: '/docs/architecture-srs',
    features: [{ t: 'Keep-rules + new arch', s: 'p' }, { t: 'Target architecture', s: 'p' }, { t: 'Code generation', s: 'a' }],
    tools: 'Go · .NET · Java · Node', llm: 'Claude Opus 4.x', llmOpt: '＋ Copilot · GPT',
  },
  {
    key: 'handoff', n: '06', phase: 'VERIFY', map: 'Test / Operate', name: 'Handoff',
    tone: 'var(--accent)', toneSoft: 'var(--accent-soft)', Icon: FlaskConical,
    agent: 'QA / Test Agent', to: '/downloads',
    features: [{ t: 'Acceptance criteria', s: 'p' }, { t: 'CI/CD + deploy', s: 'p' }, { t: 'Auto-test suite', s: 'a' }],
    tools: 'GitHub Actions · tests', llm: 'Claude Opus 4.x',
  },
]

const BANDS = [
  { Icon: Brain, label: 'AI Orchestration', title: 'Claude Code — one orchestrator', text: 'drives all six steps end-to-end, keeping a single chain of traceability (no multi-agent hand-off gaps).', spine: true },
  { Icon: BookOpen, label: 'Knowledge', title: 'The source repository itself', text: 'legacy code + verified facts are the knowledge base. No external RAG / knowledge graph that can drift from reality.' },
  { Icon: UserCheck, label: 'Human-in-loop', title: 'You stay in control', text: 'scope, steer and approve at each gate; the agent does the heavy lifting. Review is targeted, never the bottleneck.' },
]

const KPIS = [
  { value: 2860, suffix: '', label: 'VB6 lines verified' },
  { value: 100, suffix: '%', label: 'traced to source' },
  { value: 15, suffix: '+', label: 'docs delivered' },
  { value: 1, suffix: '', label: 'any-language brief' },
]

const WINS = [
  { label: 'Evidence, not AI-guessing', tone: 'var(--role-branch)' },
  { label: 'Portable — no vendor lock-in', tone: 'var(--role-central)' },
  { label: 'Interactive deliverables', tone: 'var(--role-member)' },
  { label: 'Lean — one agent, open tools', tone: 'var(--role-purple)' },
]

const LEGEND = [
  { s: 'p', label: 'Proven (ServiceTransfer)' },
  { s: 'a', label: 'Available' },
  { s: 'r', label: 'Roadmap' },
]

const COLS = PHASES.length
const colCenter = (i) => ((i + 0.5) / COLS) * 100

/* ------------------------------------------------------------------ page --- */

function Methodology() {
  const [active, setActive] = useState(0)
  const [playing, setPlaying] = useState(!prefersReduced)
  const [mounted, setMounted] = useState(false)
  const pausedRef = useRef(false)
  const fitRef = useRef(null)

  // staggered reveal on mount
  useEffect(() => {
    const t = setTimeout(() => setMounted(true), 60)
    return () => clearTimeout(t)
  }, [])

  // Slide-like fit: scale the matrix so the whole thing always fits one screen
  // (the user's intent — see the entire flow at a glance, no scroll, no "next").
  // transform never changes scrollHeight, so re-measuring stays stable.
  useLayoutEffect(() => {
    const el = fitRef.current
    if (!el) return
    const stage = el.parentElement
    const measure = () => {
      if (window.innerWidth < 720) { el.style.transform = 'none'; return }
      const s = Math.min(1, stage.clientHeight / el.scrollHeight)
      el.style.transform = `scale(${s})`
    }
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(stage)
    window.addEventListener('resize', measure)
    const t = setTimeout(measure, 260) // re-measure after webfont settles
    return () => { ro.disconnect(); window.removeEventListener('resize', measure); clearTimeout(t) }
  }, [])

  // auto-sweeping spotlight (the "storytelling" loop) — pauses on hover
  useEffect(() => {
    if (!playing || prefersReduced) return
    const id = setInterval(() => {
      if (!pausedRef.current) setActive((a) => (a + 1) % COLS)
    }, 2600)
    return () => clearInterval(id)
  }, [playing])

  const focus = useCallback((i) => { pausedRef.current = true; setActive(i) }, [])
  const release = useCallback(() => { pausedRef.current = false }, [])

  const colClass = (i) => `cap-col ${active === i ? 'is-active' : ''} ${active !== i ? 'is-dim' : ''}`
  const inClass = mounted ? 'is-in' : ''

  return (
    <div className={`cap ${inClass}`} style={{ '--active-c': PHASES[active].tone }}>
     <div className="cap-stage">
      <div className="cap-fit" ref={fitRef}>
      {/* glow that follows the active phase */}
      <div className="cap-spotlight" aria-hidden="true" style={{ left: `${colCenter(active)}%` }} />

      {/* ---- header ---- */}
      <header className="cap-head">
        <div className="cap-head__main">
          <span className="cap-kicker" style={{ '--r': 0 }}>Modernization Methodology · Capability Map</span>
          <h1 className="cap-title" style={{ '--r': 1 }}>Evidence-Driven Modernization</h1>
          <p className="cap-intro" style={{ '--r': 2 }}>
            The same end-to-end lifecycle as an AI modernization platform — planning, asset visualization,
            design, build and test — driven by <strong>one AI orchestrator</strong> with specialized agents.
            The difference: every step is <strong>traced back to the real source code</strong>, producing a
            build-ready directive for <strong>any language</strong>.
          </p>
        </div>
        <div className="cap-head__badge" style={{ '--r': 1 }}>
          <span className="cap-badge"><Sparkles size={13} /> WE DON&apos;T GUESS — WE VERIFY</span>
          <span className="cap-badge__sub">Proven end-to-end on a live legacy system</span>
        </div>
      </header>

      {/* ---- the matrix ---- */}
      <div className="cap-matrix" onMouseLeave={release}>

        {/* phase header row */}
        <div className="cap-row cap-row--phases" style={{ '--r': 3 }}>
          <div className="cap-rowlab"><span>PHASE →</span><small>vs platform lifecycle</small></div>
          <div className="cap-cells">
            {PHASES.map((p, i) => (
              <button key={p.key} className={`${colClass(i)} cap-phase`} style={{ '--c': p.tone, '--cs': p.toneSoft }}
                onMouseEnter={() => focus(i)} onFocus={() => focus(i)}>
                <span className="cap-phase__n">{p.n} · {p.phase}</span>
                <span className="cap-phase__name">{p.name}</span>
                <span className="cap-phase__map">＝ {p.map}</span>
              </button>
            ))}
          </div>
        </div>

        {/* orchestration band (the animated spine origin) */}
        <div className="cap-row cap-row--band" style={{ '--r': 4 }}>
          <div className="cap-rowlab"><span>AI Orchestration</span><small>coordinator</small></div>
          <div className="cap-band cap-band--spine">
            <span className="cap-band__icon"><Brain size={17} /></span>
            <b>Claude Code agent</b>
            <span className="cap-band__txt">one orchestrator drives all six steps — a single, unbroken chain of traceability.</span>
            <span className="cap-arrow" aria-hidden="true">
              <span className="cap-arrow__flow" /><span className="cap-arrow__flow cap-arrow__flow--2" />
            </span>
          </div>
        </div>

        {/* AI agents row — nodes sit on the spine, glow lands on the active one */}
        <div className="cap-row cap-row--agents" style={{ '--r': 5 }}>
          <div className="cap-rowlab"><span>AI Agents</span><small>specialized role</small></div>
          <div className="cap-cells cap-cells--agents">
            <span className="cap-railline" aria-hidden="true" />
            <span className="cap-railpulse" aria-hidden="true" />
            {PHASES.map((p, i) => {
              const { Icon } = p
              return (
                <Link key={p.key} to={p.to} className={`${colClass(i)} cap-agent`} style={{ '--c': p.tone, '--cs': p.toneSoft }}
                  onMouseEnter={() => focus(i)} onFocus={() => focus(i)} title={`Open ${p.name} →`}>
                  <span className="cap-agent__dot"><Icon size={19} /></span>
                  <span className="cap-agent__name">{p.agent}</span>
                </Link>
              )
            })}
          </div>
        </div>

        {/* features row */}
        <div className="cap-row" style={{ '--r': 6 }}>
          <div className="cap-rowlab"><span>Features</span><small>capabilities</small></div>
          <div className="cap-cells">
            {PHASES.map((p, i) => (
              <div key={p.key} className={`${colClass(i)} cap-cell cap-feat`} onMouseEnter={() => focus(i)}>
                {p.features.map((f) => (
                  <span key={f.t} className={`cap-fc cap-fc--${f.s}`}><i className="cap-st" />{f.t}</span>
                ))}
              </div>
            ))}
          </div>
        </div>

        {/* tools row */}
        <div className="cap-row" style={{ '--r': 7 }}>
          <div className="cap-rowlab"><span>Tools / Stack</span><small>what runs it</small></div>
          <div className="cap-cells">
            {PHASES.map((p, i) => (
              <div key={p.key} className={`${colClass(i)} cap-cell cap-tool`} onMouseEnter={() => focus(i)}>{p.tools}</div>
            ))}
          </div>
        </div>

        {/* LLM row */}
        <div className="cap-row" style={{ '--r': 8 }}>
          <div className="cap-rowlab"><span>AI Model / LLM</span><small>brand</small></div>
          <div className="cap-cells">
            {PHASES.map((p, i) => (
              <div key={p.key} className={`${colClass(i)} cap-cell cap-llm`} onMouseEnter={() => focus(i)}>
                <span className="cap-llm__pri">{p.llm}</span>
                {p.llmOpt && <span className="cap-llm__opt">{p.llmOpt}</span>}
              </div>
            ))}
          </div>
        </div>

        {/* knowledge + human bands */}
        {BANDS.filter((b) => !b.spine).map((b, i) => (
          <div key={b.label} className="cap-row cap-row--band" style={{ '--r': 9 + i }}>
            <div className="cap-rowlab"><span>{b.label}</span><small>{i === 0 ? 'ground truth' : 'control'}</small></div>
            <div className="cap-band">
              <span className="cap-band__icon"><b.Icon size={17} /></span>
              <b>{b.title}</b>
              <span className="cap-band__txt">{b.text}</span>
            </div>
          </div>
        ))}
      </div>

      {/* ---- bottom strip ---- */}
      <footer className="cap-foot" style={{ '--r': 11 }}>
        <div className="cap-legend">
          {LEGEND.map((l) => (
            <span key={l.s} className="cap-legend__i"><i className={`cap-st cap-st--${l.s}`} />{l.label}</span>
          ))}
        </div>
        <div className="cap-wins">
          {WINS.map((w) => <span key={w.label} className="cap-win" style={{ '--c': w.tone }}>{w.label}</span>)}
        </div>
        <div className="cap-kpis">
          {KPIS.map((k) => (
            <div key={k.label} className="cap-kpi"><b><StatCounter value={k.value} suffix={k.suffix} /></b><span>{k.label}</span></div>
          ))}
        </div>
      </footer>

      <div className="cap-tagline" style={{ '--r': 12 }}>
        ★ Proven on <strong>ServiceTransfer</strong> (legacy VB6 POS agent) · Documentation is <em>step 1</em> — we make the foundation unbreakable.
      </div>
      </div>{/* /cap-fit */}
     </div>{/* /cap-stage */}

      {/* ---- floating controls ---- */}
      <div className="cap-controls">
        <button className="cap-ctrl-btn" onClick={() => setPlaying((p) => !p)} aria-label={playing ? 'Pause auto-tour' : 'Play auto-tour'}>
          {playing ? <Pause size={15} /> : <Play size={15} />}
          <span>{playing ? 'Auto-tour' : 'Paused'}</span>
        </button>
        <Link to="/flows" className="cap-ctrl-btn cap-ctrl-btn--accent"><Network size={15} /> <span>Flows</span></Link>
        <Link to="/downloads" className="cap-ctrl-btn"><Download size={15} /> <span>Deliverables</span></Link>
      </div>
    </div>
  )
}

export default Methodology
