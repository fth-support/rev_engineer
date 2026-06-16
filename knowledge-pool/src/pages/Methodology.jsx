import { useEffect, useRef, useState, useCallback } from 'react'
import { Link } from 'react-router-dom'
import {
  Search, ShieldCheck, BarChart3, Ruler, Hammer, FlaskConical,
  Brain, BookOpen, UserCheck, Bot, Wrench, Sparkles, ChevronDown,
  ArrowRight, Network, Download, Quote, CheckCircle2,
} from 'lucide-react'
import StatCounter from '../components/primitives/StatCounter'
import './methodology.css'

const prefersReduced =
  typeof window !== 'undefined' &&
  window.matchMedia &&
  window.matchMedia('(prefers-reduced-motion: reduce)').matches

/* ------------------------------------------------------------------ data --- */

// The six phases of Evidence-Driven Modernization, mapped to the AI-platform
// lifecycle. Sourced from pitch/Evidence_Driven_Modernization.html and the
// verified ServiceTransfer engagement.
const PHASES = [
  {
    key: 'discover', n: '01', phase: 'PLAN', map: 'Planning', name: 'Discover',
    tone: 'var(--role-branch)', toneSoft: 'var(--role-branch-soft)', Icon: Search,
    agent: 'Code Archaeologist',
    blurb: 'Read the real VB6 line by line. Every module, call graph and dependency is mapped before a single assumption is made.',
    features: [
      { t: 'Source reading', s: 'p' },
      { t: 'Dependency map', s: 'p' },
      { t: 'Structure extraction', s: 'p' },
    ],
    tools: 'Grep · Glob · Read · Git', llm: 'Claude Opus 4.x',
  },
  {
    key: 'verify', n: '02', phase: 'ASSET', map: 'Asset Visualization', name: 'Verify',
    tone: 'var(--role-central)', toneSoft: 'var(--role-central-soft)', Icon: ShieldCheck,
    agent: 'Verification Agent',
    blurb: 'Every extracted fact is traced back to an exact file and line. If it is not in the source, it does not ship — zero hallucination.',
    features: [
      { t: 'Fact-tracing to source', s: 'p' },
      { t: 'Error detection', s: 'p' },
      { t: 'No-hallucination gate', s: 'p' },
    ],
    tools: 'cross-ref · diff · trace', llm: 'Claude Opus 4.x',
  },
  {
    key: 'visualize', n: '03', phase: 'DESIGN', map: 'Requirements / Design', name: 'Visualize',
    tone: 'var(--role-member)', toneSoft: 'var(--role-member-soft)', Icon: BarChart3,
    agent: 'Docs / Frontend Agent',
    blurb: 'Turn verified facts into a living, interactive knowledge portal — clickable ER & flow diagrams, themed UX, the system made legible.',
    features: [
      { t: 'Interactive portal', s: 'p' },
      { t: 'Live ER / flow diagrams', s: 'p' },
      { t: 'Theme / UX system', s: 'p' },
    ],
    tools: 'React · Vite · SVG · matplotlib', llm: 'Claude Opus 4.x',
  },
  {
    key: 'specify', n: '04', phase: 'SPEC', map: 'Design hand-off', name: 'Specify',
    tone: 'var(--role-safenet)', toneSoft: 'var(--role-safenet-soft)', Icon: Ruler,
    agent: 'System Analyst',
    blurb: 'Produce a build-ready directive: pseudocode, business rules and data contracts a team can implement in any language.',
    features: [
      { t: 'Any-language brief', s: 'p' },
      { t: 'Pseudocode + rules', s: 'p' },
      { t: 'Bilingual docs', s: 'p' },
    ],
    tools: 'Markdown / HTML · python-docx/pptx', llm: 'Claude Opus 4.x',
  },
  {
    key: 'modernize', n: '05', phase: 'BUILD', map: 'Implementation', name: 'Modernize',
    tone: 'var(--role-purple)', toneSoft: 'var(--role-purple-soft)', Icon: Hammer,
    agent: 'Solution Architect',
    blurb: 'Design the target architecture and keep-rules, then generate code — vendor-neutral, in the stack you choose.',
    features: [
      { t: 'Keep-rules + new arch', s: 'p' },
      { t: 'Target architecture', s: 'p' },
      { t: 'Code generation', s: 'a' },
    ],
    tools: 'any: Go · .NET · Java · Node', llm: 'Claude Opus 4.x', llmOpt: '＋ pluggable: Copilot · GPT',
  },
  {
    key: 'handoff', n: '06', phase: 'VERIFY', map: 'Test / Operate', name: 'Handoff',
    tone: 'var(--accent)', toneSoft: 'var(--accent-soft)', Icon: FlaskConical,
    agent: 'QA / Test Agent',
    blurb: 'Lock it in with acceptance criteria, CI/CD and an automated test suite so the new system is provably correct.',
    features: [
      { t: 'Acceptance criteria', s: 'p' },
      { t: 'CI/CD + deploy', s: 'p' },
      { t: 'Auto-test suite', s: 'a' },
    ],
    tools: 'GitHub Actions · tests', llm: 'Claude Opus 4.x',
  },
]

const PILLARS = [
  {
    Icon: Brain, tone: 'var(--role-branch)', title: 'One orchestrator',
    text: 'A single Claude Code agent drives all six steps end-to-end — one unbroken chain of traceability, no multi-agent hand-off gaps.',
  },
  {
    Icon: BookOpen, tone: 'var(--role-central)', title: 'Source is the truth',
    text: 'The legacy repository itself is the knowledge base. No external RAG or knowledge graph that can drift away from reality.',
  },
  {
    Icon: UserCheck, tone: 'var(--role-purple)', title: 'You stay in control',
    text: 'Scope, steer and approve at every gate. The agent does the heavy lifting; review is targeted, never the bottleneck.',
  },
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

const STATUS_LABEL = { p: 'Proven', a: 'Available', r: 'Roadmap' }

/* ----------------------------------------------------------------- helpers --- */

function scrollToId(id) {
  const el = document.getElementById(id)
  if (el) el.scrollIntoView({ behavior: prefersReduced ? 'auto' : 'smooth', block: 'start' })
}

// Reveal-on-scroll: returns [ref, shown]; flips once and stays.
function useReveal(rootMargin = '0px 0px -10% 0px') {
  const ref = useRef(null)
  const [shown, setShown] = useState(false)
  useEffect(() => {
    const el = ref.current
    if (!el) return
    if (prefersReduced || typeof IntersectionObserver === 'undefined') { setShown(true); return }
    const io = new IntersectionObserver(
      ([e]) => { if (e.isIntersecting) { setShown(true); io.disconnect() } },
      { threshold: 0.15, rootMargin },
    )
    io.observe(el)
    return () => io.disconnect()
  }, [rootMargin])
  return [ref, shown]
}

function Reveal({ as: Tag = 'div', className = '', delay = 0, children, ...rest }) {
  const [ref, shown] = useReveal()
  return (
    <Tag
      ref={ref}
      className={`reveal ${shown ? 'in' : ''} ${className}`}
      style={{ transitionDelay: `${delay}ms` }}
      {...rest}
    >
      {children}
    </Tag>
  )
}

/* -------------------------------------------------------------------- hero --- */

// Ambient, decorative bezier curves with particles flowing along them.
// Fixed viewBox + slice scaling means no node-alignment math is needed.
const HERO_CURVES = [
  { d: 'M -60 130 C 300 60, 520 250, 1260 150', tone: 'var(--role-branch)', dur: 7 },
  { d: 'M -60 360 C 360 300, 720 470, 1260 380', tone: 'var(--role-central)', dur: 9 },
  { d: 'M -60 560 C 300 520, 660 640, 1260 560', tone: 'var(--role-purple)', dur: 8 },
  { d: 'M 210 -40 C 320 220, 900 250, 1080 740', tone: 'var(--role-member)', dur: 11 },
]

function Hero() {
  const [parallax, setParallax] = useState(0)
  useEffect(() => {
    if (prefersReduced) return
    let ticking = false
    const onScroll = () => {
      if (ticking) return
      ticking = true
      requestAnimationFrame(() => {
        setParallax(Math.min(1, window.scrollY / 700))
        ticking = false
      })
    }
    window.addEventListener('scroll', onScroll, { passive: true })
    return () => window.removeEventListener('scroll', onScroll)
  }, [])

  return (
    <header className="m-hero" style={{ '--px': parallax }}>
      <div className="m-hero__aurora m-hero__aurora--a" />
      <div className="m-hero__aurora m-hero__aurora--b" />
      <div className="m-hero__grid" />

      <svg className="m-hero__svg" viewBox="0 0 1200 700" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
        {HERO_CURVES.map((c, i) => (
          <g key={i}>
            <path className="m-hero__curve" d={c.d} pathLength="1" stroke={c.tone}
              style={{ animationDelay: `${0.2 + i * 0.25}s` }} />
            <circle r="3.4" fill={c.tone} className="m-hero__particle"
              style={{ offsetPath: `path("${c.d}")`, animationDuration: `${c.dur}s`, animationDelay: `${i * 1.3}s` }} />
          </g>
        ))}
      </svg>

      <div className="m-hero__inner">
        <span className="m-hero__badge">
          <Sparkles size={14} /> WE DON&apos;T GUESS — WE VERIFY
        </span>
        <h1 className="m-hero__title">
          <span className="m-hero__title-kicker">Modernization Methodology</span>
          <span className="m-hero__title-main">Evidence-Driven<br />Modernization</span>
        </h1>
        <p className="m-hero__sub">
          The same end-to-end lifecycle as an AI modernization platform — but with a
          {' '}<strong>verification spine</strong>. Every step&apos;s output is traced back to the
          {' '}<strong>real source code</strong>, and the result is a build-ready directive any team
          can implement in <strong>any language</strong>.
        </p>
        <div className="m-hero__cta">
          <button className="m-btn m-btn--primary" onClick={() => scrollToId('journey')}>
            <Network size={16} /> Walk the journey
          </button>
          <Link to="/flows" className="m-btn m-btn--ghost">
            Explore live diagrams <ArrowRight size={15} />
          </Link>
        </div>
      </div>

      <button className="m-hero__scrollcue" onClick={() => scrollToId('thesis')} aria-label="Scroll down">
        <ChevronDown size={22} />
      </button>
    </header>
  )
}

/* ------------------------------------------------------------------ thesis --- */

function Thesis() {
  return (
    <section className="m-thesis" id="thesis">
      <Reveal className="m-thesis__col m-thesis__col--bad">
        <span className="m-thesis__tag m-thesis__tag--bad">Typical AI modernization</span>
        <h3>Generate, then hope</h3>
        <p>The model rewrites from a summary of the old system. Plausible output, untraceable claims — and bugs that only surface in production.</p>
      </Reveal>

      <Reveal className="m-thesis__bridge" delay={120}>
        <div className="m-thesis__bridge-line" />
        <span className="m-thesis__bridge-dot"><ShieldCheck size={18} /></span>
        <div className="m-thesis__bridge-line" />
      </Reveal>

      <Reveal className="m-thesis__col m-thesis__col--good" delay={200}>
        <span className="m-thesis__tag m-thesis__tag--good">Evidence-driven</span>
        <h3>Verify, then build</h3>
        <p>Every fact is anchored to a file and line in the legacy code first. The brief that comes out is portable, provable, and ready for any stack.</p>
      </Reveal>
    </section>
  )
}

/* ----------------------------------------------------------- pipeline (map) --- */

function PipelineOverview({ active, onJump }) {
  return (
    <section className="m-pipeline">
      <Reveal className="m-section-head">
        <span className="m-eyebrow">The capability map</span>
        <h2>Six steps, one chain of trust</h2>
        <p>Each phase maps to a stage of the platform lifecycle — click any step to jump into it.</p>
      </Reveal>

      <Reveal className="m-pipeline__track" delay={120}>
        <div className="m-pipeline__rail">
          <span className="m-pipeline__pulse" />
        </div>
        {PHASES.map((p, i) => {
          const { Icon } = p
          return (
            <button
              key={p.key}
              className={`m-pipeline__node ${active === i ? 'is-active' : ''}`}
              style={{ '--c': p.tone, '--cs': p.toneSoft }}
              onClick={() => onJump(`phase-${p.key}`)}
            >
              <span className="m-pipeline__dot"><Icon size={20} /></span>
              <span className="m-pipeline__n">{p.n}</span>
              <span className="m-pipeline__name">{p.name}</span>
              <span className="m-pipeline__map">＝ {p.map}</span>
            </button>
          )
        })}
      </Reveal>
    </section>
  )
}

/* ----------------------------------------------------------------- journey --- */

function PhaseRow({ p, index, setActive }) {
  const ref = useRef(null)
  const [shown, setShown] = useState(false)
  const { Icon } = p

  useEffect(() => {
    const el = ref.current
    if (!el) return
    if (prefersReduced || typeof IntersectionObserver === 'undefined') {
      setShown(true)
      return
    }
    const ioReveal = new IntersectionObserver(
      ([e]) => { if (e.isIntersecting) { setShown(true); ioReveal.disconnect() } },
      { threshold: 0.2, rootMargin: '0px 0px -12% 0px' },
    )
    const ioActive = new IntersectionObserver(
      ([e]) => { if (e.isIntersecting) setActive(index) },
      { threshold: 0, rootMargin: '-45% 0px -45% 0px' },
    )
    ioReveal.observe(el)
    ioActive.observe(el)
    return () => { ioReveal.disconnect(); ioActive.disconnect() }
  }, [index, setActive])

  return (
    <article
      ref={ref}
      id={`phase-${p.key}`}
      className={`m-phase reveal ${shown ? 'in' : ''}`}
      style={{ '--c': p.tone, '--cs': p.toneSoft }}
    >
      <div className="m-phase__rail">
        <span className="m-phase__node"><Icon size={20} /></span>
      </div>

      <div className="m-phase__card glass">
        <span className="m-phase__ghost" aria-hidden="true">{p.n}</span>
        <div className="m-phase__head">
          <span className="m-phase__kicker">{p.phase}<i>＝ {p.map}</i></span>
          <h3 className="m-phase__name">{p.name}</h3>
          <span className="m-phase__agent"><Bot size={14} /> {p.agent}</span>
        </div>
        <p className="m-phase__blurb">{p.blurb}</p>
        <ul className="m-phase__feats">
          {p.features.map((f) => (
            <li key={f.t} className={`m-feat m-feat--${f.s}`}>
              <span className="m-feat__dot" />
              {f.t}
              <span className="m-feat__status">{STATUS_LABEL[f.s]}</span>
            </li>
          ))}
        </ul>
        <div className="m-phase__foot">
          <span className="m-phase__tools"><Wrench size={13} /> {p.tools}</span>
          <span className="m-phase__llm">
            {p.llm}{p.llmOpt && <em>{p.llmOpt}</em>}
          </span>
        </div>
      </div>
    </article>
  )
}

function Journey({ active, setActive }) {
  const wrapRef = useRef(null)
  const [progress, setProgress] = useState(0)

  useEffect(() => {
    if (prefersReduced) { setProgress(1); return }
    let ticking = false
    const measure = () => {
      const el = wrapRef.current
      if (!el) return
      const rect = el.getBoundingClientRect()
      const vh = window.innerHeight
      const start = vh * 0.5
      const total = Math.max(1, rect.height - vh * 0.25)
      const passed = start - rect.top
      setProgress(Math.max(0, Math.min(1, passed / total)))
    }
    const onScroll = () => {
      if (ticking) return
      ticking = true
      requestAnimationFrame(() => { measure(); ticking = false })
    }
    window.addEventListener('scroll', onScroll, { passive: true })
    window.addEventListener('resize', onScroll)
    measure()
    return () => {
      window.removeEventListener('scroll', onScroll)
      window.removeEventListener('resize', onScroll)
    }
  }, [])

  return (
    <section className="m-journey" id="journey" ref={wrapRef}
      style={{ '--progress': progress, '--active-c': PHASES[active].tone }}>
      <Reveal className="m-section-head">
        <span className="m-eyebrow">The journey</span>
        <h2>From legacy black box to build-ready brief</h2>
        <p>Scroll through the six phases. The spine fills as the evidence chain is built.</p>
      </Reveal>

      <div className="m-journey__body">
        <div className="m-journey__spine" aria-hidden="true">
          <span className="m-journey__spine-fill" />
          <span className="m-journey__spark" />
        </div>
        <div className="m-journey__rows">
          {PHASES.map((p, i) => (
            <PhaseRow key={p.key} p={p} index={i} setActive={setActive} />
          ))}
        </div>
      </div>
    </section>
  )
}

/* ----------------------------------------------------------------- pillars --- */

function Pillars() {
  return (
    <section className="m-pillars">
      <Reveal className="m-section-head">
        <span className="m-eyebrow">Why it holds</span>
        <h2>The verification spine</h2>
        <p>Three structural choices keep the output honest from first read to final test.</p>
      </Reveal>
      <div className="m-pillars__grid">
        {PILLARS.map((p, i) => {
          const { Icon } = p
          return (
            <Reveal key={p.title} className="m-pillar glass" delay={i * 110} style={{ '--c': p.tone }}>
              <span className="m-pillar__icon"><Icon size={22} /></span>
              <h3>{p.title}</h3>
              <p>{p.text}</p>
            </Reveal>
          )
        })}
      </div>
    </section>
  )
}

/* ------------------------------------------------------------------- proof --- */

function Proof() {
  return (
    <section className="m-proof">
      <Reveal className="m-proof__quote glass">
        <Quote size={26} className="m-proof__quote-mark" />
        <p>
          Proven end-to-end on <strong>ServiceTransfer</strong> — a live legacy VB6 POS sync agent.
          Documentation is <em>step one</em>; we make the foundation unbreakable.
        </p>
      </Reveal>

      <div className="m-proof__kpis">
        {KPIS.map((k, i) => (
          <Reveal key={k.label} className="m-kpi glass" delay={i * 90}>
            <b><StatCounter value={k.value} suffix={k.suffix} /></b>
            <span>{k.label}</span>
          </Reveal>
        ))}
      </div>

      <Reveal className="m-wins" delay={120}>
        {WINS.map((w) => (
          <span key={w.label} className="m-win" style={{ '--c': w.tone }}>
            <CheckCircle2 size={14} /> {w.label}
          </span>
        ))}
      </Reveal>
    </section>
  )
}

/* --------------------------------------------------------------------- cta --- */

function ClosingCTA() {
  return (
    <Reveal as="section" className="m-cta">
      <div className="m-cta__inner glass">
        <h2>Ready to make your legacy system legible?</h2>
        <p>Step through the interactive diagrams, or take the full documentation set with you.</p>
        <div className="m-cta__actions">
          <Link to="/flows" className="m-btn m-btn--primary"><Network size={16} /> Explore the flows</Link>
          <Link to="/downloads" className="m-btn m-btn--ghost"><Download size={15} /> Download deliverables</Link>
        </div>
      </div>
    </Reveal>
  )
}

/* -------------------------------------------------------------------- page --- */

function Methodology() {
  const [active, setActive] = useState(0)
  const handleJump = useCallback((id) => scrollToId(id), [])

  return (
    <div className="methodology">
      <Hero />
      <Thesis />
      <PipelineOverview active={active} onJump={handleJump} />
      <Journey active={active} setActive={setActive} />
      <Pillars />
      <Proof />
      <ClosingCTA />
    </div>
  )
}

export default Methodology
