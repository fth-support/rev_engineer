import { useEffect, useRef, useState, useCallback } from 'react'
import { Link } from 'react-router-dom'
import {
  Search, ShieldCheck, BarChart3, Ruler, Hammer, FlaskConical,
  Brain, BookOpen, UserCheck, Bot, Sparkles, ArrowRight, Network,
  Download, CheckCircle2, ChevronLeft, ChevronRight, Maximize2, Minimize2,
  FileText, MousePointerClick,
} from 'lucide-react'
import StatCounter from '../components/primitives/StatCounter'
import ArchitectureGraph from '../components/diagrams/ArchitectureGraph'
import TxLifecycleStepper from '../components/diagrams/TxLifecycleStepper'
import './methodology.css'

const prefersReduced =
  typeof window !== 'undefined' &&
  window.matchMedia &&
  window.matchMedia('(prefers-reduced-motion: reduce)').matches

/* ------------------------------------------------------------------ data --- */

const PHASES = [
  {
    key: 'discover', n: '01', phase: 'PLAN', map: 'Planning', name: 'Discover',
    tone: 'var(--role-branch)', toneSoft: 'var(--role-branch-soft)', Icon: Search,
    agent: 'Code Archaeologist',
    blurb: 'Read the real VB6 line by line — every module, call graph and dependency mapped before a single assumption is made.',
    features: [{ t: 'Source reading', s: 'p' }, { t: 'Dependency map', s: 'p' }, { t: 'Structure extraction', s: 'p' }],
    tools: 'Grep · Glob · Read · Git',
  },
  {
    key: 'verify', n: '02', phase: 'ASSET', map: 'Asset Visualization', name: 'Verify',
    tone: 'var(--role-central)', toneSoft: 'var(--role-central-soft)', Icon: ShieldCheck,
    agent: 'Verification Agent',
    blurb: 'Every extracted fact is traced back to an exact file and line. If it is not in the source, it does not ship.',
    features: [{ t: 'Fact-tracing', s: 'p' }, { t: 'Error detection', s: 'p' }, { t: 'No-hallucination gate', s: 'p' }],
    tools: 'cross-ref · diff · trace',
  },
  {
    key: 'visualize', n: '03', phase: 'DESIGN', map: 'Requirements / Design', name: 'Visualize',
    tone: 'var(--role-member)', toneSoft: 'var(--role-member-soft)', Icon: BarChart3,
    agent: 'Docs / Frontend Agent',
    blurb: 'Turn verified facts into a living, interactive portal — clickable ER & flow diagrams, themed UX, the system made legible.',
    features: [{ t: 'Interactive portal', s: 'p' }, { t: 'Live ER / flow', s: 'p' }, { t: 'Theme / UX system', s: 'p' }],
    tools: 'React · Vite · SVG',
  },
  {
    key: 'specify', n: '04', phase: 'SPEC', map: 'Design hand-off', name: 'Specify',
    tone: 'var(--role-safenet)', toneSoft: 'var(--role-safenet-soft)', Icon: Ruler,
    agent: 'System Analyst',
    blurb: 'Produce a build-ready directive: pseudocode, business rules and data contracts a team can implement in any language.',
    features: [{ t: 'Any-language brief', s: 'p' }, { t: 'Pseudocode + rules', s: 'p' }, { t: 'Bilingual docs', s: 'p' }],
    tools: 'Markdown · docx/pptx',
  },
  {
    key: 'modernize', n: '05', phase: 'BUILD', map: 'Implementation', name: 'Modernize',
    tone: 'var(--role-purple)', toneSoft: 'var(--role-purple-soft)', Icon: Hammer,
    agent: 'Solution Architect',
    blurb: 'Design the target architecture and keep-rules, then generate code — vendor-neutral, in the stack you choose.',
    features: [{ t: 'Keep-rules + new arch', s: 'p' }, { t: 'Target architecture', s: 'p' }, { t: 'Code generation', s: 'a' }],
    tools: 'Go · .NET · Java · Node',
  },
  {
    key: 'handoff', n: '06', phase: 'VERIFY', map: 'Test / Operate', name: 'Handoff',
    tone: 'var(--accent)', toneSoft: 'var(--accent-soft)', Icon: FlaskConical,
    agent: 'QA / Test Agent',
    blurb: 'Lock it in with acceptance criteria, CI/CD and an automated test suite so the new system is provably correct.',
    features: [{ t: 'Acceptance criteria', s: 'p' }, { t: 'CI/CD + deploy', s: 'p' }, { t: 'Auto-test suite', s: 'a' }],
    tools: 'GitHub Actions · tests',
  },
]

const PILLARS = [
  { Icon: Brain, tone: 'var(--role-branch)', title: 'One orchestrator', text: 'A single Claude Code agent drives all six steps end-to-end — one unbroken chain of traceability.' },
  { Icon: BookOpen, tone: 'var(--role-central)', title: 'Source is the truth', text: 'The legacy repository itself is the knowledge base. No external RAG that can drift from reality.' },
  { Icon: UserCheck, tone: 'var(--role-purple)', title: 'You stay in control', text: 'Scope, steer and approve at every gate. The agent does the heavy lifting; review is never the bottleneck.' },
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

// Slide ids drive the progress dots / nav. Keep in sync with the rendered order.
const SLIDES = [
  { id: 'cover', label: 'Cover' },
  { id: 'problem', label: 'The problem' },
  { id: 'method', label: 'The method' },
  { id: 'matrix', label: 'Capability map' },
  { id: 'architecture', label: 'Live architecture' },
  { id: 'flow', label: 'Live flow' },
  { id: 'spine', label: 'Why it holds' },
  { id: 'proof', label: 'Proof' },
  { id: 'start', label: 'Explore' },
]

const HERO_CURVES = [
  { d: 'M -60 130 C 300 60, 520 250, 1260 150', tone: 'var(--role-branch)', dur: 7 },
  { d: 'M -60 360 C 360 300, 720 470, 1260 380', tone: 'var(--role-central)', dur: 9 },
  { d: 'M -60 560 C 300 520, 660 640, 1260 560', tone: 'var(--role-purple)', dur: 8 },
  { d: 'M 210 -40 C 320 220, 900 250, 1080 740', tone: 'var(--role-member)', dur: 11 },
]

/* ------------------------------------------------------------------ deck ---- */

function Methodology() {
  const deckRef = useRef(null)
  const [active, setActive] = useState(0)
  const activeRef = useRef(0)
  const [fs, setFs] = useState(false)
  const total = SLIDES.length

  useEffect(() => { activeRef.current = active }, [active])

  const goTo = useCallback((i) => {
    const deck = deckRef.current
    if (!deck) return
    const idx = Math.max(0, Math.min(SLIDES.length - 1, i))
    deck.scrollTo({ top: idx * deck.clientHeight, behavior: prefersReduced ? 'auto' : 'smooth' })
  }, [])

  // Track the active slide as it snaps into view.
  useEffect(() => {
    const deck = deckRef.current
    if (!deck) return
    const slides = Array.from(deck.querySelectorAll('.slide'))
    const io = new IntersectionObserver(
      (entries) => {
        entries.forEach((en) => {
          if (en.isIntersecting && en.intersectionRatio >= 0.55) {
            const idx = slides.indexOf(en.target)
            if (idx >= 0) setActive(idx)
          }
        })
      },
      { root: deck, threshold: [0.55] },
    )
    slides.forEach((s) => io.observe(s))
    return () => io.disconnect()
  }, [])

  // Keyboard navigation — arrows / space / page / home / end.
  useEffect(() => {
    const onKey = (e) => {
      if (e.target && /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Spacebar'].includes(e.key)) {
        e.preventDefault(); goTo(activeRef.current + 1)
      } else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) {
        e.preventDefault(); goTo(activeRef.current - 1)
      } else if (e.key === 'Home') { e.preventDefault(); goTo(0) }
      else if (e.key === 'End') { e.preventDefault(); goTo(SLIDES.length - 1) }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [goTo])

  // Fullscreen "present" mode.
  useEffect(() => {
    const onFs = () => setFs(Boolean(document.fullscreenElement))
    document.addEventListener('fullscreenchange', onFs)
    return () => document.removeEventListener('fullscreenchange', onFs)
  }, [])
  const togglePresent = useCallback(() => {
    const deck = deckRef.current
    if (!document.fullscreenElement) deck?.requestFullscreen?.()
    else document.exitFullscreen?.()
  }, [])

  const cx = (i) => `slide ${active === i ? 'is-active' : ''}`

  return (
    <div className="deck-page">
      <div className="deck" ref={deckRef}>

        {/* 1 — COVER ----------------------------------------------------- */}
        <section className={cx(0)} id="slide-cover" data-tone="var(--accent)">
          <div className="cover-bg" aria-hidden="true">
            <span className="cover-aurora cover-aurora--a" />
            <span className="cover-aurora cover-aurora--b" />
            <span className="cover-grid" />
            <svg className="cover-svg" viewBox="0 0 1200 700" preserveAspectRatio="xMidYMid slice">
              {HERO_CURVES.map((c, i) => (
                <g key={i}>
                  <path className="cover-curve" d={c.d} pathLength="1" stroke={c.tone} style={{ animationDelay: `${0.2 + i * 0.25}s` }} />
                  <circle r="3.4" fill={c.tone} className="cover-particle"
                    style={{ offsetPath: `path("${c.d}")`, animationDuration: `${c.dur}s`, animationDelay: `${i * 1.3}s` }} />
                </g>
              ))}
            </svg>
          </div>
          <div className="slide__inner cover">
            <span className="s-anim cover__badge" style={{ '--d': 0 }}><Sparkles size={14} /> WE DON&apos;T GUESS — WE VERIFY</span>
            <span className="s-anim cover__kicker" style={{ '--d': 1 }}>Modernization Methodology · 1-page brief</span>
            <h1 className="s-anim cover__title" style={{ '--d': 2 }}>Evidence-Driven<br />Modernization</h1>
            <p className="s-anim cover__sub" style={{ '--d': 3 }}>
              The same end-to-end lifecycle as an AI modernization platform — but with a
              {' '}<strong>verification spine</strong>. Every step is traced back to the
              {' '}<strong>real source code</strong>, producing a build-ready directive for <strong>any language</strong>.
            </p>
            <button className="s-anim m-btn m-btn--primary cover__cta" style={{ '--d': 4 }} onClick={() => goTo(1)}>
              Start the story <ArrowRight size={16} />
            </button>
            <span className="s-anim cover__hint" style={{ '--d': 5 }}>
              <MousePointerClick size={13} /> Use <kbd>→</kbd> / <kbd>Space</kbd> · scroll · or the dots →
            </span>
          </div>
        </section>

        {/* 2 — PROBLEM --------------------------------------------------- */}
        <section className={cx(1)} id="slide-problem">
          <div className="slide__inner">
            <header className="slide-head s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow">The fork in the road</span>
              <h2>Two ways to modernize a legacy system</h2>
            </header>
            <div className="problem">
              <div className="s-anim problem__col problem__col--bad" style={{ '--d': 1 }}>
                <span className="problem__tag problem__tag--bad">Typical AI rewrite</span>
                <h3>Generate, then hope</h3>
                <p>The model rewrites from a summary. Plausible output, untraceable claims — and bugs that only surface in production.</p>
              </div>
              <div className="s-anim problem__bridge" style={{ '--d': 2 }}>
                <span className="problem__bridge-dot"><ShieldCheck size={20} /></span>
              </div>
              <div className="s-anim problem__col problem__col--good" style={{ '--d': 3 }}>
                <span className="problem__tag problem__tag--good">Evidence-driven</span>
                <h3>Verify, then build</h3>
                <p>Every fact is anchored to a file and line first. The brief that comes out is portable, provable and ready for any stack.</p>
              </div>
            </div>
          </div>
        </section>

        {/* 3 — METHOD (pipeline) ---------------------------------------- */}
        <section className={cx(2)} id="slide-method">
          <div className="slide__inner">
            <header className="slide-head s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow">The capability map</span>
              <h2>Six steps, one chain of trust</h2>
              <p>Each phase maps to a stage of the platform lifecycle.</p>
            </header>
            <div className="pipe">
              <div className="pipe__rail"><span className="pipe__fill" /></div>
              {PHASES.map((p, i) => {
                const { Icon } = p
                return (
                  <div key={p.key} className="s-anim pipe__node" style={{ '--c': p.tone, '--cs': p.toneSoft, '--d': i + 1 }}>
                    <span className="pipe__dot"><Icon size={22} /></span>
                    <span className="pipe__n">{p.n}</span>
                    <span className="pipe__name">{p.name}</span>
                    <span className="pipe__map">＝ {p.map}</span>
                  </div>
                )
              })}
            </div>
          </div>
        </section>

        {/* 4 — MATRIX (phase cards) ------------------------------------- */}
        <section className={cx(3)} id="slide-matrix">
          <div className="slide__inner">
            <header className="slide-head s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow">Capabilities × agents</span>
              <h2>One orchestrator, six specialist roles</h2>
            </header>
            <div className="matrix">
              {PHASES.map((p, i) => {
                const { Icon } = p
                return (
                  <article key={p.key} className="s-anim mcard" style={{ '--c': p.tone, '--cs': p.toneSoft, '--d': i + 1 }}>
                    <span className="mcard__ghost" aria-hidden="true">{p.n}</span>
                    <span className="mcard__icon"><Icon size={18} /></span>
                    <h3 className="mcard__name">{p.name}</h3>
                    <span className="mcard__agent"><Bot size={12} /> {p.agent}</span>
                    <ul className="mcard__feats">
                      {p.features.map((f) => (
                        <li key={f.t} className={`mfeat mfeat--${f.s}`}><span className="mfeat__dot" />{f.t}</li>
                      ))}
                    </ul>
                    <span className="mcard__tools">{p.tools}</span>
                  </article>
                )
              })}
            </div>
          </div>
        </section>

        {/* 5 — LIVE ARCHITECTURE ---------------------------------------- */}
        <section className={cx(4)} id="slide-architecture">
          <div className="slide__inner slide__inner--diagram">
            <header className="slide-head slide-head--tight s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow"><MousePointerClick size={12} /> Live · interactive</span>
              <h2>What we reverse-engineered</h2>
            </header>
            <div className="slide-diagram s-anim" style={{ '--d': 1 }}>
              <ArchitectureGraph />
            </div>
          </div>
        </section>

        {/* 6 — LIVE FLOW ------------------------------------------------- */}
        <section className={cx(5)} id="slide-flow">
          <div className="slide__inner slide__inner--diagram">
            <header className="slide-head slide-head--tight s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow"><MousePointerClick size={12} /> Live · step through it</span>
              <h2>The end-to-end transaction lifecycle</h2>
            </header>
            <div className="slide-diagram slide-diagram--stepper s-anim" style={{ '--d': 1 }}>
              <TxLifecycleStepper />
            </div>
          </div>
        </section>

        {/* 7 — SPINE (pillars) ------------------------------------------ */}
        <section className={cx(6)} id="slide-spine">
          <div className="slide__inner">
            <header className="slide-head s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow">Why it holds</span>
              <h2>The verification spine</h2>
            </header>
            <div className="pillars">
              {PILLARS.map((p, i) => {
                const { Icon } = p
                return (
                  <article key={p.title} className="s-anim pillar" style={{ '--c': p.tone, '--d': i + 1 }}>
                    <span className="pillar__icon"><Icon size={22} /></span>
                    <h3>{p.title}</h3>
                    <p>{p.text}</p>
                  </article>
                )
              })}
            </div>
          </div>
        </section>

        {/* 8 — PROOF ----------------------------------------------------- */}
        <section className={cx(7)} id="slide-proof">
          <div className="slide__inner">
            <header className="slide-head s-anim" style={{ '--d': 0 }}>
              <span className="m-eyebrow">Proven, not promised</span>
              <h2>End-to-end on a live legacy system</h2>
              <p>ServiceTransfer — a VB6 POS sync agent — reverse-engineered and re-specified.</p>
            </header>
            <div className="kpis">
              {KPIS.map((k, i) => (
                <div key={k.label} className="s-anim kpi" style={{ '--d': i + 1 }}>
                  <b><StatCounter value={k.value} suffix={k.suffix} /></b>
                  <span>{k.label}</span>
                </div>
              ))}
            </div>
            <div className="wins s-anim" style={{ '--d': 5 }}>
              {WINS.map((w) => (
                <span key={w.label} className="win" style={{ '--c': w.tone }}><CheckCircle2 size={14} /> {w.label}</span>
              ))}
            </div>
          </div>
        </section>

        {/* 9 — EXPLORE / CTA -------------------------------------------- */}
        <section className={cx(8)} id="slide-start">
          <div className="cover-bg" aria-hidden="true">
            <span className="cover-aurora cover-aurora--a" />
            <span className="cover-aurora cover-aurora--b" />
          </div>
          <div className="slide__inner cover">
            <span className="s-anim m-eyebrow" style={{ '--d': 0 }}>Now explore the portal</span>
            <h2 className="s-anim cover__title cover__title--sm" style={{ '--d': 1 }}>See it for yourself</h2>
            <p className="s-anim cover__sub" style={{ '--d': 2 }}>
              Step through every flow, browse the verified data dictionary, or take the full documentation set with you.
            </p>
            <div className="s-anim start__grid" style={{ '--d': 3 }}>
              <Link to="/flows" className="start__card"><Network size={20} /><b>System Flows</b><span>Interactive diagrams</span></Link>
              <Link to="/docs/data-dictionary" className="start__card"><FileText size={20} /><b>Documentation</b><span>Verified specs</span></Link>
              <Link to="/downloads" className="start__card"><Download size={20} /><b>Deliverables</b><span>docx · pptx · html</span></Link>
            </div>
          </div>
        </section>
      </div>

      {/* ---- deck chrome (progress + controls) ---- */}
      <div className={`deck-progress ${fs ? 'in-fs' : ''}`} style={{ '--p': (active + 1) / total }}>
        <span className="deck-progress__bar" />
      </div>

      <nav className={`deck-dots ${fs ? 'in-fs' : ''}`} aria-label="Slides">
        {SLIDES.map((s, i) => (
          <button key={s.id} className={`deck-dot ${active === i ? 'is-active' : ''}`} onClick={() => goTo(i)} aria-label={s.label} title={s.label}>
            <span />
          </button>
        ))}
      </nav>

      <div className={`deck-ctrl ${fs ? 'in-fs' : ''}`}>
        <button className="deck-ctrl__btn" onClick={() => goTo(active - 1)} disabled={active === 0} aria-label="Previous slide"><ChevronLeft size={18} /></button>
        <span className="deck-ctrl__count">{active + 1} / {total}</span>
        <button className="deck-ctrl__btn" onClick={() => goTo(active + 1)} disabled={active === total - 1} aria-label="Next slide"><ChevronRight size={18} /></button>
        <button className="deck-ctrl__btn deck-ctrl__present" onClick={togglePresent} aria-label={fs ? 'Exit present mode' : 'Present fullscreen'} title={fs ? 'Exit' : 'Present'}>
          {fs ? <Minimize2 size={16} /> : <Maximize2 size={16} />}
        </button>
      </div>
    </div>
  )
}

export default Methodology
