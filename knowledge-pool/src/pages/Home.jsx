import { Link } from 'react-router-dom'
import { Server, DatabaseZap, Code, ShieldCheck, Activity, Database, ArrowRight, Network, Sparkles, ScanLine } from 'lucide-react'
import StatCounter from '../components/primitives/StatCounter'
import ArchitectureGraph from '../components/diagrams/ArchitectureGraph'
import './pages.css'

const docCards = [
  { area: 'doc1', to: '/docs/architecture-srs', icon: Server, color: 'var(--role-central)', soft: 'var(--role-central-soft)', title: 'Architecture & SRS', desc: 'System overview & requirements' },
  { area: 'doc2', to: '/docs/data-dictionary', icon: DatabaseZap, color: 'var(--role-branch)', soft: 'var(--role-branch-soft)', title: 'Data Dictionary', desc: '11 core tables, verified schema' },
  { area: 'doc3', to: '/docs/program-spec', icon: Code, color: 'var(--role-purple)', soft: 'var(--role-purple-soft)', title: 'Program Spec', desc: 'Functions & control flow' },
  { area: 'doc4', to: '/docs/security', icon: ShieldCheck, color: 'var(--role-safenet)', soft: 'var(--role-safenet-soft)', title: 'Security & Tokens', desc: 'SafeNet tokenization' },
]

function Home() {
  return (
    <div className="home">
      <div className="bento">
        <section className="cell cell--hero">
          <span className="cell__tag">PROJECT SCO</span>
          <h1>Self-Checkout, your way</h1>
          <p>A modular, white-label self-checkout platform — one standard engine any retailer can recolor, rebrand and trim, integrated through a single POS API. Foodland is the reference tenant.</p>
          <div className="hero-chips">
            <span className="chip">White-label</span>
            <span className="chip">Modular</span>
            <span className="chip">Offline-first</span>
          </div>
          <div className="hero-cta">
            <Link to="/sco" className="btn primary"><ScanLine size={16} /> Open SCO Kiosk</Link>
            <Link to="/flows" className="btn ghost"><Network size={15} /> POS foundation</Link>
            <Link to="/methodology" className="btn ghost"><Sparkles size={15} /> The method <ArrowRight size={15} /></Link>
          </div>
        </section>

        <section className="cell cell--stat1 cell--stat">
          <span className="cell__tag">POLLING</span>
          <Activity size={26} color="var(--warning)" className="stat-icon" />
          <div className="stat-num"><StatCounter value={500} /><small>ms</small></div>
          <div className="stat-label">Legacy timer loop — replaced by real-time events</div>
        </section>

        <section className="cell cell--stat2 cell--stat">
          <span className="cell__tag">DATA</span>
          <Database size={26} color="var(--role-central)" className="stat-icon" />
          <div className="stat-num"><StatCounter value={3} /><small> DBs</small></div>
          <div className="stat-label">Local · Central (HQ) · Member — + SafeNet tokenizer</div>
        </section>

        <section className="cell cell--about">
          <span className="cell__tag">ABOUT</span>
          <h3>What it is</h3>
          <p>An interactive self-checkout kiosk (ไทย/EN/中文, offline-first) that reads SKU, promo, member &amp; points from the POS and writes finished sales back — the SCO lane is treated as one POS terminal on the store network.</p>
        </section>

        <section className="cell cell--live">
          <span className="cell__tag">POS FOUNDATION · LIVE</span>
          <ArchitectureGraph />
        </section>

        {docCards.map(({ area, to, icon: Icon, color, soft, title, desc }) => (
          <Link key={area} to={to} className={`cell cell--link cell--${area} doc-card`}>
            <span className="cell__tag">DOCS</span>
            <span className="doc-icon" style={{ background: soft, color }}><Icon size={22} /></span>
            <h4>{title}</h4>
            <p>{desc}</p>
          </Link>
        ))}

        <Link to="/sco" className="cell cell--link cell--flows">
          <span className="cell__tag">SCO KIOSK</span>
          <span className="flows-icon"><ScanLine size={28} /></span>
          <div>
            <h3>Interactive kiosk demo</h3>
            <p>Tap through the full shopper loop — idle → scan → member → payment → receipt — plus edge cases and staff mode.</p>
          </div>
          <ArrowRight size={22} color="var(--accent)" style={{ marginLeft: 'auto' }} />
        </Link>
      </div>
    </div>
  )
}

export default Home
