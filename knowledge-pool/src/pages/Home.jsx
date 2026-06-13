import { Link } from 'react-router-dom'
import { Server, DatabaseZap, Code, ShieldCheck, Activity, Database, Layers, ArrowRight, Network } from 'lucide-react'
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
          <span className="cell__tag">OVERVIEW</span>
          <h1>ServiceTransfer</h1>
          <p>Interactive knowledge portal for the reverse-engineering &amp; revamp of the legacy VB6 background sync agent.</p>
          <div className="hero-chips">
            <span className="chip">VB6 → Modern</span>
            <span className="chip">Event-Driven</span>
            <span className="chip">API Gateway</span>
          </div>
          <div className="hero-cta">
            <Link to="/flows" className="btn primary"><Network size={16} /> Explore Flows</Link>
            <Link to="/docs/architecture-srs" className="btn ghost">Read the SRS <ArrowRight size={15} /></Link>
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
          <h3>What it does</h3>
          <p>A background agent on every branch POS that pushes pending sales (flagged <code>FTStaSentOnOff='0'</code>) to the central server, tokenizes card data via SafeNet, and accrues member loyalty points — all on a 500ms timer.</p>
        </section>

        <section className="cell cell--live">
          <span className="cell__tag">LIVE DIAGRAM</span>
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

        <Link to="/flows" className="cell cell--link cell--flows">
          <span className="cell__tag">SYSTEM FLOWS</span>
          <span className="flows-icon"><Layers size={28} /></span>
          <div>
            <h3>Interactive Diagrams</h3>
            <p>Step through the sync cycle, tokenization, member points, and the ER schema — all clickable.</p>
          </div>
          <ArrowRight size={22} color="var(--accent)" style={{ marginLeft: 'auto' }} />
        </Link>
      </div>
    </div>
  )
}

export default Home
