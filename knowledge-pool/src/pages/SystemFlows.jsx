import { useState } from 'react'
import { Link } from 'react-router-dom'
import { Network, Database, RefreshCw, ShieldCheck, CreditCard, Rocket, ArrowRight, GitBranch } from 'lucide-react'
import ArchitectureGraph from '../components/diagrams/ArchitectureGraph'
import TxLifecycleStepper from '../components/diagrams/TxLifecycleStepper'
import SyncFlowStepper from '../components/diagrams/SyncFlowStepper'
import ERGraph from '../components/diagrams/ERGraph'
import TokenizationStepper from '../components/diagrams/TokenizationStepper'
import MemberPointsStepper from '../components/diagrams/MemberPointsStepper'
import TargetArchitectureGraph from '../components/diagrams/TargetArchitectureGraph'
import './pages.css'

const flows = [
  { id: 'architecture', title: 'System Architecture', icon: RefreshCw, Comp: ArchitectureGraph, doc: '/docs/architecture-srs', docLabel: 'Architecture & SRS' },
  { id: 'lifecycle', title: 'End-to-End Transaction', icon: GitBranch, Comp: TxLifecycleStepper, doc: '/docs/architecture-srs', docLabel: 'Architecture & SRS' },
  { id: 'sync', title: 'Data Sync Flow', icon: RefreshCw, Comp: SyncFlowStepper, doc: '/docs/program-spec', docLabel: 'Program Spec' },
  { id: 'er', title: 'ER Diagram', icon: Database, Comp: ERGraph, doc: '/docs/data-dictionary', docLabel: 'Data Dictionary' },
  { id: 'token', title: 'Tokenization Flow', icon: ShieldCheck, Comp: TokenizationStepper, doc: '/docs/security', docLabel: 'Security & Tokens' },
  { id: 'points', title: 'Member Points Flow', icon: CreditCard, Comp: MemberPointsStepper, doc: '/docs/program-spec', docLabel: 'Program Spec' },
  { id: 'target', title: 'Target Architecture', icon: Rocket, Comp: TargetArchitectureGraph, doc: '/docs/architecture-srs', docLabel: 'Architecture & SRS' },
]
// Make the nav icons match their concept
flows[0].icon = Network

function SystemFlows() {
  const [activeId, setActiveId] = useState(flows[0].id)
  const active = flows.find((f) => f.id === activeId)
  const ActiveComp = active.Comp

  return (
    <div className="flows-page">
      <div className="flows-head glass">
        <h2>System Flows &amp; Diagrams</h2>
        <p>Interactive architecture and workflow diagrams — click nodes, step through flows, and switch theme.</p>
      </div>

      <div className="flows-layout">
        <nav className="flows-nav glass">
          {flows.map(({ id, title, icon: Icon }) => (
            <button key={id} className={`flow-nav-btn ${activeId === id ? 'active' : ''}`} onClick={() => setActiveId(id)}>
              <Icon size={18} /> {title}
            </button>
          ))}
        </nav>

        <div>
          <div className="flow-canvas glass">
            <ActiveComp />
          </div>
          <div className="flow-doc glass" style={{ marginTop: '1.25rem' }}>
            <h3>Related documentation</h3>
            <p>Full written specification for this flow, verified against the legacy source code.</p>
            <Link to={active.doc} className="btn primary" style={{ display: 'inline-flex' }}>
              Open {active.docLabel} <ArrowRight size={15} />
            </Link>
          </div>
        </div>
      </div>
    </div>
  )
}

export default SystemFlows
