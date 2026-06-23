import { NavLink } from 'react-router-dom'
import { Home, FileText, Database, ShieldCheck, Download, Code, X, Network, Layers, Presentation, ScanLine } from 'lucide-react'

const docLinks = [
  { to: '/docs/architecture-srs', icon: FileText, label: 'Architecture & SRS' },
  { to: '/docs/data-dictionary', icon: Database, label: 'Data Dictionary' },
  { to: '/docs/program-spec', icon: Code, label: 'Program Spec' },
  { to: '/docs/security', icon: ShieldCheck, label: 'Security & Tokens' },
  { to: '/docs/website-guide', icon: Layers, label: 'Website Dev Guide' },
]

function Sidebar({ isOpen, onClose }) {
  const closeOnMobile = () => { if (window.innerWidth < 1024) onClose() }
  const navClass = ({ isActive }) => `nav-item ${isActive ? 'active' : ''}`

  return (
    <>
      {isOpen && <div className="sidebar-overlay" onClick={onClose} />}

      <aside className={`sidebar ${isOpen ? 'open' : ''}`}>
        <div className="sidebar__header">
          <div className="sidebar__brand">
            <span className="sidebar__brand-mark"><ScanLine size={18} /></span>
            <span className="sidebar__brand-text">
              <strong>Project SCO</strong>
              <span>Self-Checkout Platform</span>
            </span>
          </div>
          <button className="sidebar__close" onClick={onClose} aria-label="Close menu">
            <X size={22} />
          </button>
        </div>

        <nav className="sidebar__nav">
          <NavLink to="/" end className={navClass} onClick={closeOnMobile}>
            <Home size={18} /> Overview
          </NavLink>
          <NavLink to="/sco" className={navClass} onClick={closeOnMobile}>
            <ScanLine size={18} /> SCO Kiosk
          </NavLink>
          <NavLink to="/methodology" className={navClass} onClick={closeOnMobile}>
            <Presentation size={18} /> Modernization Method
          </NavLink>

          <div className="nav-group">POS Foundation</div>
          <NavLink to="/flows" className={navClass} onClick={closeOnMobile}>
            <Network size={18} /> System Flows & Diagrams
          </NavLink>
          {docLinks.map(({ to, icon: Icon, label }) => (
            <NavLink key={to} to={to} className={navClass} onClick={closeOnMobile}>
              <Icon size={18} /> {label}
            </NavLink>
          ))}
        </nav>

        <div className="sidebar__footer">
          <NavLink to="/downloads" className={({ isActive }) => `nav-item download ${isActive ? 'active' : ''}`} onClick={closeOnMobile}>
            <Download size={18} /> Download Deliverables
          </NavLink>
        </div>
      </aside>
    </>
  )
}

export default Sidebar
