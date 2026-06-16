import { Menu } from 'lucide-react'
import { useLocation } from 'react-router-dom'
import ThemeToggle from './ThemeToggle'

const TITLES = {
  '': 'Introduction',
  methodology: 'Evidence-Driven Modernization',
  flows: 'System Flows & Diagrams',
  downloads: 'Download Deliverables',
  'docs/architecture-srs': 'Architecture & SRS',
  'docs/data-dictionary': 'Data Dictionary',
  'docs/program-spec': 'Program Spec',
  'docs/security': 'Security & Tokens',
  'docs/website-guide': 'Website Dev Guide',
}

function Topbar({ onToggleSidebar }) {
  const { pathname } = useLocation()
  const key = pathname.replace(/^\/+|\/+$/g, '')
  const current = TITLES[key] || 'Documentation'

  return (
    <header className="topbar">
      <button className="topbar__hamburger" onClick={onToggleSidebar} aria-label="Open menu">
        <Menu size={22} />
      </button>
      <nav className="topbar__crumbs" aria-label="Breadcrumb">
        <span>ServiceTransfer</span>
        <span className="sep">/</span>
        <span className="current">{current}</span>
      </nav>
      <div className="topbar__spacer" />
      <ThemeToggle />
    </header>
  )
}

export default Topbar
