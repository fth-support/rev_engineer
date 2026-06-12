import { NavLink } from 'react-router-dom'
import { Home, FileText, Database, ShieldCheck, Download, Code, X, Network, Map } from 'lucide-react'

function Sidebar({ isOpen, onClose }) {
  return (
    <>
      {/* Overlay for mobile */}
      {isOpen && (
        <div 
          className="sidebar-overlay" 
          onClick={onClose}
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', 
            zIndex: 90, display: window.innerWidth < 1024 ? 'block' : 'none'
          }}
        />
      )}

      <aside className={`sidebar ${isOpen ? 'open' : ''}`}>
        <div className="sidebar-header">
          <h3>ServiceTransfer</h3>
          <button className="close-btn" onClick={onClose}>
            <X size={24} />
          </button>
        </div>

        <div className="nav-links">
          <NavLink to="/" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <Home size={18} /> Introduction & Journey
          </NavLink>

          <NavLink to="/flows" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <Network size={18} /> System Flows & Diagrams
          </NavLink>

          <div className="nav-group">Documents</div>
          
          <NavLink to="/docs/architecture-srs" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <FileText size={18} /> Architecture & SRS
          </NavLink>
          <NavLink to="/docs/data-dictionary" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <Database size={18} /> Data Dictionary
          </NavLink>
          <NavLink to="/docs/program-spec" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <Code size={18} /> Program Spec
          </NavLink>
          <NavLink to="/docs/security" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
            <ShieldCheck size={18} /> Security & Tokens
          </NavLink>
        </div>

        <div className="sidebar-footer">
        <NavLink to="/downloads" className={({isActive}) => `nav-item download-btn ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 1024 && onClose()}>
          <Download size={18} /> Download Original
        </NavLink>
      </div>
      </aside>
    </>
  )
}

export default Sidebar
