import { NavLink } from 'react-router-dom'
import { X, Home, Map, FileText, Database, Code, ShieldCheck, Download } from 'lucide-react'

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
            zIndex: 90, display: window.innerWidth < 768 ? 'block' : 'none'
          }}
        />
      )}

      <aside className={`sidebar ${isOpen ? 'open' : ''}`}>
        <div className="sidebar-header">
          <h3>AGY Docs</h3>
          <button className="close-btn" onClick={onClose}>
            <X size={24} />
          </button>
        </div>

        <div className="nav-links">
          <NavLink to="/" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <Home size={18} /> Introduction
          </NavLink>
          <NavLink to="/journey" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <Map size={18} /> App Journey
          </NavLink>

          <div className="nav-group">Documents</div>
          
          <NavLink to="/docs/architecture-srs" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <FileText size={18} /> Architecture & SRS
          </NavLink>
          <NavLink to="/docs/data-dictionary" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <Database size={18} /> Data Dictionary
          </NavLink>
          <NavLink to="/docs/program-spec" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <Code size={18} /> Program Spec
          </NavLink>
          <NavLink to="/docs/security" className={({isActive}) => `nav-item ${isActive ? 'active' : ''}`} onClick={() => window.innerWidth < 768 && onClose()}>
            <ShieldCheck size={18} /> Security & Tokens
          </NavLink>
        </div>

        <div className="download-section">
          <a href="./downloads/doc-claude-ver.zip" download className="download-btn">
            <Download size={18} /> Download Original
          </a>
        </div>
      </aside>
    </>
  )
}

export default Sidebar
