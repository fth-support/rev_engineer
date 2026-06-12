import { Server, Database, Code, ShieldCheck, MapPin, DatabaseZap, Globe, Zap, FileText, Activity, Layers, Terminal } from 'lucide-react'
import { Link } from 'react-router-dom'
import JourneyMap from '../components/JourneyMap'
import '../Bento.css'

function Home() {
  return (
    <div className="home-page" style={{ width: '100%', maxWidth: '1400px', margin: '0 auto' }}>
      <div className="bento" aria-label="Bento grid layout">
        
        {/* HERO CELL */}
        <div className="cell cell--hero">
          <span className="cell__tag">HERO</span>
          <h1>ServiceTransfer</h1>
          <p style={{ fontSize: '1.2rem', color: 'var(--text-secondary)' }}>
            Knowledge Pool & Documentation Hub for the ServiceTransfer Revamp Project.
          </p>
          <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem' }}>
            <span style={{ background: 'rgba(96, 165, 250, 0.2)', color: '#60a5fa', padding: '0.4rem 1rem', borderRadius: '20px', fontSize: '0.85rem', fontWeight: '600' }}>Event-Driven</span>
            <span style={{ background: 'rgba(52, 211, 153, 0.2)', color: '#34d399', padding: '0.4rem 1rem', borderRadius: '20px', fontSize: '0.85rem', fontWeight: '600' }}>API Gateway</span>
          </div>
        </div>

        {/* ABOUT CELL */}
        <div className="cell cell--about">
          <span className="cell__tag">ABOUT</span>
          <h2>Introduction</h2>
          <p style={{ marginTop: '0.5rem' }}>
            ServiceTransfer คือ Background Agent ที่ทำหน้าที่ดึงข้อมูลยอดขายจากเครื่อง POS สาขา ส่งขึ้นไปยังฐานข้อมูลส่วนกลาง (Head Office) พร้อมกลไก Tokenization เพื่อปกปิดข้อมูลบัตรเครดิต
          </p>
        </div>

        {/* STATS CELL */}
        <div className="cell cell--stats">
          <span className="cell__tag">STATS</span>
          <Activity size={48} className="cell-icon" color="#fbbf24" />
          <h2 style={{ fontSize: '3rem', margin: '0.5rem 0' }}>500<span style={{ fontSize: '1rem', color: 'var(--text-secondary)' }}>ms</span></h2>
          <p>Legacy Polling Time to be replaced by Real-time Events</p>
        </div>

        {/* JOURNEY / FEAT CELL */}
        <div className="cell cell--feat" style={{ padding: 0, overflow: 'hidden', minHeight: '400px' }}>
          <span className="cell__tag" style={{ left: '20px', top: 0, borderRadius: '0 0 8px 8px' }}>JOURNEY</span>
          <div style={{ padding: '3rem 1.5rem 1.5rem 1.5rem', width: '100%' }}>
            <JourneyMap />
          </div>
        </div>

        {/* PROMO / FLOWS CELL */}
        <Link to="/flows" className="cell cell--promo doc-link-card" style={{ textDecoration: 'none' }}>
          <span className="cell__tag">SYSTEM FLOWS</span>
          <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem' }}>
            <Layers size={56} className="cell-icon" color="#fbbf24" style={{ margin: 0 }} />
            <div>
              <h3>Interactive Diagrams</h3>
              <p style={{ marginTop: '0.5rem' }}>Explore ER Diagrams, Architecture, and Tokenization Flows directly.</p>
            </div>
          </div>
        </Link>

        {/* MEDIA / ARCHITECTURE DOC */}
        <Link to="/docs/architecture-srs" className="cell cell--media doc-link-card">
          <span className="cell__tag">DOCS</span>
          <Server size={48} className="cell-icon" color="#34d399" />
          <h4>Architecture & SRS</h4>
        </Link>

        {/* FEED / DATA DICT DOC */}
        <Link to="/docs/data-dictionary" className="cell cell--feed doc-link-card">
          <span className="cell__tag">DOCS</span>
          <DatabaseZap size={48} className="cell-icon" color="#34d399" />
          <h4>Database Schema</h4>
        </Link>

        {/* TASKS / PROGRAM SPEC DOC */}
        <Link to="/docs/program-spec" className="cell cell--tasks doc-link-card">
          <span className="cell__tag">DOCS</span>
          <Code size={48} className="cell-icon" color="#60a5fa" />
          <h4>Program Spec</h4>
        </Link>

        {/* CARD / SECURITY DOC */}
        <Link to="/docs/security" className="cell cell--card doc-link-card">
          <span className="cell__tag">DOCS</span>
          <ShieldCheck size={48} className="cell-icon" color="#f87171" />
          <h4>Security Matrix</h4>
        </Link>

      </div>
    </div>
  )
}

export default Home
