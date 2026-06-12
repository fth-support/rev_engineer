import { Server, Database, Code, ShieldCheck, MapPin, DatabaseZap, Globe, Zap, FileText } from 'lucide-react'
import { Link } from 'react-router-dom'

import JourneyMap from '../components/JourneyMap'

function Home() {
  return (
    <div className="home-page">
      <div className="hero" style={{ textAlign: 'center', margin: '2rem 0 4rem' }}>
        <h1 style={{ background: 'linear-gradient(90deg, #60a5fa, #a78bfa)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', fontSize: 'clamp(2rem, 5vw, 3.5rem)' }}>
          ServiceTransfer
        </h1>
        <p style={{ fontSize: '1.2rem', color: 'var(--text-secondary)', maxWidth: '600px', margin: '1rem auto' }}>
          Knowledge Pool & Documentation Hub for the ServiceTransfer Revamp Project.
        </p>
      </div>

      <div className="intro-section glass-panel" style={{ padding: '2rem', marginBottom: '4rem' }}>
        <h2>Introduction</h2>
        <p>
          ServiceTransfer คือ Background Agent รุ่นเก่า (พัฒนาด้วย VB6) ที่ทำหน้าที่ในการดึงข้อมูลยอดขายจากเครื่อง POS ของแต่ละสาขา ส่งขึ้นไปยังฐานข้อมูลส่วนกลาง (Head Office) พร้อมทั้งมีกลไกสำคัญในการทำ Tokenization เพื่อปกปิดข้อมูลบัตรเครดิตและบัตรสมาชิกตามมาตรฐานความปลอดภัย
        </p>
        <p>
          เป้าหมายหลักของการ Revamp ระบบนี้คือ การเปลี่ยนสถาปัตยกรรมจากรูปแบบ Polling (วนลูป 500ms) และเชื่อมต่อฐานข้อมูลโดยตรงข้ามเครือข่าย ให้กลายเป็น <strong>Event-Driven Architecture</strong> ควบคู่กับการใช้ <strong>API Gateway</strong> เพื่อเพิ่มความปลอดภัยและประสิทธิภาพ
        </p>
      </div>

      <div className="journey-section" style={{ width: '100%', margin: '0 auto' }}>
        <h2 style={{ textAlign: 'center', marginBottom: '3rem' }}>User & Application Journey</h2>
        <JourneyMap />
      </div>

      <div className="docs-links" style={{ marginTop: '5rem', paddingTop: '3rem', borderTop: '1px solid var(--border-color)' }}>
        <h2 style={{ textAlign: 'center', marginBottom: '2rem' }}>Explore the Documentation</h2>
        <div className="features" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '2rem' }}>
          <Link to="/docs/architecture-srs" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
            <Server size={40} color="#60a5fa" style={{ margin: '0 auto 1rem' }} />
            <h4>Architecture & SRS</h4>
          </Link>

          <Link to="/docs/data-dictionary" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
            <DatabaseZap size={40} color="#34d399" style={{ margin: '0 auto 1rem' }} />
            <h4>Database Schema</h4>
          </Link>

          <Link to="/docs/program-spec" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
            <Code size={40} color="#fbbf24" style={{ margin: '0 auto 1rem' }} />
            <h4>Program Spec</h4>
          </Link>

          <Link to="/docs/security" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
            <ShieldCheck size={40} color="#f87171" style={{ margin: '0 auto 1rem' }} />
            <h4>Security Matrix</h4>
          </Link>
        </div>
      </div>
    </div>
  )
}

export default Home
