import { Server, Database, Code, ShieldCheck, MapPin, DatabaseZap, Globe, Zap, FileText } from 'lucide-react'
import { Link } from 'react-router-dom'

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

      <div className="journey-section">
        <h2 style={{ textAlign: 'center', marginBottom: '3rem' }}>User & Application Journey</h2>
        
        <div className="timeline">
          {/* Node 1 */}
          <div className="timeline-item">
            <div className="timeline-icon" style={{ background: '#3b82f6' }}>
              <MapPin size={24} color="#fff" />
            </div>
            <div className="timeline-content glass-panel">
              <h3>1. Transaction Capture</h3>
              <p>ลูกค้าชำระเงินที่เคาน์เตอร์ POS สาขา ข้อมูลธุรกรรมจะถูกบันทึกลงในฐานข้อมูล <strong>Local SQL Express</strong> ทันที (Offline First) โดยเรคอร์ดจะถูกกำหนด Flag เบื้องต้นเป็น <code>FTStaSentOnOff = '0'</code> (รอดำเนินการ)</p>
            </div>
          </div>

          {/* Node 2 */}
          <div className="timeline-item">
            <div className="timeline-icon" style={{ background: '#8b5cf6' }}>
              <Zap size={24} color="#fff" />
            </div>
            <div className="timeline-content glass-panel">
              <h3>2. Background Sync Agent</h3>
              <p>โปรแกรม ServiceTransfer ที่ทำงานอยู่เบื้องหลัง จะทำ Timer Polling ทุก ๆ 500 มิลลิวินาที ตรวจสอบตารางทั้งหมด หากพบ Flag = '0' จะดึงข้อมูลขึ้นมาเตรียมส่ง โดยรอให้ส่วนประกอบย่อยของบิล (เช่น รายการสินค้า, การชำระ) ครบถ้วนก่อนส่ง Header เสมอ (HD-First Rule)</p>
            </div>
          </div>

          {/* Node 3 */}
          <div className="timeline-item">
            <div className="timeline-icon" style={{ background: '#f59e0b' }}>
              <ShieldCheck size={24} color="#fff" />
            </div>
            <div className="timeline-content glass-panel">
              <h3>3. Data Tokenization</h3>
              <p>ก่อนส่งข้อมูลออกจากสาขา ข้อมูลที่เป็นความลับ (เลขบัตรเครดิต, เลขสมาชิก) จะถูกดักจับและส่งผ่าน SafeNet SOAP Web Service เพื่อแปลงให้เป็น <strong>Token (First 6 digits)</strong> ป้องกันไม่ให้ข้อมูลจริงรั่วไหลระหว่างการเดินทางข้ามเครือข่าย</p>
            </div>
          </div>

          {/* Node 4 */}
          <div className="timeline-item">
            <div className="timeline-icon" style={{ background: '#10b981' }}>
              <Globe size={24} color="#fff" />
            </div>
            <div className="timeline-content glass-panel">
              <h3>4. Central Delivery & Member Points</h3>
              <p>ข้อมูลที่ถูก Tokenize แล้วจะถูกต่อข้อความเป็นคำสั่ง SQL INSERT/UPDATE แล้วยิงไปยัง <strong>Central DB (Head Office)</strong> และมีการนำคะแนนการซื้อไปคำนวณสะสมให้สมาชิกในระบบ <strong>Member DB (MallCard)</strong> ต่อไป</p>
            </div>
          </div>
        </div>
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
