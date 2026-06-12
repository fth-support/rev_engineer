import { ShieldCheck, Server, Database, Code } from 'lucide-react'
import { Link } from 'react-router-dom'

function Home() {
  return (
    <div className="home-page">
      <div className="hero" style={{ textAlign: 'center', margin: '2rem 0 4rem' }}>
        <h1 style={{ background: 'linear-gradient(90deg, #60a5fa, #a78bfa)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', fontSize: '3.5rem' }}>
          ServiceTransfer
        </h1>
        <p style={{ fontSize: '1.2rem', color: 'var(--text-secondary)', maxWidth: '600px', margin: '1rem auto' }}>
          Knowledge Pool & Documentation Hub for the Revamp Project
        </p>
      </div>

      <div className="features" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '2rem' }}>
        <Link to="/docs/architecture-srs" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
          <Server size={48} color="#60a5fa" style={{ margin: '0 auto 1rem' }} />
          <h3>Architecture</h3>
          <p>Target Event-Driven Architecture & SRS</p>
        </Link>

        <Link to="/docs/data-dictionary" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
          <Database size={48} color="#34d399" style={{ margin: '0 auto 1rem' }} />
          <h3>Database Schema</h3>
          <p>11 Core Tables and Data Dictionary</p>
        </Link>

        <Link to="/docs/program-spec" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
          <Code size={48} color="#fbbf24" style={{ margin: '0 auto 1rem' }} />
          <h3>Program Spec</h3>
          <p>Detailed logic and Pseudocode</p>
        </Link>

        <Link to="/docs/security" className="feature-card glass-panel" style={{ padding: '2rem', textAlign: 'center', transition: 'transform 0.2s', display: 'block' }}>
          <ShieldCheck size={48} color="#f87171" style={{ margin: '0 auto 1rem' }} />
          <h3>Security</h3>
          <p>Tokenization flow and PCI-DSS compliance</p>
        </Link>
      </div>
    </div>
  )
}

export default Home
