import { useState } from 'react'
import { Routes, Route } from 'react-router-dom'
import { Menu } from 'lucide-react'
import Sidebar from './components/Sidebar'
import Home from './pages/Home'
import Journey from './pages/Journey'
import MarkdownViewer from './components/MarkdownViewer'
import './App.css'

function App() {
  const [sidebarOpen, setSidebarOpen] = useState(false)

  const toggleSidebar = () => {
    setSidebarOpen(!sidebarOpen)
  }

  return (
    <div className="app-container">
      <Sidebar isOpen={sidebarOpen} onClose={() => setSidebarOpen(false)} />
      
      <main className={`main-content ${sidebarOpen ? 'sidebar-open' : ''}`}>
        <header className="top-header">
          <button className="hamburger-btn" onClick={toggleSidebar}>
            <Menu size={28} />
          </button>
          <h2>Knowledge Pool</h2>
        </header>

        <div className="content-area glass-panel">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/journey" element={<Journey />} />
            <Route path="/docs/architecture-srs" element={<MarkdownViewer file="01_System_Architecture_and_SRS.md" />} />
            <Route path="/docs/data-dictionary" element={<MarkdownViewer file="02_Data_Dictionary.md" />} />
            <Route path="/docs/program-spec" element={<MarkdownViewer file="03_Program_Specification.md" />} />
            <Route path="/docs/security" element={<MarkdownViewer file="04_Tokenization_and_Security.md" />} />
          </Routes>
        </div>
      </main>
    </div>
  )
}

export default App
