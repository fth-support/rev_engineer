import { useState, useEffect } from 'react'
import { Routes, Route } from 'react-router-dom'
import { Menu } from 'lucide-react'
import Sidebar from './components/Sidebar'
import Home from './pages/Home'
import SystemFlows from './pages/SystemFlows'
import MarkdownViewer from './components/MarkdownViewer'
import Downloads from './pages/Downloads'
import './App.css'

function App() {
  const [sidebarOpen, setSidebarOpen] = useState(false)

  // Auto-open sidebar on large screens, close on small screens
  useEffect(() => {
    const handleResize = () => {
      if (window.innerWidth >= 1024) {
        setSidebarOpen(true)
      } else {
        setSidebarOpen(false)
      }
    }
    
    // Set initial state
    handleResize()
    
    window.addEventListener('resize', handleResize)
    return () => window.removeEventListener('resize', handleResize)
  }, [])

  const toggleSidebar = () => {
    setSidebarOpen(!sidebarOpen)
  }

  return (
    <div className="app-container">
      <Sidebar isOpen={sidebarOpen} onClose={() => setSidebarOpen(false)} />
      
      {/* Fixed Hamburger Menu at Top Left */}
      <button 
        className={`hamburger-btn-fixed ${sidebarOpen ? 'hidden-on-desktop' : ''}`} 
        onClick={toggleSidebar}
      >
        <Menu size={28} />
      </button>

      <main className={`main-content ${sidebarOpen ? 'sidebar-open' : ''}`}>
        <div className="content-area glass-panel">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/flows" element={<SystemFlows />} />
            <Route path="/downloads" element={<Downloads />} />
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
