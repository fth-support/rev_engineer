import { useState, useEffect } from 'react'
import Sidebar from './Sidebar'
import Topbar from './Topbar'

function Layout({ children }) {
  const [sidebarOpen, setSidebarOpen] = useState(false)

  useEffect(() => {
    const handleResize = () => setSidebarOpen(window.innerWidth >= 1024)
    handleResize()
    window.addEventListener('resize', handleResize)
    return () => window.removeEventListener('resize', handleResize)
  }, [])

  return (
    <div className="app-shell">
      <Sidebar isOpen={sidebarOpen} onClose={() => setSidebarOpen(false)} />
      <div className={`app-main ${sidebarOpen ? 'with-sidebar' : ''}`}>
        <Topbar onToggleSidebar={() => setSidebarOpen((o) => !o)} />
        <main className="app-content">{children}</main>
      </div>
    </div>
  )
}

export default Layout
