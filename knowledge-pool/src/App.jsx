import { Routes, Route } from 'react-router-dom'
import Layout from './components/layout/Layout'
import Home from './pages/Home'
import Methodology from './pages/Methodology'
import SystemFlows from './pages/SystemFlows'
import Downloads from './pages/Downloads'
import MarkdownViewer from './components/MarkdownViewer'

function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/methodology" element={<Methodology />} />
        <Route path="/flows" element={<SystemFlows />} />
        <Route path="/downloads" element={<Downloads />} />
        <Route path="/docs/architecture-srs" element={<MarkdownViewer file="01_System_Architecture_and_SRS.md" />} />
        <Route path="/docs/data-dictionary" element={<MarkdownViewer file="02_Data_Dictionary.md" />} />
        <Route path="/docs/program-spec" element={<MarkdownViewer file="03_Program_Specification.md" />} />
        <Route path="/docs/security" element={<MarkdownViewer file="04_Tokenization_and_Security.md" />} />
        <Route path="/docs/website-guide" element={<MarkdownViewer file="05_Website_Development_Guide.md" />} />
      </Routes>
    </Layout>
  )
}

export default App
