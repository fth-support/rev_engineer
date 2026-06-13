import { FileText, Presentation, Globe, Download, HardDrive, Package } from 'lucide-react'
import './pages.css'

const BASE = '/doc-claude-ver'

const groups = [
  {
    title: 'Specification Documents',
    icon: FileText,
    items: [
      { name: '01_SRS_Document.docx', ext: 'DOCX', path: `${BASE}/01_SRS_Document.docx`, view: false },
      { name: '02_Data_Dictionary.docx', ext: 'DOCX', path: `${BASE}/02_Data_Dictionary.docx`, view: false },
      { name: '03_Program_Specification.docx', ext: 'DOCX', path: `${BASE}/03_Program_Specification.docx`, view: false },
    ],
  },
  {
    title: 'Presentation',
    icon: Presentation,
    items: [
      { name: 'ServiceTransfer_Complete_Presentation.pptx', ext: 'PPTX', path: `${BASE}/ServiceTransfer_Complete_Presentation.pptx`, view: false },
    ],
  },
  {
    title: 'Interactive HTML Diagrams',
    icon: Globe,
    items: [
      { name: '00_Index.html', ext: 'HTML', path: `${BASE}/Diagrams/00_Index.html`, view: true },
      { name: '01_ER_Diagram.html', ext: 'HTML', path: `${BASE}/Diagrams/01_ER_Diagram.html`, view: true },
      { name: '02_System_Architecture.html', ext: 'HTML', path: `${BASE}/Diagrams/02_System_Architecture.html`, view: true },
      { name: '03_Sync_Flow.html', ext: 'HTML', path: `${BASE}/Diagrams/03_Sync_Flow.html`, view: true },
      { name: '04_Tokenization_Flow.html', ext: 'HTML', path: `${BASE}/Diagrams/04_Tokenization_Flow.html`, view: true },
      { name: '05_Member_Points_Flow.html', ext: 'HTML', path: `${BASE}/Diagrams/05_Member_Points_Flow.html`, view: true },
    ],
  },
]

function Row({ item }) {
  return (
    <div className="dl-row">
      <span className="dl-row__name">
        <span className="ext">{item.ext}</span>
        {item.name}
      </span>
      <span className="dl-actions">
        {item.view && (
          <a className="dl-btn" href={item.path} target="_blank" rel="noreferrer"><Globe size={14} /> View</a>
        )}
        <a className="dl-btn primary" href={item.path} download><Download size={14} /> Download</a>
      </span>
    </div>
  )
}

function Downloads() {
  return (
    <div className="downloads">
      <div className="dl-head glass">
        <HardDrive size={30} color="var(--accent)" />
        <div>
          <h2>Download Deliverables</h2>
          <p>The Claude documentation set — Word specs, the slide deck, and interactive HTML diagrams.</p>
        </div>
      </div>

      <div className="dl-group glass">
        <h3><Package size={18} color="var(--accent)" /> Everything (zip archive)</h3>
        <div className="dl-row">
          <span className="dl-row__name"><span className="ext">ZIP</span> doc-claude-ver.zip</span>
          <span className="dl-actions">
            <a className="dl-btn primary" href="/downloads/doc-claude-ver.zip" download><Download size={14} /> Download all</a>
          </span>
        </div>
      </div>

      {groups.map(({ title, icon: Icon, items }) => (
        <div key={title} className="dl-group glass">
          <h3><Icon size={18} color="var(--accent)" /> {title}</h3>
          {items.map((item) => <Row key={item.name} item={item} />)}
        </div>
      ))}
    </div>
  )
}

export default Downloads
