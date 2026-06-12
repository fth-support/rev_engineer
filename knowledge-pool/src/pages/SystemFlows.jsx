import { useState } from 'react';
import { Network, Database, RefreshCw, ShieldCheck, CreditCard } from 'lucide-react';
import MarkdownViewer from '../components/MarkdownViewer';
import './SystemFlows.css';

const flows = [
  {
    id: 'architecture',
    title: 'System Architecture',
    icon: <Network size={20} />,
    description: 'ภาพรวมสถาปัตยกรรมของ ServiceTransfer แบบดั้งเดิม (Legacy) และการเชื่อมต่อข้ามระบบ',
    relatedDoc: '01_System_Architecture_and_SRS.md',
    iframeUrl: '/doc-claude-ver/Diagrams/02_System_Architecture.html'
  },
  {
    id: 'sync_flow',
    title: 'Data Sync Flow',
    icon: <RefreshCw size={20} />,
    description: 'กระบวนการดึงข้อมูลจาก Local DB และส่งต่อขึ้นส่วนกลาง',
    relatedDoc: '03_Program_Specification.md',
    iframeUrl: '/doc-claude-ver/Diagrams/03_Sync_Flow.html'
  },
  {
    id: 'er_diagram',
    title: 'ER Diagram (Core Tables)',
    icon: <Database size={20} />,
    description: 'ความสัมพันธ์ของตารางหลักที่ใช้ในการโอนถ่ายข้อมูลและ Token',
    relatedDoc: '02_Data_Dictionary.md',
    iframeUrl: '/doc-claude-ver/Diagrams/01_ER_Diagram.html'
  },
  {
    id: 'tokenization',
    title: 'Tokenization Flow',
    icon: <ShieldCheck size={20} />,
    description: 'กระบวนการแปลงรหัสบัตรให้เป็น Token ก่อนนำออกจากสาขา',
    relatedDoc: '04_Tokenization_and_Security.md',
    iframeUrl: '/doc-claude-ver/Diagrams/04_Tokenization_Flow.html'
  },
  {
    id: 'member_points',
    title: 'Member Points Flow',
    icon: <CreditCard size={20} />,
    description: 'กระบวนการสะสมแต้มสมาชิกจากยอดขาย',
    relatedDoc: '03_Program_Specification.md',
    iframeUrl: '/doc-claude-ver/Diagrams/05_Member_Points_Flow.html'
  }
];

function SystemFlows() {
  const [activeFlow, setActiveFlow] = useState(flows[0]);

  return (
    <div className="system-flows-container">
      <div className="flows-header glass-panel">
        <h2>System Flows & Diagrams</h2>
        <p>แหล่งรวบรวม Diagram ทั้งหมดของระบบ เลือกดูแต่ละ Flow ได้จากเมนูด้านซ้าย</p>
      </div>

      <div className="flows-layout">
        <div className="flows-sidebar glass-panel">
          {flows.map((flow) => (
            <button
              key={flow.id}
              className={`flow-nav-btn ${activeFlow.id === flow.id ? 'active' : ''}`}
              onClick={() => setActiveFlow(flow)}
            >
              <span className="icon">{flow.icon}</span>
              {flow.title}
            </button>
          ))}
        </div>

        <div className="flows-content glass-panel" style={{ display: 'flex', flexDirection: 'column' }}>
          <div className="flow-content-header" style={{ flexShrink: 0 }}>
            <h3>{activeFlow.title}</h3>
            <p>{activeFlow.description}</p>
          </div>
          <div className="flow-render-area" style={{ flex: 1, padding: 0, overflow: 'hidden', minHeight: '600px', background: '#f0f4f8', borderRadius: '0 0 16px 16px' }}>
            <iframe 
              src={activeFlow.iframeUrl} 
              style={{ width: '100%', height: '100%', border: 'none' }}
              title={activeFlow.title}
            />
          </div>
        </div>
      </div>

      {activeFlow.relatedDoc && (
        <div className="flow-related-doc glass-panel" style={{ padding: '2rem', marginTop: '2rem' }}>
          <h3 style={{ margin: '0 0 1rem 0', paddingBottom: '0.5rem', borderBottom: '1px solid var(--border-color)', color: '#60a5fa' }}>
            Detailed Documentation
          </h3>
          <MarkdownViewer file={activeFlow.relatedDoc} />
        </div>
      )}
    </div>
  );
}

export default SystemFlows;
