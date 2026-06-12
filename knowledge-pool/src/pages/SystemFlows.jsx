import { useState, useEffect, useRef } from 'react';
import mermaid from 'mermaid';
import { Network, Database, RefreshCw, ShieldCheck, CreditCard } from 'lucide-react';
import './SystemFlows.css';

const flows = [
  {
    id: 'architecture',
    title: 'System Architecture',
    icon: <Network size={20} />,
    description: 'ภาพรวมสถาปัตยกรรมของ ServiceTransfer แบบดั้งเดิม (Legacy) และการเชื่อมต่อข้ามระบบ',
    code: `flowchart TD
    subgraph POS_Terminal["POS Terminal (Branch)"]
        POS_App[POS Application]
        LocalDB[(Local DB SQL Express)]
        ST[ServiceTransfer.exe VB6 Agent]
        Config[AdaIni.Ada Access DB Config]
        
        POS_App -->|บันทึกธุรกรรม| LocalDB
        Config -.->|อ่านรหัสผ่าน| ST
        LocalDB <-->|SELECT / UPDATE Flag| ST
    end

    subgraph Head_Office["Head Office (HQ)"]
        CentralDB[(Central Server Online DB)]
        MemberDB[(Member DB)]
    end

    subgraph Security_Layer["Security Layer"]
        SafeNet[SafeNet Tokenizer SOAP Web Service]
    end

    ST == "INSERT/UPDATE (SQLOLEDB.1)" ==> CentralDB
    ST == "UPDATE Points" ==> MemberDB
    ST <-->|"Tokenize (FIRST_SIX)"| SafeNet`
  },
  {
    id: 'sync_flow',
    title: 'Data Sync Flow',
    icon: <RefreshCw size={20} />,
    description: 'กระบวนการดึงข้อมูลจาก Local DB และส่งต่อขึ้นส่วนกลาง',
    code: `sequenceDiagram
    participant C as Customer
    participant P as POS Terminal (Local)
    participant S as ServiceTransfer Agent
    participant T as SafeNet Tokenizer
    participant HQ as Central Server (HQ)
    participant M as Member DB

    C->>P: 1. ซื้อสินค้า / ทำธุรกรรม
    P->>P: 2. บันทึกลง Local DB (Offline Mode)
    loop Every 500 ms
        S->>P: 3. ดึงข้อมูลที่ยังไม่ส่ง (Flag=0)
        S->>T: 4. ส่งเลขบัตรเพื่อสร้าง Token
        T-->>S: คืนค่า Token
        S->>HQ: 5. INSERT/UPDATE ข้อมูลธุรกรรมพร้อม Token
        S->>M: 6. ประมวลผลและอัปเดตคะแนนสมาชิก
    end`
  },
  {
    id: 'er_diagram',
    title: 'ER Diagram (Core Tables)',
    icon: <Database size={20} />,
    description: 'ความสัมพันธ์ของตารางหลักที่ใช้ในการโอนถ่ายข้อมูลและ Token',
    code: `erDiagram
    TPSTSalHD ||--o{ TPSTSalDT : "1 to N via TmnNum TransNo Date"
    TPSTSalHD ||--o{ TPSTSalRC : "1 to N"
    TPSTSalHD ||--o{ TPSTSalCD : "1 to N"
    TPSTSalHD ||--o{ TPSTSalePoint : "1 to N"
    
    TPSTSalePoint }o--|| TCNMMallCard : "Lookup via Token Member ID"
    TCNMMallCard ||--o{ TPSTBPHis : "1 to N History"
    
    TPSTTokenLst ||--o{ TPSTSalHD : "Tokenize CardCode"
    TPSTTokenLst ||--o{ TPSTSalePoint : "Tokenize MemID"`
  },
  {
    id: 'tokenization',
    title: 'Tokenization Flow',
    icon: <ShieldCheck size={20} />,
    description: 'กระบวนการแปลงรหัสบัตรให้เป็น Token ก่อนนำออกจากสาขา',
    code: `flowchart TD
    Start([เริ่มฟังก์ชัน Tokenize]) --> G1{เลขยาว < 7 หลัก?}
    G1 -- Yes --> RetOrig([คืนค่าเดิม])
    G1 -- No --> G2{เลขสั้นกว่า 14 หรือว่าง?}
    G2 -- Yes --> RetOrig
    G2 -- No --> DBSearch[(ค้นหาใน TPSTTokenLst)]
    
    DBSearch --> CheckFound{พบค่า Token เดิม?}
    CheckFound -- Yes --> RetToken([คืนค่า Token เดิม])
    CheckFound -- No --> InitSOAP[Initialize SOAP Client]
    
    InitSOAP --> CallWS[เรียก InsertToken() ผ่าน Web Service]
    CallWS --> ResWS{สำเร็จ?}
    ResWS -- Yes --> RetNew([คืนค่า Token ใหม่ และบันทึกลง DB])
    ResWS -- No --> Error([Error 557: Token Failed])`
  }
];

function SystemFlows() {
  const [activeFlow, setActiveFlow] = useState(flows[0]);
  const mermaidRef = useRef(null);

  useEffect(() => {
    mermaid.initialize({
      startOnLoad: true,
      theme: 'dark',
      fontFamily: 'Inter, sans-serif'
    });
  }, []);

  useEffect(() => {
    if (mermaidRef.current) {
      mermaidRef.current.innerHTML = '';
      mermaid.render(`mermaid-${activeFlow.id}`, activeFlow.code).then(({ svg }) => {
        mermaidRef.current.innerHTML = svg;
      });
    }
  }, [activeFlow]);

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

        <div className="flows-content glass-panel">
          <div className="flow-content-header">
            <h3>{activeFlow.title}</h3>
            <p>{activeFlow.description}</p>
          </div>
          <div className="flow-render-area" ref={mermaidRef}>
            {/* SVG is rendered here */}
          </div>
        </div>
      </div>
    </div>
  );
}

export default SystemFlows;
