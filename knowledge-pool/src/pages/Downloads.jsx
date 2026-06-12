import { useState } from 'react';
import { Folder, FileText, Globe, Download, ChevronRight, ChevronDown, HardDrive } from 'lucide-react';

const fileTree = [
  {
    name: 'Diagrams',
    type: 'folder',
    children: [
      { name: '00_Index.html', type: 'html', path: '/doc-claude-ver/Diagrams/00_Index.html' },
      { name: '01_ER_Diagram.html', type: 'html', path: '/doc-claude-ver/Diagrams/01_ER_Diagram.html' },
      { name: '02_System_Architecture.html', type: 'html', path: '/doc-claude-ver/Diagrams/02_System_Architecture.html' },
      { name: '03_Sync_Flow.html', type: 'html', path: '/doc-claude-ver/Diagrams/03_Sync_Flow.html' },
      { name: '04_Tokenization_Flow.html', type: 'html', path: '/doc-claude-ver/Diagrams/04_Tokenization_Flow.html' },
      { name: '05_Member_Points_Flow.html', type: 'html', path: '/doc-claude-ver/Diagrams/05_Member_Points_Flow.html' }
    ]
  },
  { name: '01_System_Architecture_and_SRS.md', type: 'md', path: '/doc-claude-ver/01_System_Architecture_and_SRS.md' },
  { name: '02_Data_Dictionary.md', type: 'md', path: '/doc-claude-ver/02_Data_Dictionary.md' },
  { name: '03_Program_Specification.md', type: 'md', path: '/doc-claude-ver/03_Program_Specification.md' },
  { name: '04_Tokenization_and_Security.md', type: 'md', path: '/doc-claude-ver/04_Tokenization_and_Security.md' }
];

const FileItem = ({ item }) => {
  if (item.type === 'folder') {
    return <FolderItem folder={item} />;
  }

  const getIcon = () => {
    if (item.type === 'html') return <Globe size={18} color="#60a5fa" />;
    return <FileText size={18} color="#34d399" />;
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0.75rem 1rem', borderBottom: '1px solid rgba(255,255,255,0.05)', transition: 'background 0.2s' }} className="file-item-hover">
      <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
        {getIcon()}
        <span style={{ fontSize: '0.95rem' }}>{item.name}</span>
      </div>
      <div style={{ display: 'flex', gap: '0.5rem' }}>
        <a href={item.path} target="_blank" rel="noreferrer" style={{ padding: '0.4rem 0.8rem', borderRadius: '6px', background: 'rgba(255,255,255,0.1)', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.4rem', color: 'var(--text-primary)' }}>
          <Globe size={14} /> View
        </a>
        <a href={item.path} download style={{ padding: '0.4rem 0.8rem', borderRadius: '6px', background: 'var(--accent)', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.4rem', color: '#fff' }}>
          <Download size={14} /> Download
        </a>
      </div>
    </div>
  );
};

const FolderItem = ({ folder }) => {
  const [isOpen, setIsOpen] = useState(true);

  return (
    <div style={{ marginBottom: '0.5rem' }}>
      <div 
        onClick={() => setIsOpen(!isOpen)}
        style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', padding: '0.75rem 1rem', background: 'rgba(255,255,255,0.03)', cursor: 'pointer', borderRadius: '8px' }}
      >
        {isOpen ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
        <Folder size={18} color="#fbbf24" fill="#fbbf24" fillOpacity={0.2} />
        <strong style={{ fontSize: '1rem' }}>{folder.name}</strong>
      </div>
      {isOpen && (
        <div style={{ paddingLeft: '1.5rem', borderLeft: '1px solid rgba(255,255,255,0.1)', marginLeft: '1.5rem', marginTop: '0.5rem' }}>
          {folder.children.map((child, idx) => (
            <FileItem key={idx} item={child} />
          ))}
        </div>
      )}
    </div>
  );
};

function Downloads() {
  return (
    <div className="downloads-page" style={{ width: '100%', maxWidth: '1000px', margin: '0 auto', padding: '2rem 1rem' }}>
      <style>{`
        .file-item-hover:hover {
          background: rgba(255, 255, 255, 0.03);
        }
      `}</style>
      
      <div className="glass-panel" style={{ padding: '2rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '2rem', paddingBottom: '1rem', borderBottom: '1px solid rgba(255,255,255,0.1)' }}>
          <HardDrive size={32} color="#a78bfa" />
          <div>
            <h2 style={{ margin: 0, border: 'none', padding: 0 }}>File Explorer</h2>
            <p style={{ margin: 0, fontSize: '0.9rem', color: 'var(--text-secondary)', marginTop: '0.25rem' }}>
              Browse and download original Claude-generated artifacts
            </p>
          </div>
        </div>

        <div className="explorer-container" style={{ background: '#0f172a', borderRadius: '12px', border: '1px solid rgba(255,255,255,0.1)', padding: '1rem' }}>
          <div style={{ padding: '0.5rem 1rem', borderBottom: '1px solid rgba(255,255,255,0.1)', marginBottom: '1rem', color: 'var(--text-secondary)', fontSize: '0.85rem', display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
            <HardDrive size={14} /> / doc-claude-ver /
          </div>
          
          {fileTree.map((item, idx) => (
            <FileItem key={idx} item={item} />
          ))}
        </div>
      </div>
    </div>
  );
}

export default Downloads;
