import { X } from 'lucide-react'
import { ROLE_VARS } from '../../data/diagramData'

// Slide-up detail panel shown when a graph node is selected.
// Adapted from ag-kit-main/uxui/src/components/DetailPanel.tsx.
function DetailPanel({ node, detail, onClose }) {
  const roleVar = `var(--${ROLE_VARS[node.role] || 'role-config'})`
  return (
    <div className="detail-panel" role="region" aria-label={`Details for ${node.title}`}>
      <div className="detail-panel__head">
        <div className="detail-panel__title">
          <span className="detail-panel__dot" style={{ background: roleVar }} />
          <strong>{node.title}</strong>
          {detail.tag && <span className="detail-panel__tag" style={{ color: roleVar }}>{detail.tag}</span>}
        </div>
        <button type="button" className="detail-panel__close" onClick={onClose} aria-label="Close details">
          <X size={16} />
        </button>
      </div>
      {detail.desc && <p className="detail-panel__desc">{detail.desc}</p>}
      {detail.facts && detail.facts.length > 0 && (
        <ul className="detail-panel__facts">
          {detail.facts.map((f, i) => (
            <li key={i}><code>{f}</code></li>
          ))}
        </ul>
      )}
    </div>
  )
}

export default DetailPanel
