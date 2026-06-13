import { ROLE_VARS } from '../../data/diagramData'

// A rectangular graph node (title bar + bullet items). Used by the architecture graphs.
function GraphNode({ node, isSelected, onSelect }) {
  const base = ROLE_VARS[node.role] || 'role-config'
  const style = {
    '--role-color': `var(--${base})`,
    '--role-soft': `var(--${base}-soft)`,
    '--gnode-w': `${node.w || 220}px`,
  }
  return (
    <button
      type="button"
      className={`gnode ${isSelected ? 'selected' : ''}`}
      style={style}
      onClick={() => onSelect(node.id)}
      aria-pressed={isSelected}
    >
      <div className="gnode__head">
        <span className="gnode__badge" />
        <div className="gnode__titles">
          <div className="gnode__title">{node.title}</div>
          {node.subtitle && <div className="gnode__sub">{node.subtitle}</div>}
        </div>
      </div>
      {node.items && (
        <ul className="gnode__items">
          {node.items.map((it, i) => <li key={i}>{it}</li>)}
        </ul>
      )}
    </button>
  )
}

export default GraphNode
