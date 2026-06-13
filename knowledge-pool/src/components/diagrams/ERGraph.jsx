import { Database } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import NodeGraph from '../primitives/NodeGraph'
import { erDiagram, ROLE_VARS } from '../../data/diagramData'

const TYPE_CLASS = (t) => {
  const v = (t || '').toLowerCase()
  if (v.includes('pk')) return 'pk'
  if (v.includes('fk')) return 'fk'
  if (v.includes('token')) return 'token'
  return ''
}

function EntityCard({ node, isSelected, onSelect }) {
  const base = ROLE_VARS[node.role] || 'role-config'
  const style = { '--role-color': `var(--${base})`, '--gnode-w': '208px' }
  return (
    <button
      type="button"
      className={`entity ${isSelected ? 'selected' : ''}`}
      style={style}
      onClick={() => onSelect(node.id)}
      aria-pressed={isSelected}
    >
      <div className="entity__head">
        <span className="entity__name">{node.title}</span>
        <span className="entity__sub">{node.sub}</span>
      </div>
      <ul className="entity__fields">
        {node.fields.map((f, i) => (
          <li key={i}>
            <span className="entity__fname">{f.name}</span>
            {f.type && <span className={`entity__ftype ${TYPE_CLASS(f.type)}`}>{f.type}</span>}
          </li>
        ))}
      </ul>
    </button>
  )
}

function ERGraph() {
  const getDetail = (node) => ({
    tag: node.sub,
    desc: `${node.title} — ${node.fields.length} columns. Fields marked TOKEN are tokenized via SafeNet before leaving the branch.`,
    facts: node.fields.map((f) => (f.type ? `${f.name} · ${f.type}` : f.name)),
  })

  return (
    <DiagramFrame
      icon={<Database size={20} color="var(--accent)" />}
      title="Entity-Relationship Diagram"
      subtitle="Core sales, member, and tokenization tables (verified schema field names)"
      hint="Click a table to inspect its columns"
    >
      <NodeGraph
        nodes={erDiagram.entities}
        edges={erDiagram.edges}
        aspect={1.45}
        ariaLabel="ER diagram"
        renderNode={(p) => <EntityCard {...p} />}
        getDetail={getDetail}
      />
    </DiagramFrame>
  )
}

export default ERGraph
