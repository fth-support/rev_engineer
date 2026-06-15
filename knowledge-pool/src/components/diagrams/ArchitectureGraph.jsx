import { Network } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import NodeGraph from '../primitives/NodeGraph'
import GraphNode from '../primitives/GraphNode'
import { architecture } from '../../data/diagramData'

function ArchitectureGraph() {
  return (
    <DiagramFrame
      icon={<Network size={20} color="var(--accent)" />}
      title="System Architecture"
      subtitle="POSFront (producer) + ServiceTransfer (sync) — 3-database topology + SafeNet tokenizer"
      hint="Click a node to highlight its connections and see details"
    >
      <NodeGraph
        nodes={architecture.nodes}
        edges={architecture.edges}
        aspect={1.5}
        ariaLabel="System architecture graph"
        renderNode={(p) => <GraphNode {...p} />}
        getDetail={(node) => node.detail}
      />
    </DiagramFrame>
  )
}

export default ArchitectureGraph
