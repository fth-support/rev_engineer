import { Rocket } from 'lucide-react'
import DiagramFrame from './DiagramFrame'
import NodeGraph from '../primitives/NodeGraph'
import GraphNode from '../primitives/GraphNode'
import { targetArchitecture } from '../../data/diagramData'

function TargetArchitectureGraph() {
  return (
    <DiagramFrame
      icon={<Rocket size={20} color="var(--accent)" />}
      title="Target (Revamp) Architecture"
      subtitle="Event-driven agent → message broker → API gateway with transactional writes"
      hint="Click a node to explore the modernization layer"
    >
      <NodeGraph
        nodes={targetArchitecture.nodes}
        edges={targetArchitecture.edges}
        aspect={1.8}
        ariaLabel="Target architecture graph"
        renderNode={(p) => <GraphNode {...p} />}
        getDetail={(node) => node.detail}
      />
    </DiagramFrame>
  )
}

export default TargetArchitectureGraph
