import { useMemo, useRef, useState, useLayoutEffect, useCallback } from 'react'
import AnimatedConnections from './AnimatedConnections'
import DetailPanel from './DetailPanel'

// Reusable interactive graph: animated connectors (behind) + absolutely-positioned
// nodes (in front). Clicking a node highlights it + its neighbours and dims the rest,
// then opens a DetailPanel. Powers ArchitectureGraph / TargetArchitectureGraph / ERGraph.

function NodeGraph({ nodes, edges, aspect = 1.6, renderNode, getDetail, ariaLabel = 'Diagram' }) {
  const ref = useRef(null)
  const [size, setSize] = useState({ width: 0, height: 0 })
  const [selectedId, setSelectedId] = useState(null)

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const measure = () => {
      const r = el.getBoundingClientRect()
      setSize({ width: r.width, height: r.height })
    }
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  const relatedMap = useMemo(() => {
    const m = {}
    for (const n of nodes) {
      const s = new Set([n.id])
      for (const e of edges) {
        if (e.from === n.id) s.add(e.to)
        if (e.to === n.id) s.add(e.from)
      }
      m[n.id] = s
    }
    return m
  }, [nodes, edges])

  const relatedIds = selectedId ? relatedMap[selectedId] : null

  const onSelect = useCallback((id) => {
    setSelectedId((prev) => (prev === id ? null : id))
  }, [])

  const selectedNode = selectedId ? nodes.find((n) => n.id === selectedId) : null
  const detail = selectedNode && getDetail ? getDetail(selectedNode) : null

  return (
    <div className="ngraph" aria-label={ariaLabel}>
      <div className="ngraph__canvas" ref={ref} style={{ aspectRatio: String(aspect) }}>
        <AnimatedConnections
          nodes={nodes}
          edges={edges}
          width={size.width}
          height={size.height}
          selectedId={selectedId}
          relatedIds={relatedIds}
        />
        {nodes.map((node) => {
          const isSelected = node.id === selectedId
          const isDimmed = selectedId ? !relatedIds.has(node.id) : false
          return (
            <div
              key={node.id}
              className="ngraph__node"
              style={{ left: `${node.x}%`, top: `${node.y}%`, opacity: isDimmed ? 0.32 : 1 }}
            >
              {renderNode({ node, isSelected, isDimmed, onSelect })}
            </div>
          )
        })}
      </div>
      {detail && (
        <DetailPanel node={selectedNode} detail={detail} onClose={() => setSelectedId(null)} />
      )}
    </div>
  )
}

export default NodeGraph
