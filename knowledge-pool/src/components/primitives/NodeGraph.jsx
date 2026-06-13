import { useMemo, useRef, useState, useLayoutEffect, useCallback } from 'react'
import AnimatedConnections from './AnimatedConnections'
import DetailPanel from './DetailPanel'

// Reusable interactive graph: animated connectors (behind) + absolutely-positioned
// nodes (in front). Clicking a node highlights it + its neighbours and dims the rest,
// then opens a DetailPanel. Powers ArchitectureGraph / TargetArchitectureGraph / ERGraph.
//
// Node centers are clamped so each node's full bounding box stays inside the canvas
// (no clipping at the edges), and connectors are drawn to those clamped centers.

const MARGIN = 8

function NodeGraph({ nodes, edges, aspect = 1.6, renderNode, getDetail, ariaLabel = 'Diagram' }) {
  const ref = useRef(null)
  const nodeRefs = useRef({})
  const [size, setSize] = useState({ width: 0, height: 0 })
  const [sizes, setSizes] = useState({}) // id -> { w, h } measured in px
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

  // Measure each node's rendered size once the canvas has a size.
  useLayoutEffect(() => {
    const m = {}
    let changed = false
    for (const n of nodes) {
      const el = nodeRefs.current[n.id]
      if (el) {
        m[n.id] = { w: el.offsetWidth, h: el.offsetHeight }
        if (!sizes[n.id] || sizes[n.id].w !== m[n.id].w || sizes[n.id].h !== m[n.id].h) changed = true
      }
    }
    if (changed) setSizes(m)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nodes, size.width, size.height])

  // Clamp each node's center so its box stays fully inside the canvas.
  const centers = useMemo(() => {
    const out = {}
    const W = size.width
    const H = size.height
    for (const n of nodes) {
      const sz = sizes[n.id] || { w: n.w || 210, h: 130 }
      const hw = sz.w / 2
      const hh = sz.h / 2
      let cx = (n.x / 100) * W
      let cy = (n.y / 100) * H
      if (W) cx = Math.max(hw + MARGIN, Math.min(W - hw - MARGIN, cx))
      if (H) cy = Math.max(hh + MARGIN, Math.min(H - hh - MARGIN, cy))
      out[n.id] = { cx, cy }
    }
    return out
  }, [nodes, sizes, size])

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
          centers={centers}
          edges={edges}
          width={size.width}
          height={size.height}
          selectedId={selectedId}
          relatedIds={relatedIds}
        />
        {nodes.map((node) => {
          const isSelected = node.id === selectedId
          const isDimmed = selectedId ? !relatedIds.has(node.id) : false
          const c = centers[node.id]
          return (
            <div
              key={node.id}
              ref={(el) => { nodeRefs.current[node.id] = el }}
              className="ngraph__node"
              style={{ left: `${c ? c.cx : 0}px`, top: `${c ? c.cy : 0}px`, opacity: isDimmed ? 0.32 : 1 }}
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
