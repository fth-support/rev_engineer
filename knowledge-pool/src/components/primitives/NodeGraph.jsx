import { useMemo, useRef, useState, useLayoutEffect, useCallback } from 'react'
import AnimatedConnections from './AnimatedConnections'
import DetailPanel from './DetailPanel'

// Reusable interactive graph: animated connectors (behind) + absolutely-positioned
// nodes (in front). Clicking a node highlights it + its neighbours and dims the rest.
//
// Responsive strategy (fit-to-width): the graph is laid out on a FIXED "design canvas"
// (DESIGN_W × designH) so node positions, sizes, fonts and connectors keep identical
// proportions at every screen size. A wrapper then `transform: scale()`s the whole
// canvas down to fit the available width — so it's always fully visible, never clipped
// and never overlapping, without needing full screen.

const DESIGN_W = 1080
const MARGIN = 10

function NodeGraph({ nodes, edges, aspect = 1.6, renderNode, getDetail, ariaLabel = 'Diagram' }) {
  const fitRef = useRef(null)
  const nodeRefs = useRef({})
  const [scale, setScale] = useState(1)
  const [sizes, setSizes] = useState({}) // id -> { w, h } measured at design scale
  const [selectedId, setSelectedId] = useState(null)

  const designH = Math.round(DESIGN_W / aspect)

  // Scale the fixed design canvas to fit the available width.
  useLayoutEffect(() => {
    const el = fitRef.current
    if (!el) return
    const measure = () => {
      const w = el.clientWidth
      if (w > 0) setScale(Math.min(1, w / DESIGN_W))
    }
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  // Measure each node's intrinsic (unscaled) size for clamping. Re-measure after the
  // webfont loads and whenever a node resizes — otherwise the clamp uses fallback-font
  // heights measured before Inter loads, and tall nodes overflow the top edge.
  useLayoutEffect(() => {
    const remeasure = () => {
      const m = {}
      for (const n of nodes) {
        const el = nodeRefs.current[n.id]
        if (el) m[n.id] = { w: el.offsetWidth, h: el.offsetHeight }
      }
      setSizes((prev) => {
        let changed = false
        for (const k in m) {
          if (!prev[k] || prev[k].w !== m[k].w || prev[k].h !== m[k].h) changed = true
        }
        return changed ? { ...prev, ...m } : prev
      })
    }
    remeasure()
    let cancelled = false
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(() => { if (!cancelled) remeasure() })
    }
    const ro = new ResizeObserver(remeasure)
    for (const n of nodes) {
      const el = nodeRefs.current[n.id]
      if (el) ro.observe(el)
    }
    return () => { cancelled = true; ro.disconnect() }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nodes])

  // Node centers in DESIGN pixels, clamped so each box stays inside the design canvas.
  const centers = useMemo(() => {
    const out = {}
    for (const n of nodes) {
      const sz = sizes[n.id] || { w: n.w || 210, h: 130 }
      const hw = sz.w / 2
      const hh = sz.h / 2
      let cx = (n.x / 100) * DESIGN_W
      let cy = (n.y / 100) * designH
      cx = Math.max(hw + MARGIN, Math.min(DESIGN_W - hw - MARGIN, cx))
      cy = Math.max(hh + MARGIN, Math.min(designH - hh - MARGIN, cy))
      out[n.id] = { cx, cy }
    }
    return out
  }, [nodes, sizes, designH])

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
  const onSelect = useCallback((id) => setSelectedId((prev) => (prev === id ? null : id)), [])
  const selectedNode = selectedId ? nodes.find((n) => n.id === selectedId) : null
  const detail = selectedNode && getDetail ? getDetail(selectedNode) : null

  return (
    <div className="ngraph" aria-label={ariaLabel}>
      <div className="ngraph__fit" ref={fitRef}>
        <div className="ngraph__scaler" style={{ width: DESIGN_W * scale, height: designH * scale }}>
          <div
            className="ngraph__canvas"
            style={{ width: DESIGN_W, height: designH, transform: `scale(${scale})`, transformOrigin: 'top left' }}
          >
            <AnimatedConnections
              centers={centers}
              edges={edges}
              width={DESIGN_W}
              height={designH}
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
        </div>
      </div>
      {detail && <DetailPanel node={selectedNode} detail={detail} onClose={() => setSelectedId(null)} />}
    </div>
  )
}

export default NodeGraph
