import { useMemo, useRef, useState, useLayoutEffect, useCallback } from 'react'
import DetailPanel from './DetailPanel'

// Interactive node graph rendered as ONE inline <svg viewBox>. The browser scales the
// fixed coordinate space (VBW × VBH) to the container width, preserving aspect ratio —
// so the diagram is always fully visible, never clipped, and fully responsive WITHOUT
// any JS scale math, ResizeObserver, or size measurement (which previously raced with
// the webfont). Node box heights are computed deterministically from content, calibrated
// to the loaded font, so a card never exceeds its box → no overflow, no clipping.

const VBW = 1080
const MARGIN = 14

// Deterministic node box size (in viewBox units), calibrated to the loaded Inter font.
function nodeSize(n) {
  const w = n.w || 210
  const h = n.fields
    ? 44 + n.fields.length * 36          // entity card: header + field rows
    : 62 + (n.items ? n.items.length : 0) * 30 // graph card: header + bullet items
  return { w, h }
}

function buildPath(fx, fy, tx, ty) {
  const dx = tx - fx, dy = ty - fy
  if (Math.abs(dx) >= Math.abs(dy)) {
    const mx = (fx + tx) / 2
    return `M ${fx} ${fy} C ${mx} ${fy}, ${mx} ${ty}, ${tx} ${ty}`
  }
  const my = (fy + ty) / 2
  return `M ${fx} ${fy} C ${fx} ${my}, ${tx} ${my}, ${tx} ${ty}`
}
function hashDelay(k) { let h = 0; for (let i = 0; i < k.length; i++) h = ((h << 5) - h + k.charCodeAt(i)) | 0; return (Math.abs(h) % 3000) / 1000 }

function DrawPath({ d, stroke, width, delay }) {
  const ref = useRef(null)
  const [len, setLen] = useState(0)
  useLayoutEffect(() => { if (ref.current) setLen(ref.current.getTotalLength()) }, [d])
  return (
    <path ref={ref} d={d} fill="none" stroke={stroke} strokeWidth={width} strokeLinecap="round"
      strokeDasharray={len || 1} strokeDashoffset={len || 1}
      style={{
        animationName: len ? 'draw-line' : undefined,
        animationDuration: len ? '1.3s' : undefined,
        animationTimingFunction: 'cubic-bezier(0.4,0,0.2,1)',
        animationFillMode: 'forwards',
        animationDelay: `${delay}s`,
      }} />
  )
}

function NodeGraph({ nodes, edges, aspect = 1.6, renderNode, getDetail, ariaLabel = 'Diagram' }) {
  const VBH = Math.round(VBW / aspect)
  const [selectedId, setSelectedId] = useState(null)

  // Node centers + boxes in viewBox units, clamped so every box stays inside the canvas.
  const centers = useMemo(() => {
    const out = {}
    for (const n of nodes) {
      const { w, h } = nodeSize(n)
      const hw = w / 2, hh = h / 2
      let cx = (n.x / 100) * VBW
      let cy = (n.y / 100) * VBH
      cx = Math.max(hw + MARGIN, Math.min(VBW - hw - MARGIN, cx))
      cy = Math.max(hh + MARGIN, Math.min(VBH - hh - MARGIN, cy))
      out[n.id] = { cx, cy, w, h }
    }
    return out
  }, [nodes, VBH])

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
  const onSelect = useCallback((id) => setSelectedId((p) => (p === id ? null : id)), [])
  const selectedNode = selectedId ? nodes.find((n) => n.id === selectedId) : null
  const detail = selectedNode && getDetail ? getDetail(selectedNode) : null

  return (
    <div className="ngraph" aria-label={ariaLabel}>
      <svg className="ngraph__svg" viewBox={`0 0 ${VBW} ${VBH}`} preserveAspectRatio="xMidYMid meet" role="img">
        {/* connectors */}
        {edges.map((e, i) => {
          const a = centers[e.from]
          const b = centers[e.to]
          if (!a || !b) return null
          const d = buildPath(a.cx, a.cy, b.cx, b.cy)
          const key = `${e.from}-${e.to}-${i}`
          const isRel = selectedId && relatedIds ? relatedIds.has(e.from) && relatedIds.has(e.to) : false
          const dim = selectedId ? !isRel : false
          return (
            <g key={key} opacity={dim ? 0.12 : 1} style={{ transition: 'opacity .25s' }}>
              <DrawPath d={d} stroke={isRel ? 'var(--connection-highlight)' : 'var(--connection-line)'} width={isRel ? 2.6 : 1.8} delay={0.5 + i * 0.06} />
              <circle r={isRel ? 3.4 : 2.2} fill={isRel ? 'var(--particle)' : 'var(--connection-line)'}
                style={{ offsetPath: `path("${d}")`, animationName: 'orbit', animationDuration: `${isRel ? 2.2 : 4}s`, animationTimingFunction: 'linear', animationIterationCount: 'infinite', animationDelay: `${-hashDelay(key)}s` }} />
              {e.label && (
                <foreignObject x={(a.cx + b.cx) / 2 - 85} y={(a.cy + b.cy) / 2 - 15} width="170" height="30" style={{ overflow: 'visible' }}>
                  <div xmlns="http://www.w3.org/1999/xhtml" className={`ngraph__edge-label ${isRel ? 'is-related' : ''}`} style={{ opacity: dim ? 0.15 : 1 }}>{e.label}</div>
                </foreignObject>
              )}
            </g>
          )
        })}
        {/* nodes */}
        {nodes.map((n) => {
          const c = centers[n.id]
          const isSelected = n.id === selectedId
          const isDimmed = selectedId ? !relatedIds.has(n.id) : false
          return (
            <foreignObject key={n.id} x={c.cx - c.w / 2} y={c.cy - c.h / 2} width={c.w} height={c.h + 6}
              style={{ overflow: 'visible', opacity: isDimmed ? 0.32 : 1, transition: 'opacity .25s' }}>
              <div xmlns="http://www.w3.org/1999/xhtml" style={{ width: '100%' }}>
                {renderNode({ node: n, isSelected, isDimmed, onSelect })}
              </div>
            </foreignObject>
          )
        })}
      </svg>
      {detail && <DetailPanel node={selectedNode} detail={detail} onClose={() => setSelectedId(null)} />}
    </div>
  )
}

export default NodeGraph
