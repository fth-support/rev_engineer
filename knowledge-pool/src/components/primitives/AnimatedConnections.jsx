import { useMemo, useRef, useLayoutEffect, useState } from 'react'

// SVG bezier connectors with draw-in animation + a particle that flows along each path.
// Adapted from ag-kit-main/uxui/src/components/Connections.tsx, generalized to
// rectangular nodes positioned by percent and to 2D (horizontal or vertical) routing.

function hashDelay(key) {
  let h = 0
  for (let i = 0; i < key.length; i++) h = ((h << 5) - h + key.charCodeAt(i)) | 0
  return (Math.abs(h) % 3000) / 1000
}

function buildPath(fx, fy, tx, ty) {
  const dx = tx - fx
  const dy = ty - fy
  if (Math.abs(dx) >= Math.abs(dy)) {
    const mx = (fx + tx) / 2
    return `M ${fx} ${fy} C ${mx} ${fy}, ${mx} ${ty}, ${tx} ${ty}`
  }
  const my = (fy + ty) / 2
  return `M ${fx} ${fy} C ${fx} ${my}, ${tx} ${my}, ${tx} ${ty}`
}

function DrawPath({ d, stroke, strokeWidth, delay }) {
  const ref = useRef(null)
  const [length, setLength] = useState(0)
  useLayoutEffect(() => {
    if (ref.current) setLength(ref.current.getTotalLength())
  }, [d])
  return (
    <path
      ref={ref}
      d={d}
      fill="none"
      stroke={stroke}
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeDasharray={length || 1}
      strokeDashoffset={length || 1}
      style={{
        animationName: length ? 'draw-line' : undefined,
        animationDuration: length ? '1.3s' : undefined,
        animationTimingFunction: 'cubic-bezier(0.4,0,0.2,1)',
        animationFillMode: 'forwards',
        animationDelay: `${delay}s`,
      }}
    />
  )
}

function AnimatedConnections({ nodes, edges, width, height, selectedId, relatedIds }) {
  const paths = useMemo(() => {
    if (!width || !height) return []
    const map = new Map(nodes.map((n) => [n.id, n]))
    return edges
      .map((e, i) => {
        const a = map.get(e.from)
        const b = map.get(e.to)
        if (!a || !b) return null
        const fx = (a.x / 100) * width
        const fy = (a.y / 100) * height
        const tx = (b.x / 100) * width
        const ty = (b.y / 100) * height
        const key = `${e.from}-${e.to}-${i}`
        return {
          key, index: i, from: e.from, to: e.to, label: e.label,
          d: buildPath(fx, fy, tx, ty),
          lx: (fx + tx) / 2, ly: (fy + ty) / 2,
          delay: hashDelay(key),
        }
      })
      .filter(Boolean)
  }, [nodes, edges, width, height])

  if (!width || !height) return null

  return (
    <svg className="ngraph__svg" width={width} height={height} viewBox={`0 0 ${width} ${height}`} aria-hidden="true">
      {paths.map((p) => {
        const isRelated = selectedId && relatedIds ? relatedIds.has(p.from) && relatedIds.has(p.to) : false
        const dimmed = selectedId ? !isRelated : false
        return (
          <g key={p.key} opacity={dimmed ? 0.12 : 1} style={{ transition: 'opacity .25s' }}>
            <DrawPath
              d={p.d}
              stroke={isRelated ? 'var(--connection-highlight)' : 'var(--connection-line)'}
              strokeWidth={isRelated ? 2.6 : 1.8}
              delay={0.5 + p.index * 0.06}
            />
            <circle
              r={isRelated ? 3.2 : 2}
              fill={isRelated ? 'var(--particle)' : 'var(--connection-line)'}
              style={{
                offsetPath: `path("${p.d}")`,
                animationName: 'orbit',
                animationDuration: `${isRelated ? 2.2 : 4}s`,
                animationTimingFunction: 'linear',
                animationIterationCount: 'infinite',
                animationDelay: `${-p.delay}s`,
                willChange: 'offset-distance',
              }}
            />
            {p.label && (
              <foreignObject x={p.lx - 80} y={p.ly - 14} width="160" height="28" style={{ overflow: 'visible' }}>
                <div className={`ngraph__edge-label ${isRelated ? 'is-related' : ''}`} style={{ opacity: dimmed ? 0.15 : 1 }}>
                  {p.label}
                </div>
              </foreignObject>
            )}
          </g>
        )
      })}
    </svg>
  )
}

export default AnimatedConnections
