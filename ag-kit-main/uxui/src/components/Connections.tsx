import { memo, useMemo, useRef, useLayoutEffect, useState } from "react";
import type { PackageNode, Dependency } from "../data";
import { getNodeRadius } from "../nodeLayout";

interface ConnectionsProps {
  nodes: PackageNode[];
  connections: Dependency[];
  selectedId: string | null;
  highlightedIds: Set<string> | null;
  visiblePkgIds: Set<string> | null;
  containerWidth: number;
  containerHeight: number;
}

function hashDelay(key: string): number {
  let h = 0;
  for (let i = 0; i < key.length; i++) {
    h = ((h << 5) - h + key.charCodeAt(i)) | 0;
  }
  return (Math.abs(h) % 3000) / 1000;
}

export const Connections = memo(function Connections({
  nodes,
  connections,
  selectedId,
  highlightedIds,
  visiblePkgIds,
  containerWidth,
  containerHeight,
}: ConnectionsProps) {
  const paths = useMemo(() => {
    if (containerWidth === 0 || containerHeight === 0) return [];
    const nodeMap = new Map(nodes.map((n) => [n.id, n]));
    return connections.map((conn, i) => {
      const fromNode = nodeMap.get(conn.from);
      const toNode = nodeMap.get(conn.to);
      if (!fromNode || !toNode) return null;

      const fromCx = (fromNode.x / 100) * containerWidth;
      const fromCy = (fromNode.y / 100) * containerHeight;
      const toCx = (toNode.x / 100) * containerWidth;
      const toCy = (toNode.y / 100) * containerHeight;

      const fromR = getNodeRadius(fromNode.id);
      const toR = getNodeRadius(toNode.id);

      const dy = toCy - fromCy;

      let startY: number, endY: number;
      if (Math.abs(dy) < fromR + toR) {
        startY = fromCy;
        endY = toCy;
      } else if (dy > 0) {
        startY = fromCy + fromR;
        endY = toCy - toR;
      } else {
        startY = fromCy - fromR;
        endY = toCy + toR;
      }

      const midY = (startY + endY) / 2;
      const key = `${conn.from}-${conn.to}`;

      return {
        key,
        index: i,
        fromId: conn.from,
        toId: conn.to,
        d: `M ${fromCx} ${startY} C ${fromCx} ${midY}, ${toCx} ${midY}, ${toCx} ${endY}`,
        delay: hashDelay(key),
      };
    });
  }, [nodes, connections, containerWidth, containerHeight]);

  if (containerWidth === 0 || containerHeight === 0) return null;

  return (
    <svg
      className="!absolute inset-0 pointer-events-none"
      aria-hidden="true"
      width={containerWidth}
      height={containerHeight}
      viewBox={`0 0 ${containerWidth} ${containerHeight}`}
      style={{ zIndex: 1 }}
    >
      {paths.map((p) => {
        if (!p) return null;

        const isRelated =
          selectedId && highlightedIds
            ? highlightedIds.has(p.fromId) && highlightedIds.has(p.toId)
            : false;
        const isFiltered = visiblePkgIds
          ? !visiblePkgIds.has(p.fromId) || !visiblePkgIds.has(p.toId)
          : false;
        const isDimmed = (selectedId ? !isRelated : false) || isFiltered;

        return (
          <g key={p.key} opacity={isDimmed ? 0.08 : 1}>
            <DrawPath
              d={p.d}
              stroke={
                isRelated
                  ? "var(--connection-highlight)"
                  : "var(--connection-line)"
              }
              strokeWidth={isRelated ? 2.5 : 2}
              delay={0.7 + p.index * 0.03}
            />

            <circle
              r={isRelated ? 3 : 1.5}
              fill={isRelated ? "var(--accent)" : "var(--particle-dim)"}
              className="animate-[orbit_4s_linear_infinite]"
              style={{
                offsetPath: `path("${p.d}")`,
                animationDelay: `${-p.delay}s`,
                animationDuration: isRelated ? "2s" : "4s",
                willChange: "offset-distance",
              }}
            />
          </g>
        );
      })}
    </svg>
  );
});

/** Individual path that measures itself and animates stroke-dashoffset */
function DrawPath({
  d,
  stroke,
  strokeWidth,
  delay,
}: {
  d: string;
  stroke: string;
  strokeWidth: number;
  delay: number;
}) {
  const ref = useRef<SVGPathElement>(null);
  const [length, setLength] = useState(0);

  useLayoutEffect(() => {
    if (ref.current) {
      setLength(ref.current.getTotalLength());
    }
  }, [d]);

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
        animation: length
          ? `draw-line 1.4s cubic-bezier(0.4, 0, 0.2, 1) forwards`
          : undefined,
        animationDelay: `${delay}s`,
      }}
    />
  );
}
