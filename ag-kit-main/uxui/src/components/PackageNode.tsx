import { useState, memo, useRef, useEffect, useCallback } from "react";
import { createPortal } from "react-dom";
import type { PackageNode as PackageNodeData } from "../data";
import { CATEGORY_META } from "../data";
import { getNodeRadius, getNodeDelay } from "../nodeLayout";
import { Icon } from "./Icons";

const VULN_BADGE: Record<string, { label: string }> = {
  critical: { label: "CRITICAL" },
  high: { label: "HIGH" },
  moderate: { label: "MODERATE" },
};

function TooltipPortal({
  pkg,
  anchorRect,
}: {
  pkg: PackageNodeData;
  anchorRect: DOMRect;
}) {
  const flip = anchorRect.bottom > window.innerHeight - 280;
  const catMeta = CATEGORY_META[pkg.category];

  return createPortal(
    <div
      className="node-tooltip-enter"
      style={{
        position: "fixed",
        top: flip ? anchorRect.top - 12 : anchorRect.bottom + 12,
        left: anchorRect.left + anchorRect.width / 2,
        transform: flip ? "translate(-50%, -100%)" : "translateX(-50%)",
        zIndex: 9999,
        pointerEvents: "none",
      }}
    >
      <div className="rounded-2xl px-5 py-4 min-w-[220px] max-w-[280px] relative overflow-hidden bg-[var(--tooltip-bg)] border border-[var(--tooltip-border)] shadow-[var(--tooltip-shadow)]">
        <div className="relative z-10">
          <div className="flex items-center gap-2 mb-1.5">
            <div
              className="w-5 h-5 rounded-md flex items-center justify-center shrink-0"
              style={{
                background: `var(--cat-${pkg.category})`,
                color: "#fff",
              }}
            >
              <Icon name={catMeta.icon} size={11} />
            </div>
            <span className="text-[14px] font-bold tracking-tight text-[var(--text-primary)] truncate">
              {pkg.name}
            </span>
          </div>

          <p className="text-[12px] leading-relaxed text-[var(--text-secondary)] mb-3">
            {pkg.description}
          </p>

          <div className="h-px w-full mb-2.5 bg-[var(--tooltip-divider)]" />

          <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-[12px]">
            <span className="text-[var(--text-muted)]">Version</span>
            <span className="text-[var(--text-primary)] font-mono font-medium">
              {pkg.version}
            </span>
            <span className="text-[var(--text-muted)]">License</span>
            <span className="text-[var(--text-primary)] font-medium">
              {pkg.license}
            </span>
            {pkg.size && (
              <>
                <span className="text-[var(--text-muted)]">Size</span>
                <span className="text-[var(--text-primary)] font-medium">
                  {pkg.size}
                </span>
              </>
            )}
            <span className="text-[var(--text-muted)]">Category</span>
            <span
              className="font-medium"
              style={{ color: `var(--cat-${pkg.category})` }}
            >
              {catMeta.label}
            </span>
          </div>

          {pkg.vulnerability !== "none" && (
            <div
              className="mt-2.5 flex items-center gap-1.5 text-[12px] font-bold"
              style={{ color: `var(--vuln-${pkg.vulnerability}-text)` }}
            >
              <Icon
                name="alert-triangle"
                size={12}
                style={{ color: `var(--vuln-${pkg.vulnerability}-text)` }}
              />
              {pkg.vulnCount} {VULN_BADGE[pkg.vulnerability]?.label}
            </div>
          )}

          {pkg.outdated && (
            <div className="mt-1.5 flex items-center gap-1.5 text-[12px] font-semibold text-[var(--outdated-text)]">
              <Icon name="clock" size={12} />
              Outdated — update available
            </div>
          )}
        </div>
      </div>
    </div>,
    document.body
  );
}

export const PackageNodeComponent = memo(
  function PackageNodeComponent({
    pkg,
    selectedId,
    highlightedIds,
    visiblePkgIds,
    isDark,
    onSelect,
  }: {
    pkg: PackageNodeData;
    selectedId: string | null;
    highlightedIds: Set<string> | null;
    visiblePkgIds: Set<string> | null;
    isDark: boolean;
    onSelect: (id: string) => void;
  }) {
    const isSelected = selectedId === pkg.id;
    const isHighlighted = highlightedIds ? highlightedIds.has(pkg.id) : false;
    const isDimmed =
      (selectedId !== null && !isHighlighted) ||
      (visiblePkgIds !== null && !visiblePkgIds.has(pkg.id));

    const [showTooltip, setShowTooltip] = useState(false);
    const ref = useRef<HTMLButtonElement>(null);
    const [rect, setRect] = useState<DOMRect | null>(null);

    const hasVuln = pkg.vulnerability !== "none";
    const overrideColor = isDark ? pkg.colorDark : pkg.colorLight;
    const catColor = overrideColor ?? `var(--cat-${pkg.category})`;
    const r = getNodeRadius(pkg.id);
    const size = r * 2;

    const updateRect = useCallback(() => {
      if (ref.current) setRect(ref.current.getBoundingClientRect());
    }, []);

    const show = useCallback(() => {
      setShowTooltip(true);
    }, []);
    const hide = useCallback(() => {
      setShowTooltip(false);
    }, []);

    useEffect(() => {
      if (!showTooltip) return;
      updateRect();
      window.addEventListener("scroll", updateRect, true);
      window.addEventListener("resize", updateRect);
      return () => {
        window.removeEventListener("scroll", updateRect, true);
        window.removeEventListener("resize", updateRect);
      };
    }, [showTooltip, updateRect]);

    let borderColor = "var(--node-default-border)";
    let bg = "var(--node-default-bg)";
    let shadow = "var(--node-default-shadow)";

    if (isSelected) {
      borderColor = "var(--node-selected-border)";
      bg = "var(--node-selected-bg)";
      shadow = "var(--node-selected-glow)";
    } else if (hasVuln) {
      borderColor = `var(--vuln-${pkg.vulnerability}-border)`;
      bg = `var(--vuln-${pkg.vulnerability}-bg)`;
    }

    if (isHighlighted && !isSelected) {
      shadow = "var(--node-hover-glow)";
    }

    return (
      <div
        className="absolute flex flex-col items-center node-enter"
        style={{
          left: `${pkg.x}%`,
          top: `${pkg.y}%`,
          width: size,
          marginLeft: -r,
          marginTop: -r,
          zIndex: showTooltip || isSelected ? 30 : 5,
          opacity: isDimmed ? 0.2 : 1,
          transition: "opacity 0.15s ease",
          animationDelay: `${getNodeDelay(pkg.id)}ms`,
        }}
      >
        {hasVuln && !isDimmed && (
          <div
            className="absolute rounded-full pointer-events-none animate-[pulse-ring_3s_ease-in-out_infinite]"
            style={{
              width: size + 12,
              height: size + 12,
              top: -6,
              left: -6,
              border: `2px solid var(--vuln-${pkg.vulnerability}-border)`,
            }}
          />
        )}

        <button
          ref={ref}
          type="button"
          className="node-btn relative flex items-center justify-center rounded-full border-[1.5px] select-none cursor-pointer"
          style={{
            width: size,
            height: size,
            borderColor,
            background: bg,
            boxShadow: shadow,
          }}
          onClick={() => onSelect(pkg.id)}
          onMouseEnter={show}
          onMouseLeave={hide}
          onFocus={show}
          onBlur={hide}
          aria-label={`${pkg.name} v${pkg.version}`}
        >
          <div
            className="absolute inset-0 rounded-full pointer-events-none"
            style={{
              background: `conic-gradient(from 0deg, ${catColor} 0deg, ${catColor} 90deg, transparent 90deg)`,
              opacity: 0.25,
              maskImage:
                "radial-gradient(circle, transparent 60%, black 62%, black 100%)",
              WebkitMaskImage:
                "radial-gradient(circle, transparent 60%, black 62%, black 100%)",
            }}
          />
          <Icon
            name={CATEGORY_META[pkg.category].icon}
            size={r * 0.6}
            style={{ color: catColor }}
          />
        </button>

        <div className="flex flex-col items-center mt-1 pointer-events-none">
          <span className="text-[10px] font-bold text-[var(--text-primary)] leading-tight text-center max-w-[90px] truncate">
            {pkg.name}
          </span>
          <div className="flex items-center gap-0.5 mt-0.5">
            <span className="text-[9px] font-mono text-[var(--text-muted)]">
              {pkg.version}
            </span>
            {hasVuln && (
              <span
                className="w-[5px] h-[5px] rounded-full animate-pulse"
                style={{ background: `var(--vuln-${pkg.vulnerability}-text)` }}
              />
            )}
            {pkg.outdated && (
              <span className="text-[8px] font-bold px-0.5 rounded bg-[var(--outdated-bg)] text-[var(--outdated-text)]">
                OLD
              </span>
            )}
          </div>
        </div>

        {showTooltip && rect && <TooltipPortal pkg={pkg} anchorRect={rect} />}
      </div>
    );
  },
  (prev, next) => {
    if (
      prev.pkg !== next.pkg ||
      prev.onSelect !== next.onSelect ||
      prev.isDark !== next.isDark
    )
      return false;

    const prevSelected = prev.selectedId === prev.pkg.id;
    const nextSelected = next.selectedId === next.pkg.id;
    if (prevSelected !== nextSelected) return false;

    const prevHighlighted = prev.highlightedIds
      ? prev.highlightedIds.has(prev.pkg.id)
      : false;
    const nextHighlighted = next.highlightedIds
      ? next.highlightedIds.has(next.pkg.id)
      : false;
    if (prevHighlighted !== nextHighlighted) return false;

    const prevDimmed =
      (prev.selectedId !== null && !prevHighlighted) ||
      (prev.visiblePkgIds !== null && !prev.visiblePkgIds.has(prev.pkg.id));
    const nextDimmed =
      (next.selectedId !== null && !nextHighlighted) ||
      (next.visiblePkgIds !== null && !next.visiblePkgIds.has(next.pkg.id));
    if (prevDimmed !== nextDimmed) return false;

    return true;
  }
);
