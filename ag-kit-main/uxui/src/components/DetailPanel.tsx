import type { PackageNode } from "../data";
import { CATEGORY_META, PACKAGES } from "../data";
import { Icon } from "./Icons";

const VULN_LABELS: Record<string, string> = {
  critical: "CRITICAL",
  high: "HIGH",
  moderate: "MODERATE",
};

export function DetailPanel({
  pkg,
  onClose,
}: {
  pkg: PackageNode;
  onClose: () => void;
}) {
  const dependents = PACKAGES.filter((p) => p.dependencies.includes(pkg.id));

  return (
    <div
      className="panel-enter border-t border-[var(--border-subtle)] bg-[var(--tooltip-bg)] px-4 py-3"
      role="region"
      aria-label={`Details for ${pkg.name}`}
    >
      {/* Top row: name + close */}
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2.5">
          <div
            className="w-7 h-7 rounded-full flex items-center justify-center shrink-0"
            style={{
              background: `var(--cat-${pkg.category})`,
              color: "var(--page-bg)",
            }}
          >
            <Icon name={CATEGORY_META[pkg.category].icon} size={14} />
          </div>
          <div>
            <span className="text-[13px] font-bold text-[var(--text-primary)]">
              {pkg.name}
            </span>
            <span className="text-[11px] font-mono ml-2 px-1.5 py-0.5 rounded bg-[var(--badge-bg)] text-[var(--badge-text)]">
              v{pkg.version}
            </span>
          </div>
        </div>
        <button
          type="button"
          onClick={onClose}
          className="text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors cursor-pointer p-1"
          aria-label="Close"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.5"
            strokeLinecap="round"
          >
            <path d="M18 6L6 18M6 6l12 12" />
          </svg>
        </button>
      </div>

      {/* Stats row */}
      <div className="flex items-center gap-3 flex-wrap text-[11px]">
        <span
          className="text-[var(--text-muted)]"
          style={{ color: `var(--cat-${pkg.category})` }}
        >
          {CATEGORY_META[pkg.category].label}
        </span>

        <span className="w-px h-3 bg-[var(--border-subtle)]" />

        <span className="text-[var(--text-secondary)]">{pkg.license}</span>

        {pkg.size && (
          <>
            <span className="w-px h-3 bg-[var(--border-subtle)]" />
            <span className="text-[var(--text-secondary)]">{pkg.size}</span>
          </>
        )}

        <span className="w-px h-3 bg-[var(--border-subtle)]" />

        {pkg.vulnerability !== "none" ? (
          <span
            className="font-bold flex items-center gap-1"
            style={{ color: `var(--vuln-${pkg.vulnerability}-text)` }}
          >
            <Icon
              name="alert-triangle"
              size={12}
              style={{ color: `var(--vuln-${pkg.vulnerability}-text)` }}
            />
            {pkg.vulnCount} {VULN_LABELS[pkg.vulnerability]}
          </span>
        ) : (
          <span className="font-semibold text-[var(--vuln-none-text)] flex items-center gap-1">
            <Icon name="check-circle" size={12} />
            Clean
          </span>
        )}

        {pkg.outdated && (
          <span className="font-bold text-[var(--outdated-text)] flex items-center gap-1">
            <Icon name="clock" size={12} />
            Outdated
          </span>
        )}
      </div>

      {/* Deps row */}
      {(pkg.dependencies.length > 0 || dependents.length > 0) && (
        <div className="flex items-center gap-3 flex-wrap text-[11px] mt-2 pt-2 border-t border-[var(--border-subtle)]">
          {pkg.dependencies.length > 0 && (
            <span className="text-[var(--text-muted)]">
              <span className="text-[var(--text-secondary)] font-semibold">
                Deps:
              </span>{" "}
              {pkg.dependencies
                .map((id) => PACKAGES.find((p) => p.id === id)?.name ?? id)
                .join(", ")}
            </span>
          )}

          {dependents.length > 0 && (
            <span className="text-[var(--text-muted)]">
              <span className="text-[var(--text-secondary)] font-semibold">
                Used by:
              </span>{" "}
              {dependents.map((d) => d.name).join(", ")}
            </span>
          )}
        </div>
      )}
    </div>
  );
}
