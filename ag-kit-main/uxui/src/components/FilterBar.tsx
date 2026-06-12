import type { PackageCategory } from "../data";
import { CATEGORY_META } from "../data";
import { Icon } from "./Icons";

const ALL_CATEGORIES = Object.keys(CATEGORY_META) as PackageCategory[];

export function FilterBar({
  activeFilter,
  onFilter,
  showVulnOnly,
  onToggleVuln,
}: {
  activeFilter: PackageCategory | null;
  onFilter: (cat: PackageCategory | null) => void;
  showVulnOnly: boolean;
  onToggleVuln: () => void;
}) {
  return (
    <div
      className="flex flex-wrap items-center gap-1 mb-3"
      role="toolbar"
      aria-label="Filter packages"
    >
      <button
        type="button"
        aria-pressed={activeFilter === null && !showVulnOnly}
        className={`text-[11px] font-semibold px-3 py-1 rounded-full border cursor-pointer transition-colors ${
          activeFilter === null && !showVulnOnly
            ? "bg-[var(--accent)] text-[var(--page-bg)] border-transparent"
            : "bg-transparent text-[var(--text-muted)] border-[var(--border-subtle)] hover:text-[var(--text-secondary)]"
        }`}
        onClick={() => onFilter(null)}
      >
        All
      </button>

      {ALL_CATEGORIES.map((cat) => {
        const meta = CATEGORY_META[cat];
        const active = activeFilter === cat;
        return (
          <button
            key={cat}
            type="button"
            aria-pressed={active}
            className={`flex items-center gap-1 text-[11px] font-semibold px-3 py-1 rounded-full border cursor-pointer transition-colors ${
              active
                ? "border-transparent"
                : "bg-transparent text-[var(--text-muted)] border-[var(--border-subtle)] hover:text-[var(--text-secondary)]"
            }`}
            style={
              active
                ? { background: `var(--cat-${cat})`, color: "var(--page-bg)" }
                : undefined
            }
            onClick={() => onFilter(active ? null : cat)}
          >
            <Icon name={meta.icon} size={11} />
            {meta.label}
          </button>
        );
      })}

      <button
        type="button"
        aria-pressed={showVulnOnly}
        className={`flex items-center gap-1 text-[11px] font-semibold px-3 py-1 rounded-full border cursor-pointer transition-colors ${
          showVulnOnly
            ? "bg-[var(--vuln-high-bg)] text-[var(--vuln-high-text)] border-[var(--vuln-high-border)]"
            : "bg-transparent text-[var(--text-muted)] border-[var(--border-subtle)] hover:text-[var(--vuln-high-text)]"
        }`}
        onClick={onToggleVuln}
      >
        <Icon name="alert-triangle" size={11} />
        Issues
      </button>
    </div>
  );
}
