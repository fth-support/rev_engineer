import { useState, useCallback, useMemo, useRef, useLayoutEffect } from "react";
import { themes, type Theme } from "./theme";
import { PACKAGES, DEPENDENCIES, type PackageCategory } from "./data";
import { Connections } from "./components/Connections";
import { Header } from "./components/Header";
import { FilterBar } from "./components/FilterBar";
import { PackageNodeComponent } from "./components/PackageNode";
import { DetailPanel } from "./components/DetailPanel";

function themeToVars(t: Theme): React.CSSProperties {
  return Object.fromEntries(
    Object.entries(t).map(([k, v]) => [
      "--" + k.replace(/[A-Z]/g, (c) => "-" + c.toLowerCase()),
      v,
    ])
  ) as React.CSSProperties;
}

/** Precomputed: for each node, which nodes are its direct neighbors */
const RELATED_MAP: Record<string, Set<string>> = {};
for (const pkg of PACKAGES) {
  const s = new Set<string>([pkg.id]);
  for (const dep of DEPENDENCIES) {
    if (dep.from === pkg.id) s.add(dep.to);
    if (dep.to === pkg.id) s.add(dep.from);
  }
  RELATED_MAP[pkg.id] = s;
}

const VULN_COUNT = PACKAGES.filter((p) => p.vulnerability !== "none").reduce(
  (sum, p) => sum + (p.vulnCount ?? 0),
  0
);

function App() {
  const [themeName, setThemeName] = useState("dark");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [categoryFilter, setCategoryFilter] = useState<PackageCategory | null>(
    null
  );
  const [showVulnOnly, setShowVulnOnly] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  const [containerSize, setContainerSize] = useState({ width: 0, height: 0 });

  useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    // Get initial size synchronously
    const rect = el.getBoundingClientRect();
    setContainerSize({ width: rect.width, height: rect.height });
    // Then observe for changes
    const ro = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect;
      setContainerSize({ width, height });
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const t = themes[themeName];
  const themeVars = useMemo(() => themeToVars(t), [t]);

  useLayoutEffect(() => {
    const root = document.documentElement;
    for (const [key, value] of Object.entries(themeVars)) {
      root.style.setProperty(key, value as string);
    }
  }, [themeVars]);

  const highlightedIds = selectedId ? RELATED_MAP[selectedId] : null;

  const visiblePkgIds = useMemo(() => {
    if (!categoryFilter && !showVulnOnly) return null;
    let filtered = PACKAGES;
    if (categoryFilter)
      filtered = filtered.filter((p) => p.category === categoryFilter);
    if (showVulnOnly)
      filtered = filtered.filter(
        (p) => p.vulnerability !== "none" || p.outdated
      );
    return new Set(filtered.map((p) => p.id));
  }, [categoryFilter, showVulnOnly]);

  const handleSelect = useCallback((id: string) => {
    setSelectedId((prev) => (prev === id ? null : id));
  }, []);

  const selectedPkg = selectedId
    ? PACKAGES.find((p) => p.id === selectedId) ?? null
    : null;

  return (
    <main
      className="flex items-center justify-center min-h-screen p-3 sm:p-5 bg-[var(--page-bg)]"
      style={themeVars}
    >
      <div className="theme-transition w-full max-w-[520px] mx-auto">
        <div className="card-enter rounded-xl p-3 sm:p-4 relative bg-[var(--card-bg)] shadow-[var(--card-shadow)] border border-[var(--border-subtle)]">
          <Header
            themeName={themeName}
            setThemeName={setThemeName}
            totalPackages={PACKAGES.length}
            vulnCount={VULN_COUNT}
            depCount={DEPENDENCIES.length}
          />

          <FilterBar
            activeFilter={categoryFilter}
            onFilter={(cat) => {
              setCategoryFilter(cat);
              setShowVulnOnly(false);
            }}
            showVulnOnly={showVulnOnly}
            onToggleVuln={() => {
              setShowVulnOnly((v) => !v);
              setCategoryFilter(null);
            }}
          />

          <div
            className="dot-grid rounded-lg border border-[var(--border-subtle)] bg-[var(--canvas-bg)] relative overflow-hidden"
            aria-label="Dependency graph"
          >
            <div
              ref={containerRef}
              className="relative w-full"
              style={{ aspectRatio: "9 / 14" }}
            >
              <Connections
                nodes={PACKAGES}
                connections={DEPENDENCIES}
                selectedId={selectedId}
                highlightedIds={highlightedIds}
                visiblePkgIds={visiblePkgIds}
                containerWidth={containerSize.width}
                containerHeight={containerSize.height}
              />
              {PACKAGES.map((pkg) => (
                <PackageNodeComponent
                  key={pkg.id}
                  pkg={pkg}
                  selectedId={selectedId}
                  highlightedIds={highlightedIds}
                  visiblePkgIds={visiblePkgIds}
                  isDark={themeName === "dark"}
                  onSelect={handleSelect}
                />
              ))}
            </div>

            {selectedPkg && (
              <div className="absolute bottom-0 left-0 right-0 z-40">
                <DetailPanel
                  pkg={selectedPkg}
                  onClose={() => setSelectedId(null)}
                />
              </div>
            )}
          </div>

          <div className="flex items-center justify-between mt-3 pt-2.5 border-t border-[var(--border-subtle)]">
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-mono font-medium px-1.5 py-0.5 rounded bg-[var(--badge-bg)] text-[var(--badge-text)]">
                package.json
              </span>
              <span className="text-[10px] text-[var(--text-muted)]">
                {PACKAGES.length} packages · {DEPENDENCIES.length} dependencies
              </span>
            </div>
            {VULN_COUNT > 0 && (
              <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-[var(--vuln-high-bg)] text-[var(--vuln-high-text)] border border-[var(--vuln-high-border)]">
                npm audit
              </span>
            )}
          </div>
        </div>
      </div>
    </main>
  );
}

export default App;
