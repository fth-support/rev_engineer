import { Icon } from "./Icons";

export function Header({
  themeName,
  setThemeName,
  totalPackages,
  vulnCount,
  depCount,
}: {
  themeName: string;
  setThemeName: (name: string) => void;
  totalPackages: number;
  vulnCount: number;
  depCount: number;
}) {
  const isDark = themeName === "dark";

  return (
    <div className="flex items-center justify-between mb-4">
      <div className="flex items-center gap-3">
        <div className="w-8 h-8 rounded-lg bg-[var(--accent)] flex items-center justify-center">
          <Icon name="package" size={16} style={{ color: "var(--page-bg)" }} />
        </div>
        <div>
          <h1 className="text-[15px] font-bold tracking-tight m-0 leading-tight text-[var(--text-primary)]">
            npm dependencies
          </h1>
          <div className="flex items-center gap-2 mt-0.5">
            <span className="text-[11px] text-[var(--text-muted)]">
              {totalPackages} packages
            </span>
            <span className="text-[11px] text-[var(--text-muted)]">·</span>
            <span className="text-[11px] text-[var(--text-muted)]">
              {depCount} deps
            </span>
            {vulnCount > 0 && (
              <>
                <span className="text-[11px] text-[var(--text-muted)]">·</span>
                <span className="flex items-center gap-0.5 text-[11px] text-[var(--vuln-high-text)] font-semibold">
                  <Icon name="alert-triangle" size={10} />
                  {vulnCount}
                </span>
              </>
            )}
          </div>
        </div>
      </div>

      <button
        type="button"
        role="switch"
        aria-checked={isDark}
        aria-label="Toggle dark mode"
        className="relative w-[42px] h-[22px] rounded-full border border-[var(--border-subtle)] cursor-pointer shrink-0 transition-colors"
        style={{
          background: isDark ? "rgba(125,211,252,0.12)" : "rgba(0,0,0,0.06)",
        }}
        onClick={() => setThemeName(isDark ? "light" : "dark")}
      >
        <div
          className="absolute top-[2px] w-[16px] h-[16px] rounded-full flex items-center justify-center shadow-sm transition-all duration-200"
          style={{
            left: isDark ? 22 : 3,
            background: isDark ? "var(--accent)" : "var(--text-secondary)",
          }}
        >
          {isDark ? (
            <svg
              width="9"
              height="9"
              viewBox="0 0 24 24"
              fill="none"
              stroke="#fff"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
            </svg>
          ) : (
            <svg
              width="9"
              height="9"
              viewBox="0 0 24 24"
              fill="none"
              stroke="#fff"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <circle cx="12" cy="12" r="5" />
              <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
            </svg>
          )}
        </div>
      </button>
    </div>
  );
}
