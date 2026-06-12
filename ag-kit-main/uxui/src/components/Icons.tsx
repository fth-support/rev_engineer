import type { ReactNode, CSSProperties } from "react";

const PATHS: Record<string, ReactNode> = {
  atom: (
    <g transform="scale(1.15) translate(-1.57 -1.57)">
      <ellipse cx="12" cy="12" rx="10" ry="3.5" strokeWidth="1" />
      <ellipse
        cx="12"
        cy="12"
        rx="10"
        ry="3.5"
        strokeWidth="1"
        transform="rotate(60 12 12)"
      />
      <ellipse
        cx="12"
        cy="12"
        rx="10"
        ry="3.5"
        strokeWidth="1"
        transform="rotate(120 12 12)"
      />
      <circle cx="12" cy="12" r="2" fill="currentColor" stroke="none" />
    </g>
  ),
  triangle: <path d="M12 3L22 20H2z" />,
  brush: (
    <>
      <path d="M12 3c4.97 0 9 2.69 9 6 0 2.21-1.79 4-4 4h-1.5c-1.1 0-2 .9-2 2 0 .53.21 1.01.55 1.36.23.24.45.58.45 1.14 0 1.38-1.12 2.5-2.5 2.5C7.03 20 3 16.42 3 12S7.03 4 12 3z" />
      <circle cx="7.5" cy="11" r="1.5" fill="currentColor" />
      <circle cx="12" cy="7.5" r="1.5" fill="currentColor" />
      <circle cx="16.5" cy="11" r="1.5" fill="currentColor" />
    </>
  ),
  container: (
    <>
      <path d="M22 8.5l-10-5-10 5" />
      <path d="M2 8.5l10 5 10-5" />
      <path d="M12 13.5v8" />
      <path d="M2 8.5v8l10 5 10-5v-8" />
    </>
  ),
  flask: (
    <>
      <path d="M9 2h6" />
      <path d="M10 2v7.53L4.89 18.12A2 2 0 0 0 6.62 21h10.76a2 2 0 0 0 1.73-2.88L14 9.53V2" />
      <path d="M6.5 17h11" />
    </>
  ),
  plug: (
    <>
      <path d="M7 2v6" />
      <path d="M17 2v6" />
      <path d="M4 8h16v4a8 8 0 0 1-8 8 8 8 0 0 1-8-8V8z" />
      <path d="M12 20v2" />
    </>
  ),
  package: (
    <>
      <path d="M16.5 9.4l-9-5.19" />
      <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
      <path d="M3.27 6.96L12 12.01l8.73-5.05" />
      <path d="M12 22.08V12" />
    </>
  ),
  shield: (
    <>
      <path d="M12 2l8 4v5c0 5.25-3.5 9.74-8 11-4.5-1.26-8-5.75-8-11V6l8-4z" />
      <path d="M9 12l2 2 4-4" />
    </>
  ),
  "alert-triangle": (
    <>
      <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
      <line x1="12" y1="9" x2="12" y2="13" />
      <line x1="12" y1="17" x2="12.01" y2="17" />
    </>
  ),
  "check-circle": (
    <>
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
      <path d="M22 4L12 14.01l-3-3" />
    </>
  ),
  git: (
    <>
      <circle cx="12" cy="18" r="3" />
      <circle cx="6" cy="6" r="3" />
      <circle cx="18" cy="6" r="3" />
      <path d="M18 9v2c0 .6-.4 1-1 1H7c-.6 0-1 .4-1 1v2" />
      <path d="M12 12v3" />
    </>
  ),
  clock: (
    <>
      <circle cx="12" cy="12" r="10" />
      <path d="M12 6v6l4 2" />
    </>
  ),
  download: (
    <>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </>
  ),
  filter: <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />,
};

export function Icon({
  name,
  size = 22,
  className,
  style,
}: {
  name: string;
  size?: number;
  className?: string;
  style?: CSSProperties;
}) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.8}
      strokeLinecap="round"
      strokeLinejoin="round"
      className={className}
      style={style}
      aria-hidden="true"
    >
      {PATHS[name] ?? PATHS.package}
    </svg>
  );
}
