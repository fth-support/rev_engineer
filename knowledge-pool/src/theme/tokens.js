// Design tokens for the ServiceTransfer Documentation Portal.
// Each theme is a flat map of camelCase keys -> CSS values.
// ThemeContext converts every key to a `--kebab-case` CSS custom property on :root,
// so any component (or raw CSS) can read e.g. var(--accent), var(--node-bg).
//
// Modeled on ag-kit-main/uxui/src/theme.ts but tailored for a docs portal:
// shared semantic tokens + a fixed set of "role" colors used by the diagrams
// (branch=blue, central=green, member=orange, safenet=red, token=red, config=gray).

export const dark = {
  // page / surfaces
  pageBg: '#0b1120',
  pageBgAlt: '#0f172a',
  surface: 'rgba(30, 41, 59, 0.72)',
  surfaceSolid: '#111c33',
  surfaceRaised: 'rgba(51, 65, 85, 0.55)',
  surfaceSunken: 'rgba(2, 6, 23, 0.45)',
  border: 'rgba(148, 163, 184, 0.16)',
  borderStrong: 'rgba(148, 163, 184, 0.32)',

  // text
  textPrimary: '#f1f5f9',
  textSecondary: '#94a3b8',
  textMuted: '#64748b',

  // accent
  accent: '#60a5fa',
  accentStrong: '#3b82f6',
  accentSoft: 'rgba(59, 130, 246, 0.16)',
  accentContrast: '#0b1120',

  // role colors (diagram domains) — strong + soft tint
  roleBranch: '#60a5fa',
  roleBranchSoft: 'rgba(96, 165, 250, 0.14)',
  roleCentral: '#34d399',
  roleCentralSoft: 'rgba(52, 211, 153, 0.14)',
  roleMember: '#fbbf24',
  roleMemberSoft: 'rgba(251, 191, 36, 0.14)',
  roleSafenet: '#f87171',
  roleSafenetSoft: 'rgba(248, 113, 113, 0.14)',
  roleToken: '#fb7185',
  roleTokenSoft: 'rgba(251, 113, 133, 0.14)',
  roleConfig: '#94a3b8',
  roleConfigSoft: 'rgba(148, 163, 184, 0.14)',
  rolePurple: '#c084fc',
  rolePurpleSoft: 'rgba(192, 132, 252, 0.14)',

  // node / graph
  nodeBg: 'rgba(30, 41, 59, 0.85)',
  nodeBorder: 'rgba(148, 163, 184, 0.22)',
  nodeShadow: '0 6px 22px rgba(0, 0, 0, 0.45)',
  nodeHoverGlow: '0 0 26px rgba(96, 165, 250, 0.22)',
  nodeSelectedBorder: 'rgba(96, 165, 250, 0.7)',
  nodeSelectedGlow: '0 0 30px rgba(96, 165, 250, 0.28), 0 0 0 1px rgba(96, 165, 250, 0.4)',
  connectionLine: 'rgba(148, 163, 184, 0.35)',
  connectionHighlight: 'rgba(96, 165, 250, 0.85)',
  particle: 'rgba(96, 165, 250, 0.9)',
  gridDot: 'rgba(148, 163, 184, 0.06)',

  // status
  success: '#34d399',
  successBg: 'rgba(52, 211, 153, 0.12)',
  warning: '#fbbf24',
  warningBg: 'rgba(251, 191, 36, 0.12)',
  danger: '#f87171',
  dangerBg: 'rgba(248, 113, 113, 0.12)',
  info: '#60a5fa',
  infoBg: 'rgba(96, 165, 250, 0.12)',

  // code / misc
  codeBg: 'rgba(2, 6, 23, 0.6)',
  codeText: '#e2e8f0',
  badgeBg: 'rgba(96, 165, 250, 0.14)',
  badgeText: '#93c5fd',
  shadowSm: '0 1px 3px rgba(0, 0, 0, 0.4)',
  shadowMd: '0 8px 30px rgba(0, 0, 0, 0.35)',
  shadowLg: '0 20px 60px rgba(0, 0, 0, 0.5)',
  scrollbar: 'rgba(148, 163, 184, 0.28)',
}

export const light = {
  pageBg: '#eef2f8',
  pageBgAlt: '#e6ecf5',
  surface: 'rgba(255, 255, 255, 0.86)',
  surfaceSolid: '#ffffff',
  surfaceRaised: '#ffffff',
  surfaceSunken: '#f1f5f9',
  border: 'rgba(15, 23, 42, 0.1)',
  borderStrong: 'rgba(15, 23, 42, 0.2)',

  textPrimary: '#0f172a',
  textSecondary: '#475569',
  textMuted: '#64748b',

  accent: '#2563eb',
  accentStrong: '#1d4ed8',
  accentSoft: 'rgba(37, 99, 235, 0.1)',
  accentContrast: '#ffffff',

  roleBranch: '#1d4ed8',
  roleBranchSoft: 'rgba(29, 78, 216, 0.1)',
  roleCentral: '#047857',
  roleCentralSoft: 'rgba(4, 120, 87, 0.1)',
  roleMember: '#b45309',
  roleMemberSoft: 'rgba(180, 83, 9, 0.1)',
  roleSafenet: '#b91c1c',
  roleSafenetSoft: 'rgba(185, 28, 28, 0.08)',
  roleToken: '#be123c',
  roleTokenSoft: 'rgba(190, 18, 60, 0.08)',
  roleConfig: '#475569',
  roleConfigSoft: 'rgba(71, 85, 105, 0.1)',
  rolePurple: '#7c3aed',
  rolePurpleSoft: 'rgba(124, 58, 237, 0.1)',

  nodeBg: '#ffffff',
  nodeBorder: 'rgba(15, 23, 42, 0.14)',
  nodeShadow: '0 4px 16px rgba(15, 23, 42, 0.08)',
  nodeHoverGlow: '0 0 18px rgba(37, 99, 235, 0.16)',
  nodeSelectedBorder: '#2563eb',
  nodeSelectedGlow: '0 0 22px rgba(37, 99, 235, 0.18), 0 0 0 1.5px #2563eb',
  connectionLine: 'rgba(100, 116, 139, 0.5)',
  connectionHighlight: '#2563eb',
  particle: '#2563eb',
  gridDot: 'rgba(15, 23, 42, 0.05)',

  success: '#059669',
  successBg: 'rgba(5, 150, 105, 0.1)',
  warning: '#b45309',
  warningBg: 'rgba(180, 83, 9, 0.1)',
  danger: '#dc2626',
  dangerBg: 'rgba(220, 38, 38, 0.08)',
  info: '#2563eb',
  infoBg: 'rgba(37, 99, 235, 0.1)',

  codeBg: '#f1f5f9',
  codeText: '#0f172a',
  badgeBg: 'rgba(37, 99, 235, 0.1)',
  badgeText: '#1d4ed8',
  shadowSm: '0 1px 3px rgba(15, 23, 42, 0.08)',
  shadowMd: '0 8px 30px rgba(15, 23, 42, 0.1)',
  shadowLg: '0 20px 50px rgba(15, 23, 42, 0.16)',
  scrollbar: 'rgba(15, 23, 42, 0.22)',
}

export const themes = { dark, light }

// camelCase -> --kebab-case
export function toCssVarName(key) {
  return '--' + key.replace(/[A-Z]/g, (c) => '-' + c.toLowerCase())
}

export function themeToCssVars(theme) {
  const out = {}
  for (const [k, v] of Object.entries(theme)) {
    out[toCssVarName(k)] = v
  }
  return out
}
