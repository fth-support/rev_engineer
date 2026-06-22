// Design tokens for the Project SCO portal.
// Themed on the Foodland Self-Checkout (SCO) brand: warm cream + Foodland red + gold + green.
// Each theme is a flat map of camelCase keys -> CSS values. ThemeContext converts every key
// to a `--kebab-case` CSS custom property on :root, so any component or raw CSS can read
// e.g. var(--accent), var(--accent-2), var(--node-bg).
//
// Brand source: sco-kiosk design — red #E2231A/#C2161A, dark #211B1A, cream #F6F1EE,
// gold #F5B800, green #1F8A4C. role-* hues are kept for the technical diagrams.

export const dark = {
  // page / surfaces — warm charcoal
  pageBg: '#1a1310',
  pageBgAlt: '#211b1a',
  surface: 'rgba(58, 44, 36, 0.55)',
  surfaceSolid: '#241a15',
  surfaceRaised: 'rgba(92, 79, 73, 0.40)',
  surfaceSunken: 'rgba(15, 10, 8, 0.50)',
  border: 'rgba(201, 182, 172, 0.16)',
  borderStrong: 'rgba(201, 182, 172, 0.30)',

  // text — warm cream
  textPrimary: '#f6f1ee',
  textSecondary: '#c9b6ac',
  textMuted: '#8a7e79',

  // accent — Foodland red
  accent: '#e2231a',
  accentStrong: '#c2161a',
  accentSoft: 'rgba(226, 35, 26, 0.16)',
  accentContrast: '#ffffff',
  // secondary brand accent — gold
  accent2: '#f5b800',
  accent2Soft: 'rgba(245, 184, 0, 0.16)',

  // role colors (diagram domains) — strong + soft tint
  roleBranch: '#e2231a',
  roleBranchSoft: 'rgba(226, 35, 26, 0.14)',
  roleCentral: '#2bbf6a',
  roleCentralSoft: 'rgba(43, 191, 106, 0.14)',
  roleMember: '#f5b800',
  roleMemberSoft: 'rgba(245, 184, 0, 0.14)',
  roleSafenet: '#fb7185',
  roleSafenetSoft: 'rgba(251, 113, 133, 0.14)',
  roleToken: '#f59e0b',
  roleTokenSoft: 'rgba(245, 158, 11, 0.14)',
  roleConfig: '#c9b6ac',
  roleConfigSoft: 'rgba(201, 182, 172, 0.14)',
  rolePurple: '#c084fc',
  rolePurpleSoft: 'rgba(192, 132, 252, 0.14)',

  // node / graph
  nodeBg: 'rgba(58, 44, 36, 0.85)',
  nodeBorder: 'rgba(201, 182, 172, 0.22)',
  nodeShadow: '0 6px 22px rgba(0, 0, 0, 0.5)',
  nodeHoverGlow: '0 0 26px rgba(226, 35, 26, 0.22)',
  nodeSelectedBorder: 'rgba(226, 35, 26, 0.7)',
  nodeSelectedGlow: '0 0 30px rgba(226, 35, 26, 0.28), 0 0 0 1px rgba(226, 35, 26, 0.4)',
  connectionLine: 'rgba(201, 182, 172, 0.35)',
  connectionHighlight: 'rgba(226, 35, 26, 0.85)',
  particle: 'rgba(245, 184, 0, 0.95)',
  gridDot: 'rgba(201, 182, 172, 0.07)',

  // status
  success: '#2bbf6a',
  successBg: 'rgba(43, 191, 106, 0.14)',
  warning: '#f5b800',
  warningBg: 'rgba(245, 184, 0, 0.14)',
  danger: '#f87171',
  dangerBg: 'rgba(248, 113, 113, 0.14)',
  info: '#e2231a',
  infoBg: 'rgba(226, 35, 26, 0.14)',

  // code / misc
  codeBg: 'rgba(15, 10, 8, 0.6)',
  codeText: '#f1e6df',
  badgeBg: 'rgba(226, 35, 26, 0.16)',
  badgeText: '#fca5a5',
  shadowSm: '0 1px 3px rgba(0, 0, 0, 0.45)',
  shadowMd: '0 8px 30px rgba(0, 0, 0, 0.4)',
  shadowLg: '0 20px 60px rgba(0, 0, 0, 0.55)',
  scrollbar: 'rgba(201, 182, 172, 0.28)',
}

export const light = {
  // page / surfaces — warm cream
  pageBg: '#f6f1ee',
  pageBgAlt: '#ece3de',
  surface: 'rgba(255, 255, 255, 0.88)',
  surfaceSolid: '#ffffff',
  surfaceRaised: '#fdfbf9',
  surfaceSunken: '#f1e6df',
  border: 'rgba(58, 44, 36, 0.12)',
  borderStrong: 'rgba(58, 44, 36, 0.22)',

  textPrimary: '#211b1a',
  textSecondary: '#5c4f49',
  textMuted: '#8a7e79',

  accent: '#e2231a',
  accentStrong: '#c2161a',
  accentSoft: 'rgba(226, 35, 26, 0.10)',
  accentContrast: '#ffffff',
  accent2: '#d69a00',
  accent2Soft: 'rgba(214, 154, 0, 0.12)',

  roleBranch: '#c2161a',
  roleBranchSoft: 'rgba(194, 22, 26, 0.10)',
  roleCentral: '#1a7a42',
  roleCentralSoft: 'rgba(26, 122, 66, 0.10)',
  roleMember: '#b8860b',
  roleMemberSoft: 'rgba(184, 134, 11, 0.12)',
  roleSafenet: '#be123c',
  roleSafenetSoft: 'rgba(190, 18, 60, 0.08)',
  roleToken: '#b45309',
  roleTokenSoft: 'rgba(180, 83, 9, 0.10)',
  roleConfig: '#5c4f49',
  roleConfigSoft: 'rgba(92, 79, 73, 0.10)',
  rolePurple: '#7c3aed',
  rolePurpleSoft: 'rgba(124, 58, 237, 0.10)',

  nodeBg: '#ffffff',
  nodeBorder: 'rgba(58, 44, 36, 0.14)',
  nodeShadow: '0 4px 16px rgba(58, 44, 36, 0.10)',
  nodeHoverGlow: '0 0 18px rgba(226, 35, 26, 0.16)',
  nodeSelectedBorder: '#e2231a',
  nodeSelectedGlow: '0 0 22px rgba(226, 35, 26, 0.18), 0 0 0 1.5px #e2231a',
  connectionLine: 'rgba(138, 126, 121, 0.5)',
  connectionHighlight: '#e2231a',
  particle: '#d69a00',
  gridDot: 'rgba(58, 44, 36, 0.06)',

  success: '#1a7a42',
  successBg: 'rgba(26, 122, 66, 0.10)',
  warning: '#b8860b',
  warningBg: 'rgba(184, 134, 11, 0.12)',
  danger: '#dc2626',
  dangerBg: 'rgba(220, 38, 38, 0.08)',
  info: '#c2161a',
  infoBg: 'rgba(194, 22, 26, 0.10)',

  codeBg: '#f1e6df',
  codeText: '#211b1a',
  badgeBg: 'rgba(226, 35, 26, 0.10)',
  badgeText: '#c2161a',
  shadowSm: '0 1px 3px rgba(58, 44, 36, 0.08)',
  shadowMd: '0 8px 30px rgba(58, 44, 36, 0.10)',
  shadowLg: '0 20px 50px rgba(58, 44, 36, 0.16)',
  scrollbar: 'rgba(58, 44, 36, 0.22)',
}

export const themes = { dark, light }

// camelCase -> --kebab-case
export function toCssVarName(key) {
  return '--' + key.replace(/[A-Z0-9]/g, (c) => '-' + c.toLowerCase())
}

export function themeToCssVars(theme) {
  const out = {}
  for (const [k, v] of Object.entries(theme)) {
    out[toCssVarName(k)] = v
  }
  return out
}
