import { PACKAGES } from "./data";

const BASE_R = 24;
const SCALE_PER_DEP = 3;

const NODE_RADII: Record<string, number> = {};
const NODE_DELAY: Record<string, number> = {};
for (let i = 0; i < PACKAGES.length; i++) {
  const pkg = PACKAGES[i];
  const depCount = PACKAGES.filter((p) =>
    p.dependencies.includes(pkg.id)
  ).length;
  NODE_RADII[pkg.id] = BASE_R + depCount * SCALE_PER_DEP;
  NODE_DELAY[pkg.id] = 100 + i * 30;
}

export function getNodeRadius(id: string): number {
  return NODE_RADII[id] ?? BASE_R;
}

export function getNodeDelay(id: string): number {
  return NODE_DELAY[id] ?? 0;
}
