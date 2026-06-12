// ─── Types ───────────────────────────────────────────────────────────────────

export interface Quest {
  id: string;
  icon: string;
  title: string;
  description: string;
  xpReward: number;
  completed: boolean;
}
