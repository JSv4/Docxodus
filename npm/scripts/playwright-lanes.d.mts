export const LANES: readonly string[];
export const serialLaneSpecs: readonly string[];
export function planLaneRuns(
  args: readonly string[],
  lanes?: readonly string[],
): { lane: string; env: Record<string, string>; args: string[] }[];
