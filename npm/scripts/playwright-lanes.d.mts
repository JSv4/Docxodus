export const LANES: readonly string[];
export const serialLaneSpecs: readonly string[];
export function splitSpecFiles(specFiles: readonly string[], index: number, total: number): string[];
export function planLaneRuns(
  args: readonly string[],
  options?: { lanes?: readonly string[]; specFiles?: readonly string[] },
): { lane: string; env: Record<string, string>; args: string[] }[];
