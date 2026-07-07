import type { JwtClaims } from './types';

export function normalizeAudience(audience: JwtClaims['aud']): string[] {
  if (!audience) return [];
  if (typeof audience === 'string') return [audience];

  return audience.filter((value): value is string => typeof value === 'string');
}
