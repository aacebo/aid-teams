import type { JwtClaims } from './types';
import { normalizeAudience } from './normalize-audience';

export function formatAudience(audience: JwtClaims['aud']): string {
  const values = normalizeAudience(audience);

  return values.length === 0 ? 'unknown' : values.join(',');
}
