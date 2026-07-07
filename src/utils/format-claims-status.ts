import type { JwtClaims } from './types';
import { formatAudience } from './format-audience';

export function formatClaimsStatus(label: string, claims: JwtClaims | undefined): string {
  return [
    `${label}:`,
    `aud=${formatAudience(claims?.aud)}`,
    `oid=${claims?.oid ?? 'unknown'}`,
    `tid=${claims?.tid ?? 'unknown'}`,
    `user=${claims?.preferred_username ?? claims?.upn ?? claims?.name ?? 'unknown'}`,
    `scp=${claims?.scp ?? 'none'}`,
  ].join(' ');
}
