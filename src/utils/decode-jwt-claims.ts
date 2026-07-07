import type { JwtClaims } from './types';

export function decodeJwtClaims(token: string): JwtClaims | undefined {
  const payload = token.split('.')[1];

  if (!payload) return undefined;

  try {
    const normalized = payload.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, '=');

    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8')) as JwtClaims;
  } catch {
    return undefined;
  }
}
