import { ExpressAdapter, type HttpMethod, type HttpRouteHandler } from '@microsoft/teams.apps/dist/http';
import { randomUUID } from 'crypto';
import type { HeaderMap } from './utils/index';

export const RAW_REQUEST_ID_FIELD = '__aidTeamsRequestId';
export const requestHeadersById = new Map<string, HeaderMap>();

export class RequestCapturingExpressAdapter extends ExpressAdapter {
  override registerRoute(method: HttpMethod, path: string, handler: HttpRouteHandler): void {
    super.registerRoute(method, path, async (request) => {
      const requestId = randomUUID();
      const requestBody = request.body;
      let body = requestBody;
      const normalizedHeaders: Record<string, string[]> = {};

      if (requestBody && typeof requestBody === 'object' && !Array.isArray(requestBody)) {
        const bodyRecord = requestBody as Record<string, unknown>;
        const channelData = bodyRecord.channelData;
        body = {
          ...bodyRecord,
          channelData: {
            ...(channelData && typeof channelData === 'object' && !Array.isArray(channelData) ? channelData : {}),
            [RAW_REQUEST_ID_FIELD]: requestId,
          },
        };
      }

      for (const [name, value] of Object.entries(request.headers)) {
        normalizedHeaders[name.toLowerCase()] = Array.isArray(value) ? value : [value];
      }

      requestHeadersById.set(requestId, normalizedHeaders);

      try {
        return await handler({ ...request, body });
      } finally {
        requestHeadersById.delete(requestId);
      }
    });
  }
}
