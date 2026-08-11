import { vi } from 'vitest';

type Route = (options: Record<string, any>) => unknown;

/**
 * Fake GraphClient. Routes are keyed "METHOD /path-prefix" and matched by
 * method plus endpoint prefix, so query strings do not have to be repeated.
 * An unmatched call throws in the same shape GraphClient.makeRequest does.
 */
export function fakeGraph(routes: Record<string, unknown | Route>) {
  const calls: Array<{ endpoint: string; options: Record<string, any> }> = [];

  const makeRequest = vi.fn(async (endpoint: string, options: Record<string, any> = {}) => {
    calls.push({ endpoint, options });
    const method = (options.method ?? 'GET').toUpperCase();
    for (const [key, value] of Object.entries(routes)) {
      const [routeMethod, routePath] = key.split(' ');
      if (routeMethod === method && endpoint.startsWith(routePath)) {
        return typeof value === 'function' ? (value as Route)(options) : value;
      }
    }
    throw new Error(
      `Microsoft Graph API error: 404 Not Found - no fake route for ${method} ${endpoint}`
    );
  });

  return { makeRequest, calls };
}

/** Error in the exact shape GraphClient.makeRequest throws on a 412. */
export function preconditionFailed(): Error {
  return new Error('Microsoft Graph API error: 412 Precondition Failed - {"error":{"code":""}}');
}
