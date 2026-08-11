/** Group and user ids are GUIDs. */
export const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Plan, bucket and task ids are 28-character opaque strings, NOT GUIDs. */
export const PLANNER_ID = /^[A-Za-z0-9_-]{28}$/;

/** The slice of GraphClient these helpers need. */
export interface GraphLike {
  makeRequest(endpoint: string, options?: Record<string, any>): Promise<any>;
}

export async function getJson<T>(g: GraphLike, endpoint: string): Promise<T> {
  return (await g.makeRequest(endpoint)) as T;
}

export async function postJson<T>(g: GraphLike, endpoint: string, body: unknown): Promise<T> {
  return (await g.makeRequest(endpoint, {
    method: 'POST',
    body: JSON.stringify(body),
  })) as T;
}

/** Read an entity's ETag. Planner returns it in the body as @odata.etag. */
export async function getEtag(g: GraphLike, endpoint: string): Promise<string> {
  const entity = await g.makeRequest(endpoint, { includeHeaders: true });
  const etag = entity?.['@odata.etag'] ?? entity?._etag;
  if (!etag || etag === 'no-etag-found') {
    throw new Error(`Graph returned no ETag for ${endpoint}; cannot safely write to it.`);
  }
  return etag as string;
}

export async function patchWithEtag(
  g: GraphLike,
  endpoint: string,
  body: unknown,
  etag: string
): Promise<void> {
  await g.makeRequest(endpoint, {
    method: 'PATCH',
    headers: { 'If-Match': etag },
    body: JSON.stringify(body),
  });
}

export async function deleteWithEtag(g: GraphLike, endpoint: string, etag: string): Promise<void> {
  await g.makeRequest(endpoint, {
    method: 'DELETE',
    headers: { 'If-Match': etag },
  });
}

export function isPreconditionFailed(err: unknown): boolean {
  return /\b412\b/.test(err instanceof Error ? err.message : String(err));
}

/**
 * Run a write under optimistic concurrency: read the ETag, write, and on a
 * 412 refetch and retry EXACTLY once. A second 412 propagates so the caller
 * can report a genuine concurrent edit.
 */
export async function withEtagRetry<T>(
  read: () => Promise<string>,
  write: (etag: string) => Promise<T>
): Promise<T> {
  try {
    return await write(await read());
  } catch (err) {
    if (!isPreconditionFailed(err)) throw err;
    return await write(await read());
  }
}

/**
 * Follow @odata.nextLink up to `cap` items. Planner ignores $top, so the cap
 * is enforced here and reported to the caller.
 */
export async function paginate<T>(
  g: GraphLike,
  endpoint: string,
  cap: number
): Promise<{ items: T[]; truncated: boolean }> {
  const items: T[] = [];
  let next: string | undefined = endpoint;
  let truncated = false;

  while (next) {
    const page: any = await g.makeRequest(next);
    for (const item of (page?.value ?? []) as T[]) {
      if (items.length >= cap) {
        truncated = true;
        return { items, truncated };
      }
      items.push(item);
    }
    const link: string | undefined = page?.['@odata.nextLink'];
    next = link ? link.replace(/^https:\/\/[^/]+\/(v1\.0|beta)/, '') : undefined;
  }

  return { items, truncated };
}
