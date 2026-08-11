import { describe, it, expect, vi } from 'vitest';
import {
  getJson,
  getEtag,
  patchWithEtag,
  deleteWithEtag,
  paginate,
  isPreconditionFailed,
  withEtagRetry,
  graphStatus,
  PLANNER_ID,
  GUID,
} from '../graph.js';
import { fakeGraph, preconditionFailed, graphError, graphScopeError } from './fake-graph.js';

describe('id patterns', () => {
  it('recognises a 28-char Planner id and rejects a GUID', () => {
    expect(PLANNER_ID.test('IbwGMWqfsUy1MJQsY6XBvWQABDrS')).toBe(true);
    expect(PLANNER_ID.test('38efbc25-b868-481e-85af-88d7caf77779')).toBe(false);
  });

  it('recognises a GUID and rejects a Planner id', () => {
    expect(GUID.test('38efbc25-b868-481e-85af-88d7caf77779')).toBe(true);
    expect(GUID.test('IbwGMWqfsUy1MJQsY6XBvWQABDrS')).toBe(false);
  });
});

describe('getEtag', () => {
  it('prefers the @odata.etag body field', async () => {
    const g = fakeGraph({ 'GET /planner/tasks/T': { '@odata.etag': 'W/"abc"' } });
    expect(await getEtag(g, '/planner/tasks/T')).toBe('W/"abc"');
  });

  it('falls back to the _etag header field', async () => {
    const g = fakeGraph({ 'GET /planner/tasks/T': { _etag: 'W/"xyz"' } });
    expect(await getEtag(g, '/planner/tasks/T')).toBe('W/"xyz"');
  });

  it('throws when neither is present', async () => {
    const g = fakeGraph({ 'GET /planner/tasks/T': { id: 'T' } });
    await expect(getEtag(g, '/planner/tasks/T')).rejects.toThrow(/no ETag/i);
  });
});

describe('patchWithEtag / deleteWithEtag', () => {
  it('sends If-Match and a JSON string body on PATCH', async () => {
    const g = fakeGraph({ 'PATCH /planner/tasks/T': {} });
    await patchWithEtag(g, '/planner/tasks/T', { title: 'x' }, 'W/"abc"');
    const call = g.calls[0];
    expect(call.options.method).toBe('PATCH');
    expect(call.options.headers['If-Match']).toBe('W/"abc"');
    expect(call.options.body).toBe(JSON.stringify({ title: 'x' }));
  });

  it('sends If-Match on DELETE', async () => {
    const g = fakeGraph({ 'DELETE /planner/tasks/T': {} });
    await deleteWithEtag(g, '/planner/tasks/T', 'W/"abc"');
    expect(g.calls[0].options.method).toBe('DELETE');
    expect(g.calls[0].options.headers['If-Match']).toBe('W/"abc"');
  });
});

describe('graphStatus', () => {
  it('parses the status from the "Microsoft Graph API error:" prefix', () => {
    expect(graphStatus(graphError(403, 'Forbidden'))).toBe(403);
  });

  it('parses the status from the "Microsoft Graph API scope error:" prefix', () => {
    expect(graphStatus(graphScopeError(401, 'Unauthorized'))).toBe(401);
  });

  it('does not match a digit that only appears in the body', () => {
    const trap = graphError(
      400,
      'Bad Request',
      '{"error":{"code":"BadRequest","message":"Value 412 exceeds max allowed length."}}'
    );
    expect(graphStatus(trap)).toBe(400);
  });

  it('returns null for a non-Graph error', () => {
    expect(graphStatus(new Error('socket hang up'))).toBeNull();
  });
});

describe('isPreconditionFailed', () => {
  it('detects a 412 error', () => {
    expect(isPreconditionFailed(preconditionFailed())).toBe(true);
  });

  it('ignores other errors', () => {
    expect(isPreconditionFailed(new Error('403 Forbidden'))).toBe(false);
  });

  it('does not false-positive when the body merely contains "412"', () => {
    const trap = graphError(
      400,
      'Bad Request',
      '{"error":{"code":"BadRequest","message":"Value 412 exceeds max allowed length."}}'
    );
    expect(isPreconditionFailed(trap)).toBe(false);
  });
});

describe('withEtagRetry', () => {
  it('does not retry when the write succeeds', async () => {
    const read = vi.fn(async () => 'W/"1"');
    const write = vi.fn(async () => 'done');
    expect(await withEtagRetry(read, write)).toBe('done');
    expect(read).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledTimes(1);
  });

  it('refetches once and retries on 412', async () => {
    const read = vi.fn().mockResolvedValueOnce('W/"1"').mockResolvedValueOnce('W/"2"');
    const write = vi.fn().mockRejectedValueOnce(preconditionFailed()).mockResolvedValueOnce('done');
    expect(await withEtagRetry(read, write)).toBe('done');
    expect(read).toHaveBeenCalledTimes(2);
    expect(write).toHaveBeenNthCalledWith(2, 'W/"2"');
  });

  it('propagates a second 412', async () => {
    const read = vi.fn(async () => 'W/"1"');
    const write = vi.fn().mockRejectedValue(preconditionFailed());
    await expect(withEtagRetry(read, write)).rejects.toThrow(/412/);
    expect(write).toHaveBeenCalledTimes(2);
  });

  it('does not retry a non-412 failure', async () => {
    const read = vi.fn(async () => 'W/"1"');
    const write = vi.fn().mockRejectedValue(new Error('403 Forbidden'));
    await expect(withEtagRetry(read, write)).rejects.toThrow(/403/);
    expect(write).toHaveBeenCalledTimes(1);
  });

  it('does not retry a 400 whose body merely contains "412"', async () => {
    const read = vi.fn(async () => 'W/"1"');
    const trap = graphError(
      400,
      'Bad Request',
      '{"error":{"code":"BadRequest","message":"Value 412 exceeds max allowed length."}}'
    );
    const write = vi.fn().mockRejectedValue(trap);
    await expect(withEtagRetry(read, write)).rejects.toThrow(/400/);
    expect(write).toHaveBeenCalledTimes(1);
  });
});

describe('paginate', () => {
  it('follows @odata.nextLink and strips the absolute prefix', async () => {
    // Routes are matched in declaration order by prefix, so the more
    // specific page-2 route must come first.
    const g = fakeGraph({
      'GET /x?page=2': { value: [{ id: 'b' }] },
      'GET /x': {
        value: [{ id: 'a' }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?page=2',
      },
    });
    const result = await paginate<{ id: string }>(g, '/x', 100);
    expect(result.items.map((i) => i.id)).toEqual(['a', 'b']);
    expect(result.truncated).toBe(false);
    // The absolute nextLink must have been reduced to a relative path.
    expect(g.calls[1].endpoint).toBe('/x?page=2');
  });

  it('stops at the cap and reports truncation', async () => {
    const g = fakeGraph({
      'GET /x': {
        value: [{ id: 'a' }, { id: 'b' }, { id: 'c' }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?page=2',
      },
    });
    const result = await paginate<{ id: string }>(g, '/x', 2);
    expect(result.items.length).toBe(2);
    expect(result.truncated).toBe(true);
  });
});

describe('getJson', () => {
  it('returns the parsed body', async () => {
    const g = fakeGraph({ 'GET /planner/plans/P': { id: 'P', title: 'Q3' } });
    expect(await getJson<{ title: string }>(g, '/planner/plans/P')).toEqual({
      id: 'P',
      title: 'Q3',
    });
  });
});
