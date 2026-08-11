import { describe, it, expect } from 'vitest';
import { ok, toolError } from '../result.js';
import { graphError } from './fake-graph.js';

function errorText(result: ReturnType<typeof toolError>): string {
  return JSON.parse(result.content[0].text).error;
}

describe('ok', () => {
  it('serializes the payload as JSON text with no isError flag', () => {
    const result = ok({ id: 'T', title: 'x' });
    expect(result.isError).toBeUndefined();
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'T', title: 'x' });
  });
});

describe('toolError', () => {
  it('does not mislabel a 400 whose body merely contains "412" as a concurrency conflict', () => {
    const trap = graphError(
      400,
      'Bad Request',
      '{"error":{"code":"BadRequest","message":"Value 412 exceeds max allowed length."}}'
    );
    const result = toolError(trap);
    expect(result.isError).toBe(true);
    expect(errorText(result)).not.toMatch(/modified by someone else/i);
    expect(errorText(result)).toBe(trap.message);
  });

  it('does not mislabel a 403 whose body merely contains "404" as not-found', () => {
    const trap = graphError(
      403,
      'Forbidden',
      '{"error":{"code":"Forbidden","message":"See status 404 in related docs."}}'
    );
    const result = toolError(trap);
    expect(errorText(result)).toMatch(/don't have access/i);
    expect(errorText(result)).not.toMatch(/no such planner item/i);
  });

  it('reports a real 412 as a concurrency conflict', () => {
    const result = toolError(graphError(412, 'Precondition Failed'));
    expect(errorText(result)).toMatch(/modified by someone else/i);
  });

  it('reports a real 403 as an access error', () => {
    const result = toolError(graphError(403, 'Forbidden'));
    expect(errorText(result)).toMatch(/don't have access/i);
  });

  it('reports a real 404 as a not-found error', () => {
    const result = toolError(graphError(404, 'Not Found'));
    expect(errorText(result)).toMatch(/no such planner item/i);
  });

  it('passes the original message through for a non-Graph error', () => {
    const err = new Error('socket hang up');
    const result = toolError(err);
    expect(errorText(result)).toBe('socket hang up');
    expect(errorText(result)).not.toMatch(
      /don't have access|no such planner item|modified by someone else/i
    );
  });
});
