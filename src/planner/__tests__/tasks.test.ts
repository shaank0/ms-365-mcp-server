import { describe, it, expect, afterEach } from 'vitest';
import { z } from 'zod';
import { deletePlannerTaskTool } from '../tools/tasks.js';
import { fakeGraph, preconditionFailed } from './fake-graph.js';

describe('delete-planner-task', () => {
  it('reads the ETag then deletes with If-Match', async () => {
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: 'T1' }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    const del = g.calls.find((c) => c.options.method === 'DELETE');
    expect(del?.options.headers['If-Match']).toBe('W/"1"');
  });

  it('retries once on a 412', async () => {
    let attempts = 0;
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': () => {
        attempts += 1;
        if (attempts === 1) throw preconditionFailed();
        return {};
      },
    });
    const result = await deletePlannerTaskTool.execute({ taskId: 'T1' }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    expect(attempts).toBe(2);
  });

  it('returns a friendly error on a persistent 412', async () => {
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': () => {
        throw preconditionFailed();
      },
    });
    const result = await deletePlannerTaskTool.execute({ taskId: 'T1' }, { graphClient: g });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/modified by someone else/i);
  });
});

describe('delete-planner-task confirm gate', () => {
  const ORIGINAL = process.env.MS365_MCP_REQUIRE_CONFIRM;

  afterEach(() => {
    if (ORIGINAL === undefined) delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    else process.env.MS365_MCP_REQUIRE_CONFIRM = ORIGINAL;
  });

  it('gate off (default): deletes without a confirm param', async () => {
    delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: 'T1' }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'DELETE')).toBe(true);
  });

  it('gate on: refuses without confirm: true and never touches Graph', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: 'T1' }, { graphClient: g });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/confirmation_required/);
    expect(g.calls).toHaveLength(0);
  });

  it('gate on: proceeds with confirm: true', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph({
      'GET /planner/tasks/T1': { id: 'T1', '@odata.etag': 'W/"1"' },
      'DELETE /planner/tasks/T1': {},
    });
    const result = await deletePlannerTaskTool.execute(
      { taskId: 'T1', confirm: true },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'DELETE')).toBe(true);
  });
});

describe('delete-planner-task schema', () => {
  const schema = z.object(deletePlannerTaskTool.buildSchema());

  it('accepts a well-formed 28-character Planner task id', () => {
    expect(schema.safeParse({ taskId: 'A'.repeat(28) }).success).toBe(true);
  });

  it('rejects a path-traversal-style id', () => {
    const result = schema.safeParse({ taskId: '../../etc/passwd' });
    expect(result.success).toBe(false);
  });

  it('rejects an id containing a slash', () => {
    const result = schema.safeParse({ taskId: 'abc/def' });
    expect(result.success).toBe(false);
  });
});
