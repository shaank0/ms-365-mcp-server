import { describe, it, expect } from 'vitest';
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
