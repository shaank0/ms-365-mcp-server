import { describe, it, expect, afterEach } from 'vitest';
import { z } from 'zod';
import { deletePlannerTaskTool, createPlannerTaskTool } from '../tools/tasks.js';
import { fakeGraph, preconditionFailed } from './fake-graph.js';

// A valid 28-character Planner id (matches PLANNER_ID), now that execute()
// re-validates params.taskId at runtime (see validate.ts) rather than only
// trusting the Zod schema.
const T1 = 'T'.repeat(28);

describe('delete-planner-task', () => {
  it('reads the ETag then deletes with If-Match', async () => {
    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    const del = g.calls.find((c) => c.options.method === 'DELETE');
    expect(del?.options.headers['If-Match']).toBe('W/"1"');
  });

  it('retries once on a 412', async () => {
    let attempts = 0;
    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: () => {
        attempts += 1;
        if (attempts === 1) throw preconditionFailed();
        return {};
      },
    });
    const result = await deletePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    expect(attempts).toBe(2);
  });

  it('returns a friendly error on a persistent 412', async () => {
    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: () => {
        throw preconditionFailed();
      },
    });
    const result = await deletePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/modified by someone else/i);
  });

  // Discovery mode's execute-tool calls utility.execute(parameters, ctx) directly with
  // raw, unparsed client input - no Zod runs on that path (see graph-tools.ts's
  // execute-tool handler). Prove the runtime check in validate.ts closes that hole
  // independently of the schema.
  it('rejects a traversal-style taskId even when the schema is bypassed, and never calls Graph', async () => {
    const g = fakeGraph({
      'GET /planner/tasks/': { id: 'nope' },
      'DELETE /planner/tasks/': {},
    });
    const result = await deletePlannerTaskTool.execute(
      { taskId: '../../me/messages/XYZ' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/taskId/);
    expect(g.calls).toHaveLength(0);
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
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'DELETE')).toBe(true);
  });

  it('gate on: refuses without confirm: true and never touches Graph', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: {},
    });
    const result = await deletePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/confirmation_required/);
    expect(g.calls).toHaveLength(0);
  });

  it('gate on: proceeds with confirm: true', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, '@odata.etag': 'W/"1"' },
      [`DELETE /planner/tasks/${T1}`]: {},
    });
    const result = await deletePlannerTaskTool.execute(
      { taskId: T1, confirm: true },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'DELETE')).toBe(true);
  });
});

describe('delete-planner-task schema', () => {
  const schema = z.object(deletePlannerTaskTool.buildSchema());

  it('accepts a well-formed 28-character Planner task id', () => {
    expect(schema.safeParse({ taskId: T1 }).success).toBe(true);
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

const PLAN = 'P'.repeat(28);

const CREATE_TASK_ROUTES = {
  'POST /planner/tasks': { id: T1, title: 'Ship it' },
  [`GET /planner/tasks/${T1}/details`]: { id: T1, '@odata.etag': 'W/"d1"' },
  [`PATCH /planner/tasks/${T1}/details`]: {},
};

describe('create-planner-task', () => {
  it('makes a single POST when no description or checklist is given', async () => {
    const g = fakeGraph(CREATE_TASK_ROUTES);
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it' },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.length).toBe(1);
    expect(g.calls[0].options.method).toBe('POST');
  });

  it('writes details in a second call when a description is given', async () => {
    const g = fakeGraph(CREATE_TASK_ROUTES);
    await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', description: 'the details' },
      { graphClient: g }
    );
    const patch = g.calls.find((c) => c.options.method === 'PATCH');
    expect(patch!.options.headers['If-Match']).toBe('W/"d1"');
    expect(JSON.parse(patch!.options.body).description).toBe('the details');
  });

  it('maps a priority word and builds the checklist', async () => {
    const g = fakeGraph(CREATE_TASK_ROUTES);
    await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', priority: 'urgent', checklist: ['a'] },
      { graphClient: g }
    );
    const post = g.calls.find((c) => c.options.method === 'POST');
    expect(JSON.parse(post!.options.body).priority).toBe(1);
    const patch = g.calls.find((c) => c.options.method === 'PATCH');
    const checklist = JSON.parse(patch!.options.body).checklist;
    expect(Object.values(checklist)[0]).toMatchObject({ title: 'a', isChecked: false });
  });

  it('returns success WITH a warning when the details write fails', async () => {
    const g = fakeGraph({
      ...CREATE_TASK_ROUTES,
      [`PATCH /planner/tasks/${T1}/details`]: () => {
        throw new Error('Microsoft Graph API error: 500 Server Error - boom');
      },
    });
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', description: 'the details' },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    const payload = JSON.parse(result.content[0].text);
    expect(payload.id).toBe(T1);
    expect(payload.warning).toMatch(/description/i);
    expect(payload.warning).toMatch(new RegExp(T1));
  });

  it('fails before creating anything when the bucket name is unknown', async () => {
    const g = fakeGraph({
      ...CREATE_TASK_ROUTES,
      [`GET /planner/plans/${PLAN}/buckets`]: { value: [{ id: 'B', name: 'To Do' }] },
    });
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', bucket: 'Backlog' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(g.calls.some((c) => c.options.method === 'POST')).toBe(false);
  });

  // Discovery mode's execute-tool calls utility.execute(parameters, ctx) directly with
  // raw, unparsed client input - no Zod runs on that path (see graph-tools.ts's
  // execute-tool handler). Prove the runtime check in validate.ts closes that hole
  // independently of the schema, same as delete-planner-task above.
  it('rejects a traversal-style planId even when the schema is bypassed, and never calls Graph', async () => {
    const g = fakeGraph(CREATE_TASK_ROUTES);
    const result = await createPlannerTaskTool.execute(
      { planId: '../../me/messages/XYZ', title: 'Ship it' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/planId/);
    expect(g.calls).toHaveLength(0);
  });
});

describe('create-planner-task confirm gate', () => {
  const ORIGINAL = process.env.MS365_MCP_REQUIRE_CONFIRM;

  afterEach(() => {
    if (ORIGINAL === undefined) delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    else process.env.MS365_MCP_REQUIRE_CONFIRM = ORIGINAL;
  });

  it('gate off (default): creates without a confirm param', async () => {
    delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    const g = fakeGraph(CREATE_TASK_ROUTES);
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it' },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'POST')).toBe(true);
  });

  it('gate on: refuses without confirm: true and never touches Graph', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph(CREATE_TASK_ROUTES);
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/confirmation_required/);
    expect(g.calls).toHaveLength(0);
  });

  it('gate on: proceeds with confirm: true', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph(CREATE_TASK_ROUTES);
    const result = await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', confirm: true },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'POST')).toBe(true);
  });
});

describe('create-planner-task schema', () => {
  const schema = z.object(createPlannerTaskTool.buildSchema());

  it('accepts a well-formed request', () => {
    expect(schema.safeParse({ planId: PLAN, title: 'Ship it' }).success).toBe(true);
  });

  it('rejects a path-traversal-style planId', () => {
    const result = schema.safeParse({ planId: '../../etc/passwd', title: 'Ship it' });
    expect(result.success).toBe(false);
  });

  it('rejects a planId containing a slash', () => {
    const result = schema.safeParse({ planId: 'abc/def', title: 'Ship it' });
    expect(result.success).toBe(false);
  });
});
