import { describe, it, expect, afterEach } from 'vitest';
import { z } from 'zod';
import {
  deletePlannerTaskTool,
  createPlannerTaskTool,
  updatePlannerTaskTool,
} from '../tools/tasks.js';
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

  // Unlike update-planner-task's replacement semantics, create-planner-task targets
  // a brand-new task with nothing to merge against - its checklist body must contain
  // ONLY the new items, never a null entry (toReplacement is update-only; this pins
  // that createPlannerTaskTool never starts calling it).
  it('sends no null entries in the checklist body (nothing to clear on a new task)', async () => {
    const g = fakeGraph(CREATE_TASK_ROUTES);
    await createPlannerTaskTool.execute(
      { planId: PLAN, title: 'Ship it', checklist: ['a', 'b'] },
      { graphClient: g }
    );
    const patch = g.calls.find((c) => c.options.method === 'PATCH');
    const checklist = JSON.parse(patch!.options.body).checklist as Record<string, unknown>;
    expect(Object.keys(checklist)).toHaveLength(2);
    expect(Object.values(checklist).every((v) => v !== null)).toBe(true);
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
    // Discriminate the failure mode from a generic 404: the message must name
    // the bucket that was requested AND list what actually exists in the
    // plan, not just fail for any reason (fakeGraph 404s any unrouted call,
    // so `isError: true` alone would pass even if the buckets route were
    // missing or resolveBucket failed for an unrelated reason).
    expect(result.content[0].text).toMatch(/Backlog/);
    expect(result.content[0].text).toMatch(/To Do/);
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

  // The bucket/priority/checklist tests above each pin ONE resolved value in
  // isolation, which does not prove it landed in the RIGHT body field - e.g.
  // a bucketId/assignments/dueDateTime/startDateTime typo in the wiring would
  // pass every test above (resolveBucket/resolveAssignees are correct in
  // isolation; toPriority is correct in isolation) as long as nothing reads
  // that specific field back out of the POST body. Supply every wired field
  // at once and assert each lands under its correct Graph body key.
  it('wires bucket, assignees, dates and priority into the correct POST body fields', async () => {
    const groupId = '11111111-1111-1111-1111-111111111111';
    const memberId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
    const bucketId = 'B'.repeat(28);

    const g = fakeGraph({
      'POST /planner/tasks': { id: T1, title: 'Ship it' },
      // Longer/more specific paths must be listed before shorter prefixes of
      // themselves - fakeGraph matches by endpoint.startsWith(routePath), so
      // `/planner/plans/${PLAN}` would otherwise swallow the `/buckets` call.
      [`GET /planner/plans/${PLAN}/buckets`]: { value: [{ id: bucketId, name: 'To Do' }] },
      [`GET /planner/plans/${PLAN}`]: {
        id: PLAN,
        container: { containerId: groupId, type: 'group' },
      },
      [`GET /groups/${groupId}/members`]: {
        value: [{ id: memberId, mail: 'will@kw-corp.com', userPrincipalName: 'will@kw-corp.com' }],
      },
    });

    const result = await createPlannerTaskTool.execute(
      {
        planId: PLAN,
        title: 'Ship it',
        bucket: 'To Do',
        assignees: ['will@kw-corp.com'],
        dueDateTime: '2026-09-01T00:00:00Z',
        startDateTime: '2026-08-01T00:00:00Z',
        priority: 'urgent',
      },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();

    const post = g.calls.find(
      (c) => c.options.method === 'POST' && c.endpoint === '/planner/tasks'
    );
    const body = JSON.parse(post!.options.body);
    expect(body.bucketId).toBe(bucketId);
    expect(body.assignments).toEqual({
      [memberId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
    });
    expect(body.dueDateTime).toBe('2026-09-01T00:00:00Z');
    expect(body.startDateTime).toBe('2026-08-01T00:00:00Z');
    expect(body.priority).toBe(1);
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

// NOTE on route ordering: fakeGraph matches by method + endpoint PREFIX in
// declaration order (see fake-graph.ts). "/planner/tasks/T...T" is a PREFIX
// of "/planner/tasks/T...T/details", so for both GET and PATCH the /details
// route MUST be declared first below - otherwise the bare-task route would
// shadow it and a "details" call would silently be answered by the task
// route's fixture instead, making a test assert against the wrong call.
const UPDATE_ROUTES = {
  [`GET /planner/tasks/${T1}/details`]: { '@odata.etag': 'W/"d1"' },
  [`GET /planner/tasks/${T1}`]: { id: T1, title: 'Ship it', '@odata.etag': 'W/"t1"' },
  [`PATCH /planner/tasks/${T1}/details`]: {},
  [`PATCH /planner/tasks/${T1}`]: {},
};

describe('update-planner-task', () => {
  it('patches the task with its own ETag', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    await updatePlannerTaskTool.execute({ taskId: T1, title: 'Renamed' }, { graphClient: g });
    const patch = g.calls.find(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    expect(patch!.options.headers['If-Match']).toBe('W/"t1"');
    expect(JSON.parse(patch!.options.body).title).toBe('Renamed');
  });

  it('maps status to percentComplete', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    await updatePlannerTaskTool.execute({ taskId: T1, status: 'complete' }, { graphClient: g });
    const patch = g.calls.find(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    expect(JSON.parse(patch!.options.body).percentComplete).toBe(100);
  });

  it('accepts percentComplete: 0 and priority: 0, not silently dropped by a truthy check', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    await updatePlannerTaskTool.execute(
      { taskId: T1, percentComplete: 0, priority: 0 },
      { graphClient: g }
    );
    const patch = g.calls.find(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    const body = JSON.parse(patch!.options.body);
    expect(body.percentComplete).toBe(0);
    expect(body.priority).toBe(0);
  });

  it('an explicit percentComplete overrides status when both are given', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    await updatePlannerTaskTool.execute(
      { taskId: T1, status: 'complete', percentComplete: 42 },
      { graphClient: g }
    );
    const patch = g.calls.find(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    expect(JSON.parse(patch!.options.body).percentComplete).toBe(42);
  });

  it('touches only the details entity when only a description is given', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    await updatePlannerTaskTool.execute(
      { taskId: T1, description: 'new text' },
      { graphClient: g }
    );
    const taskPatches = g.calls.filter(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    expect(taskPatches.length).toBe(0);
    const detailPatch = g.calls.find(
      (c) => c.options.method === 'PATCH' && c.endpoint.endsWith('/details')
    );
    expect(JSON.parse(detailPatch!.options.body).description).toBe('new text');
    // Distinguish the details ETag (W/"d1") from the task ETag (W/"t1") - if
    // applyDetails were ever handed the TASK's ETag instead of the details
    // entity's own, this would still pass without this explicit check.
    expect(detailPatch!.options.headers['If-Match']).toBe('W/"d1"');
  });

  it('errors when no updatable field is supplied', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    const result = await updatePlannerTaskTool.execute({ taskId: T1 }, { graphClient: g });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/at least one/i);
  });

  it('warns when the task updated but details failed', async () => {
    const g = fakeGraph({
      ...UPDATE_ROUTES,
      [`PATCH /planner/tasks/${T1}/details`]: () => {
        throw new Error('Microsoft Graph API error: 500 Server Error - boom');
      },
    });
    const result = await updatePlannerTaskTool.execute(
      { taskId: T1, title: 'Renamed', description: 'new text' },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    const payload = JSON.parse(result.content[0].text);
    expect(payload.warning).toMatch(/description/i);
    expect(payload.warning).toMatch(/update-planner-task/);
    expect(payload.warning).toMatch(new RegExp(T1));
  });

  // Discovery mode's execute-tool calls utility.execute(parameters, ctx) directly with
  // raw, unparsed client input - no Zod runs on that path (see graph-tools.ts's
  // execute-tool handler). Prove the runtime check in validate.ts closes that hole
  // independently of the schema, same as delete/create-planner-task above.
  it('rejects a traversal-style taskId even when the schema is bypassed, and never calls Graph', async () => {
    const g = fakeGraph(UPDATE_ROUTES);
    const result = await updatePlannerTaskTool.execute(
      { taskId: '../../me/messages/XYZ', title: 'Renamed' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/taskId/);
    expect(g.calls).toHaveLength(0);
  });

  // Mirror of create-planner-task's "wires bucket, assignees, dates and priority
  // into the correct POST body fields" test above: each field-in-isolation test
  // proves resolveBucket/resolveAssignees/toPriority are correct in isolation, but
  // not that the result landed under the RIGHT body key - a bucketId/assignments
  // typo in the wiring would pass every test above. The task here starts with no
  // existing assignees, so the resolved assignment is a pure addition (no nulled
  // keys), isolating the "does it land in the right field" question from the
  // merge/replacement behavior covered separately below.
  it('wires bucket, assignees, dates, priority, status and percentComplete into the correct PATCH body fields', async () => {
    const groupId = '11111111-1111-1111-1111-111111111111';
    const memberId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
    const bucketId = 'B'.repeat(28);

    const g = fakeGraph({
      [`GET /planner/tasks/${T1}`]: { id: T1, planId: PLAN, '@odata.etag': 'W/"t1"' },
      // Longer/more specific paths must be listed before shorter prefixes of
      // themselves - fakeGraph matches by endpoint.startsWith(routePath), so
      // `/planner/plans/${PLAN}` would otherwise swallow the `/buckets` call.
      [`GET /planner/plans/${PLAN}/buckets`]: { value: [{ id: bucketId, name: 'To Do' }] },
      [`GET /planner/plans/${PLAN}`]: {
        id: PLAN,
        container: { containerId: groupId, type: 'group' },
      },
      [`GET /groups/${groupId}/members`]: {
        value: [{ id: memberId, mail: 'will@kw-corp.com', userPrincipalName: 'will@kw-corp.com' }],
      },
      [`PATCH /planner/tasks/${T1}`]: {},
    });

    const result = await updatePlannerTaskTool.execute(
      {
        taskId: T1,
        bucket: 'To Do',
        assignees: ['will@kw-corp.com'],
        dueDateTime: '2026-09-01T00:00:00Z',
        startDateTime: '2026-08-01T00:00:00Z',
        priority: 'urgent',
        percentComplete: 77,
      },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();

    const patch = g.calls.find(
      (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
    );
    const body = JSON.parse(patch!.options.body);
    expect(body.bucketId).toBe(bucketId);
    expect(body.assignments).toEqual({
      [memberId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
    });
    expect(body.dueDateTime).toBe('2026-09-01T00:00:00Z');
    expect(body.startDateTime).toBe('2026-08-01T00:00:00Z');
    expect(body.priority).toBe(1);
    expect(body.percentComplete).toBe(77);

    // Reuses the task GET's own ETag for the PATCH instead of fetching the
    // identical endpoint a second time. A plain GET call carries no explicit
    // "method" in options (fakeGraph/GraphClient default it), so it must be
    // matched as (options.method ?? 'GET') === 'GET', not a literal 'GET'.
    expect(
      g.calls.filter(
        (c) => (c.options.method ?? 'GET') === 'GET' && c.endpoint === `/planner/tasks/${T1}`
      )
    ).toHaveLength(1);
    expect(patch!.options.headers['If-Match']).toBe('W/"t1"');
  });

  describe('assignments and checklist are replaced, not merged', () => {
    const groupId = '22222222-2222-2222-2222-222222222222';
    const oldMemberId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
    const newMemberId = 'cccccccc-cccc-cccc-cccc-cccccccccccc';

    // plannerTask.assignments and plannerTaskDetails.checklist are Graph OPEN
    // TYPES: PATCH merges keys already in the body, leaving anything absent
    // untouched. A body containing only the resolved/new entries would silently
    // fail to remove stale ones - these routes give the task existing assignees
    // and the details existing checklist items to prove they get explicitly
    // nulled, not left dangling.
    const routesWithExisting = {
      [`GET /planner/tasks/${T1}/details`]: {
        '@odata.etag': 'W/"d1"',
        checklist: {
          'existing-1': { title: 'old item 1', isChecked: false },
          'existing-2': { title: 'old item 2', isChecked: true },
        },
      },
      [`GET /planner/tasks/${T1}`]: {
        id: T1,
        planId: PLAN,
        '@odata.etag': 'W/"t1"',
        assignments: {
          [oldMemberId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
        },
      },
      [`GET /planner/plans/${PLAN}`]: {
        id: PLAN,
        container: { containerId: groupId, type: 'group' },
      },
      [`GET /groups/${groupId}/members`]: {
        value: [{ id: newMemberId, mail: 'new@kw-corp.com', userPrincipalName: 'new@kw-corp.com' }],
      },
      [`PATCH /planner/tasks/${T1}/details`]: {},
      [`PATCH /planner/tasks/${T1}`]: {},
    };

    it('clearing assignees (empty array) nulls every existing key, unassigning everyone', async () => {
      const g = fakeGraph(routesWithExisting);
      await updatePlannerTaskTool.execute({ taskId: T1, assignees: [] }, { graphClient: g });
      const patch = g.calls.find(
        (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
      );
      expect(JSON.parse(patch!.options.body).assignments).toEqual({ [oldMemberId]: null });
    });

    it('replacing assignees nulls the stale assignee AND adds the new one', async () => {
      const g = fakeGraph(routesWithExisting);
      await updatePlannerTaskTool.execute(
        { taskId: T1, assignees: ['new@kw-corp.com'] },
        { graphClient: g }
      );
      const patch = g.calls.find(
        (c) => c.options.method === 'PATCH' && !c.endpoint.endsWith('/details')
      );
      expect(JSON.parse(patch!.options.body).assignments).toEqual({
        [oldMemberId]: null,
        [newMemberId]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      });
    });

    it('clearing the checklist (empty array) nulls every existing item', async () => {
      const g = fakeGraph(routesWithExisting);
      await updatePlannerTaskTool.execute({ taskId: T1, checklist: [] }, { graphClient: g });
      const detailPatch = g.calls.find(
        (c) => c.options.method === 'PATCH' && c.endpoint.endsWith('/details')
      );
      expect(JSON.parse(detailPatch!.options.body).checklist).toEqual({
        'existing-1': null,
        'existing-2': null,
      });
    });

    it('replacing the checklist nulls the stale items AND adds the new one', async () => {
      const g = fakeGraph(routesWithExisting);
      await updatePlannerTaskTool.execute(
        { taskId: T1, checklist: ['new item'] },
        { graphClient: g }
      );
      const detailPatch = g.calls.find(
        (c) => c.options.method === 'PATCH' && c.endpoint.endsWith('/details')
      );
      const checklist = JSON.parse(detailPatch!.options.body).checklist as Record<string, unknown>;
      expect(checklist['existing-1']).toBeNull();
      expect(checklist['existing-2']).toBeNull();
      const newEntries = Object.entries(checklist).filter(([key]) => !key.startsWith('existing-'));
      expect(newEntries).toHaveLength(1);
      expect(newEntries[0][1]).toMatchObject({ title: 'new item', isChecked: false });
    });

    // Reuses the GET already performed to read the existing checklist keys
    // for the details PATCH's ETag, instead of fetching /details a second time.
    it('fetches the details entity exactly once when clearing the checklist', async () => {
      const g = fakeGraph(routesWithExisting);
      await updatePlannerTaskTool.execute({ taskId: T1, checklist: [] }, { graphClient: g });
      // A plain GET carries no explicit "method" in options (fakeGraph/GraphClient
      // default it), so it must be matched as (options.method ?? 'GET') === 'GET'.
      const detailsGets = g.calls.filter(
        (c) => (c.options.method ?? 'GET') === 'GET' && c.endpoint.endsWith('/details')
      );
      expect(detailsGets).toHaveLength(1);
      const detailPatch = g.calls.find(
        (c) => c.options.method === 'PATCH' && c.endpoint.endsWith('/details')
      );
      expect(detailPatch!.options.headers['If-Match']).toBe('W/"d1"');
    });
  });
});

describe('update-planner-task confirm gate', () => {
  const ORIGINAL = process.env.MS365_MCP_REQUIRE_CONFIRM;

  afterEach(() => {
    if (ORIGINAL === undefined) delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    else process.env.MS365_MCP_REQUIRE_CONFIRM = ORIGINAL;
  });

  it('gate off (default): updates without a confirm param', async () => {
    delete process.env.MS365_MCP_REQUIRE_CONFIRM;
    const g = fakeGraph(UPDATE_ROUTES);
    const result = await updatePlannerTaskTool.execute(
      { taskId: T1, title: 'Renamed' },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'PATCH')).toBe(true);
  });

  it('gate on: refuses without confirm: true and never touches Graph', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph(UPDATE_ROUTES);
    const result = await updatePlannerTaskTool.execute(
      { taskId: T1, title: 'Renamed' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/confirmation_required/);
    expect(g.calls).toHaveLength(0);
  });

  it('gate on: proceeds with confirm: true', async () => {
    process.env.MS365_MCP_REQUIRE_CONFIRM = 'true';
    const g = fakeGraph(UPDATE_ROUTES);
    const result = await updatePlannerTaskTool.execute(
      { taskId: T1, title: 'Renamed', confirm: true },
      { graphClient: g }
    );
    expect(result.isError).toBeUndefined();
    expect(g.calls.some((c) => c.options.method === 'PATCH')).toBe(true);
  });
});

describe('update-planner-task schema', () => {
  const schema = z.object(updatePlannerTaskTool.buildSchema());

  it('accepts a well-formed request', () => {
    expect(schema.safeParse({ taskId: T1, title: 'Renamed' }).success).toBe(true);
  });

  it('accepts percentComplete: 0 and priority: 0', () => {
    expect(schema.safeParse({ taskId: T1, percentComplete: 0, priority: 0 }).success).toBe(true);
  });

  it('rejects a path-traversal-style taskId', () => {
    const result = schema.safeParse({ taskId: '../../etc/passwd', title: 'Renamed' });
    expect(result.success).toBe(false);
  });

  it('rejects a taskId containing a slash', () => {
    const result = schema.safeParse({ taskId: 'abc/def', title: 'Renamed' });
    expect(result.success).toBe(false);
  });

  it('rejects a percentComplete above 100', () => {
    expect(schema.safeParse({ taskId: T1, percentComplete: 101 }).success).toBe(false);
  });

  it('rejects a percentComplete below 0', () => {
    expect(schema.safeParse({ taskId: T1, percentComplete: -1 }).success).toBe(false);
  });

  it('rejects a non-integer percentComplete', () => {
    expect(schema.safeParse({ taskId: T1, percentComplete: 3.7 }).success).toBe(false);
  });
});
