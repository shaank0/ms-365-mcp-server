import { describe, it, expect } from 'vitest';
import { createPlannerPlanTool, listPlannerPlansTool } from '../tools/plans.js';
import { fakeGraph } from './fake-graph.js';

const MEMBER_OF = {
  'GET /me/memberOf': {
    value: [
      {
        id: '11111111-1111-1111-1111-111111111111',
        displayName: 'Marketing',
        groupTypes: ['Unified'],
      },
    ],
  },
};

describe('create-planner-plan', () => {
  it('creates a plan in the resolved group', async () => {
    const g = fakeGraph({
      ...MEMBER_OF,
      'POST /planner/plans': { id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', title: 'Q3 Rollout' },
    });
    const result = await createPlannerPlanTool.execute(
      { group: 'Marketing', title: 'Q3 Rollout' },
      { graphClient: g }
    );
    const payload = JSON.parse(result.content[0].text);
    expect(result.isError).toBeUndefined();
    expect(payload.id).toBe('PPPPPPPPPPPPPPPPPPPPPPPPPPPP');
    const post = g.calls.find((c) => c.options.method === 'POST');
    expect(JSON.parse(post!.options.body).container.url).toBe(
      'https://graph.microsoft.com/v1.0/groups/11111111-1111-1111-1111-111111111111'
    );
  });

  it('creates the requested buckets in order', async () => {
    const g = fakeGraph({
      ...MEMBER_OF,
      'POST /planner/plans': { id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', title: 'Q3' },
      'POST /planner/buckets': {},
    });
    await createPlannerPlanTool.execute(
      { group: 'Marketing', title: 'Q3', buckets: ['To Do', 'Doing'] },
      { graphClient: g }
    );
    const names = g.calls
      .filter((c) => c.endpoint === '/planner/buckets')
      .map((c) => JSON.parse(c.options.body).name);
    expect(names).toEqual(['To Do', 'Doing']);
  });

  it('reports partial failure when a bucket cannot be created', async () => {
    const g = fakeGraph({
      ...MEMBER_OF,
      'POST /planner/plans': { id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', title: 'Q3' },
      'POST /planner/buckets': (opts: any) => {
        if (JSON.parse(opts.body).name === 'Doing') throw new Error('500 Server Error');
        return {};
      },
    });
    const result = await createPlannerPlanTool.execute(
      { group: 'Marketing', title: 'Q3', buckets: ['To Do', 'Doing'] },
      { graphClient: g }
    );
    const payload = JSON.parse(result.content[0].text);
    expect(payload.buckets).toEqual(['To Do']);
    expect(payload.warning).toMatch(/Doing/);
  });

  it('surfaces an unresolvable group without creating anything', async () => {
    const g = fakeGraph(MEMBER_OF);
    const result = await createPlannerPlanTool.execute(
      { group: 'Finance', title: 'Q3' },
      { graphClient: g }
    );
    expect(result.isError).toBe(true);
    expect(g.calls.some((c) => c.options.method === 'POST')).toBe(false);
  });
});

describe('list-planner-plans', () => {
  it('returns plans across the caller groups with the group name attached', async () => {
    const g = fakeGraph({
      ...MEMBER_OF,
      'GET /groups/11111111-1111-1111-1111-111111111111/planner/plans': {
        value: [{ id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', title: 'Q3 Rollout' }],
      },
    });
    const result = await listPlannerPlansTool.execute({}, { graphClient: g });
    const payload = JSON.parse(result.content[0].text);
    expect(payload.plans).toEqual([
      {
        id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP',
        title: 'Q3 Rollout',
        group: 'Marketing',
        groupId: '11111111-1111-1111-1111-111111111111',
      },
    ]);
  });

  it('skips groups whose plans cannot be read', async () => {
    const g = fakeGraph({
      'GET /me/memberOf': {
        value: [
          {
            id: '11111111-1111-1111-1111-111111111111',
            displayName: 'Marketing',
            groupTypes: ['Unified'],
          },
          {
            id: '99999999-9999-9999-9999-999999999999',
            displayName: 'Locked',
            groupTypes: ['Unified'],
          },
        ],
      },
      'GET /groups/11111111-1111-1111-1111-111111111111/planner/plans': {
        value: [{ id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', title: 'Q3' }],
      },
      'GET /groups/99999999-9999-9999-9999-999999999999/planner/plans': () => {
        throw new Error('Microsoft Graph API error: 403 Forbidden - nope');
      },
    });
    const result = await listPlannerPlansTool.execute({}, { graphClient: g });
    const payload = JSON.parse(result.content[0].text);
    expect(payload.plans.length).toBe(1);
    expect(payload.skippedGroups).toEqual(['Locked']);
  });
});
