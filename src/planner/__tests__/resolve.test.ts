import { describe, it, expect } from 'vitest';
import { resolveGroup, resolveBucket, resolveAssignees, PlannerResolveError } from '../resolve.js';
import { fakeGraph } from './fake-graph.js';

const GROUPS = {
  'GET /me/memberOf': {
    value: [
      {
        id: '11111111-1111-1111-1111-111111111111',
        displayName: 'Marketing',
        groupTypes: ['Unified'],
      },
      { id: '22222222-2222-2222-2222-222222222222', displayName: 'Sales', groupTypes: ['Unified'] },
      { id: '33333333-3333-3333-3333-333333333333', displayName: 'Sales', groupTypes: ['Unified'] },
      { id: '44444444-4444-4444-4444-444444444444', displayName: 'Global Admins', groupTypes: [] },
    ],
  },
};

describe('resolveGroup', () => {
  it('passes a GUID straight through without calling Graph', async () => {
    const g = fakeGraph({});
    expect(await resolveGroup(g, '11111111-1111-1111-1111-111111111111')).toBe(
      '11111111-1111-1111-1111-111111111111'
    );
    expect(g.calls.length).toBe(0);
  });

  it('resolves a unique display name case-insensitively', async () => {
    const g = fakeGraph(GROUPS);
    expect(await resolveGroup(g, '  marketing ')).toBe('11111111-1111-1111-1111-111111111111');
  });

  it('lists candidates when the name is ambiguous', async () => {
    const g = fakeGraph(GROUPS);
    await expect(resolveGroup(g, 'Sales')).rejects.toThrow(/22222222.*33333333/s);
  });

  it('errors when no group matches', async () => {
    const g = fakeGraph(GROUPS);
    await expect(resolveGroup(g, 'Finance')).rejects.toThrow(PlannerResolveError);
  });

  it('ignores non-Unified directory objects', async () => {
    const g = fakeGraph(GROUPS);
    await expect(resolveGroup(g, 'Global Admins')).rejects.toThrow(/not a member of any group/i);
  });
});

const BUCKETS = {
  'GET /planner/plans/PPPPPPPPPPPPPPPPPPPPPPPPPPPP/buckets': {
    value: [
      { id: 'BBBBBBBBBBBBBBBBBBBBBBBBBBBB', name: 'To Do' },
      { id: 'CCCCCCCCCCCCCCCCCCCCCCCCCCCC', name: 'Doing' },
    ],
  },
};

describe('resolveBucket', () => {
  it('passes a 28-char Planner id through', async () => {
    const g = fakeGraph({});
    expect(
      await resolveBucket(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', 'BBBBBBBBBBBBBBBBBBBBBBBBBBBB')
    ).toBe('BBBBBBBBBBBBBBBBBBBBBBBBBBBB');
    expect(g.calls.length).toBe(0);
  });

  it('resolves a bucket name case-insensitively', async () => {
    const g = fakeGraph(BUCKETS);
    expect(await resolveBucket(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', 'to do')).toBe(
      'BBBBBBBBBBBBBBBBBBBBBBBBBBBB'
    );
  });

  it('lists available buckets on a miss and never creates one', async () => {
    const g = fakeGraph(BUCKETS);
    await expect(resolveBucket(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', 'Backlog')).rejects.toThrow(
      /To Do, Doing/
    );
    expect(g.calls.every((c) => (c.options.method ?? 'GET') === 'GET')).toBe(true);
  });
});

const PLAN_AND_MEMBERS = {
  'GET /planner/plans/PPPPPPPPPPPPPPPPPPPPPPPPPPPP': {
    id: 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP',
    container: { containerId: '11111111-1111-1111-1111-111111111111', type: 'group' },
  },
  'GET /groups/11111111-1111-1111-1111-111111111111/members': {
    value: [
      {
        id: 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa',
        mail: 'will@kw-corp.com',
        userPrincipalName: 'will@kw-corp.com',
      },
      {
        id: 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb',
        mail: null,
        userPrincipalName: 'josh@kw-corp.com',
      },
    ],
  },
};

describe('resolveAssignees', () => {
  it('maps an email to the assignment shape', async () => {
    const g = fakeGraph(PLAN_AND_MEMBERS);
    const result = await resolveAssignees(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', ['WILL@kw-corp.com']);
    expect(result).toEqual({
      'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa': {
        '@odata.type': '#microsoft.graph.plannerAssignment',
        orderHint: ' !',
      },
    });
  });

  it('matches on userPrincipalName when mail is null', async () => {
    const g = fakeGraph(PLAN_AND_MEMBERS);
    const result = await resolveAssignees(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', ['josh@kw-corp.com']);
    expect(Object.keys(result)).toEqual(['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb']);
  });

  it('passes a GUID through without a member lookup match', async () => {
    const g = fakeGraph(PLAN_AND_MEMBERS);
    const result = await resolveAssignees(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', [
      'cccccccc-cccc-cccc-cccc-cccccccccccc',
    ]);
    expect(Object.keys(result)).toEqual(['cccccccc-cccc-cccc-cccc-cccccccccccc']);
  });

  it('lists valid members when someone is not in the group', async () => {
    const g = fakeGraph(PLAN_AND_MEMBERS);
    await expect(
      resolveAssignees(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', ['stranger@example.com'])
    ).rejects.toThrow(/will@kw-corp.com/);
  });

  it('errors when the plan has no group container', async () => {
    const g = fakeGraph({
      'GET /planner/plans/PPPPPPPPPPPPPPPPPPPPPPPPPPPP': { id: 'P', container: {} },
    });
    await expect(
      resolveAssignees(g, 'PPPPPPPPPPPPPPPPPPPPPPPPPPPP', ['will@kw-corp.com'])
    ).rejects.toThrow(/no group container/i);
  });
});
