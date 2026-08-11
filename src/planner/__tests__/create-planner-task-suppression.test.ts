import { describe, it, expect, vi } from 'vitest';
import { registerGraphTools, buildToolsRegistry } from '../../graph-tools.js';
import type GraphClient from '../../graph-client.js';

/**
 * override-suppression.test.ts proves the suppression MECHANISM using
 * update-planner-task as a mocked stand-in for a not-yet-written override
 * (it mocks ../index.js so it doesn't have to wait for a real collision to
 * exist). Task 7 ships the first REAL one: create-planner-task is a genuine
 * toolName in the real generated client and a genuine entry in
 * endpoints.json, and src/planner/index.js's PLANNER_TOOLS now really does
 * carry a same-named override (src/planner/tools/tasks.ts's
 * createPlannerTaskTool). This file deliberately does NOT mock ../index.js -
 * it drives the real PLANNER_TOOLS / OVERRIDDEN_TOOL_NAMES against a fake
 * declarative client shaped like the real collision, so the suppression
 * mechanism is exercised on the exact tool it was built for, not a stand-in.
 * (A single file, rather than a test nested inside override-suppression.test.ts,
 * because vi.mock/vi.unmock calls are hoisted to module scope in Vitest - they
 * can't be toggled per-test within one file without fighting that hoisting.)
 */

vi.mock('../../logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../../generated/client-beta.js', () => ({ api: { endpoints: [] } }));

vi.mock('../../generated/client.js', () => ({
  api: {
    endpoints: [
      {
        // Upstream's real declarative registration: a raw POST /planner/tasks
        // passthrough, with no description/assignees/checklist handling -
        // exactly what createPlannerTaskTool replaces.
        alias: 'create-planner-task',
        method: 'post',
        path: '/planner/tasks',
        description: 'Create planner task',
        parameters: [],
      },
      {
        alias: 'list-plan-tasks',
        method: 'get',
        path: '/planner/plans/{plannerPlan-id}/tasks',
        description: 'List plan tasks',
        parameters: [],
      },
      {
        alias: 'send-mail',
        method: 'post',
        path: '/me/sendMail',
        description: 'Send mail',
        parameters: [],
      },
    ],
  },
}));

describe('create-planner-task: a genuine, shipped collision', () => {
  it('registers exactly once, the registration is ours, and unrelated upstream tools are unaffected', () => {
    const mockServer = { tool: vi.fn(), registerTool: vi.fn() };

    registerGraphTools(mockServer as never, {} as GraphClient, false);

    const declarativeNames = mockServer.registerTool.mock.calls.map((call) => call[0]);
    const utilityCalls = mockServer.tool.mock.calls;
    const utilityNames = utilityCalls.map((call) => call[0]);

    // The declarative (generated-client) registration must NOT fire for
    // create-planner-task - the overlay is the sole registration.
    expect(declarativeNames).not.toContain('create-planner-task');
    expect(utilityNames.filter((n) => n === 'create-planner-task')).toHaveLength(1);

    // Unrelated declarative endpoints are unaffected by the override.
    expect(declarativeNames).toContain('list-plan-tasks');
    expect(declarativeNames).toContain('send-mail');

    // Prove the surviving registration is OURS, not a degraded stand-in: its
    // schema exposes description/assignees/checklist, which upstream's
    // declarative create-planner-task (a raw body passthrough) does not.
    const ourRegistration = utilityCalls.find((call) => call[0] === 'create-planner-task');
    expect(ourRegistration).toBeDefined();
    const schemaShape = ourRegistration![2] as Record<string, unknown>;
    expect(Object.keys(schemaShape)).toEqual(
      expect.arrayContaining(['description', 'assignees', 'checklist'])
    );
  });

  // buildToolsRegistry backs discovery mode's search-tools/get-tool-schema/execute-tool
  // triad and reads the SAME allEndpoints array as registerGraphTools above, but is a
  // separate function - pin it directly, same as override-suppression.test.ts does for
  // its update-planner-task stand-in.
  it('excludes the declarative entry from buildToolsRegistry (discovery mode) too', () => {
    const registry = buildToolsRegistry(false, false);
    expect(registry.has('create-planner-task')).toBe(false);
    expect(registry.has('list-plan-tasks')).toBe(true);
    expect(registry.has('send-mail')).toBe(true);
  });
});
