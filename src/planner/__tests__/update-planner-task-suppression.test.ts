import { describe, it, expect, vi } from 'vitest';
import { registerGraphTools, buildToolsRegistry } from '../../graph-tools.js';
import type GraphClient from '../../graph-client.js';

/**
 * Sibling of create-planner-task-suppression.test.ts (Task 7), for the second
 * genuine collision this fork ships: update-planner-task is a real toolName in
 * the real generated client and a real entry in endpoints.json, and
 * src/planner/index.js's PLANNER_TOOLS now really does carry a same-named
 * override (src/planner/tools/tasks.ts's updatePlannerTaskTool). Like its
 * sibling, this file deliberately does NOT mock ../index.js - it drives the
 * real PLANNER_TOOLS / OVERRIDDEN_TOOL_NAMES against a fake declarative client
 * shaped like the real collision, so the suppression mechanism is exercised on
 * the exact tool it was built for, not a stand-in.
 *
 * override-suppression.test.ts now covers the MECHANISM in the abstract with a
 * synthetic tool name (update-planner-task stopped being a safe stand-in the
 * moment this collision became real); this file is the one that proves the
 * real update-planner-task collision behaves correctly.
 */

vi.mock('../../logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../../generated/client-beta.js', () => ({ api: { endpoints: [] } }));

vi.mock('../../generated/client.js', () => ({
  api: {
    endpoints: [
      {
        // Upstream's real declarative registration: a raw PATCH
        // /planner/tasks/{id} passthrough requiring a manually supplied
        // If-Match ETag, and no access to the description/checklist entity -
        // exactly what updatePlannerTaskTool replaces.
        alias: 'update-planner-task',
        method: 'patch',
        path: '/planner/tasks/{plannerTask-id}',
        description: 'Update planner task',
        parameters: [],
      },
      {
        alias: 'get-planner-task',
        method: 'get',
        path: '/planner/tasks/{plannerTask-id}',
        description: 'Get planner task',
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

describe('update-planner-task: a genuine, shipped collision', () => {
  it('registers exactly once, the registration is ours, and unrelated upstream tools are unaffected', () => {
    const mockServer = { tool: vi.fn(), registerTool: vi.fn() };

    registerGraphTools(mockServer as never, {} as GraphClient, false);

    const declarativeNames = mockServer.registerTool.mock.calls.map((call) => call[0]);
    const utilityCalls = mockServer.tool.mock.calls;
    const utilityNames = utilityCalls.map((call) => call[0]);

    // The declarative (generated-client) registration must NOT fire for
    // update-planner-task - the overlay is the sole registration.
    expect(declarativeNames).not.toContain('update-planner-task');
    expect(utilityNames.filter((n) => n === 'update-planner-task')).toHaveLength(1);

    // Unrelated declarative endpoints are unaffected by the override.
    expect(declarativeNames).toContain('get-planner-task');
    expect(declarativeNames).toContain('send-mail');

    // Prove the surviving registration is OURS, not a degraded stand-in: its
    // schema exposes description/checklist/status, which upstream's
    // declarative update-planner-task (a raw task-entity-only PATCH) does not.
    const ourRegistration = utilityCalls.find((call) => call[0] === 'update-planner-task');
    expect(ourRegistration).toBeDefined();
    const schemaShape = ourRegistration![2] as Record<string, unknown>;
    expect(Object.keys(schemaShape)).toEqual(
      expect.arrayContaining(['description', 'checklist', 'status'])
    );
  });

  // buildToolsRegistry backs discovery mode's search-tools/get-tool-schema/execute-tool
  // triad and reads the SAME allEndpoints array as registerGraphTools above, but is a
  // separate function - pin it directly, same as the other suppression tests do.
  it('excludes the declarative entry from buildToolsRegistry (discovery mode) too', () => {
    const registry = buildToolsRegistry(false, false);
    expect(registry.has('update-planner-task')).toBe(false);
    expect(registry.has('get-planner-task')).toBe(true);
    expect(registry.has('send-mail')).toBe(true);
  });
});
