import { describe, it, expect, vi } from 'vitest';
import { registerGraphTools, buildToolsRegistry } from '../../graph-tools.js';
import type GraphClient from '../../graph-client.js';

/**
 * Mandatory regression test for the "endpointsData vs allEndpoints" suppression bug
 * (fix round 3): the old fix filtered OVERRIDDEN_TOOL_NAMES out of endpointsData, a
 * metadata LOOKUP TABLE keyed by toolName, not the registration source. Registration
 * iterates allEndpoints (built from the generated client) and merely does a
 * config-lookup against endpointsData, so removing an entry from endpointsData left
 * the declarative tool registered (with a degraded, config-less registration) while
 * the overlay's own registerTool call silently lost (McpServer throws "already
 * registered", swallowed by the try/catch around the utility-tool loop).
 *
 * This test originally used update-planner-task as a mocked stand-in for a
 * not-yet-written override, since at the time it was a real toolName in the
 * generated client/endpoints.json but had no real PLANNER_TOOLS entry yet.
 * Task 8 shipped that real override (updatePlannerTaskTool), so reusing the
 * same name here would now just restate production behavior that
 * update-planner-task-suppression.test.ts already covers end-to-end against
 * the REAL PLANNER_TOOLS / OVERRIDDEN_TOOL_NAMES. To keep this file doing its
 * original job — proving the generic suppression MECHANISM in isolation,
 * independent of any specific collision — it now uses a synthetic tool name
 * that has no real counterpart in the generated client or endpoints.json, and
 * still mocks ./planner/index.js — the one seam graph-tools.ts imports
 * PLANNER_TOOLS / OVERRIDDEN_TOOL_NAMES through — rather than driving the real
 * module. This is deliberately NOT a test of OVERRIDDEN_TOOL_NAMES in
 * isolation: it drives the real registerGraphTools() and asserts on what
 * actually got registered on the mock MCP server.
 */

vi.mock('../../logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../../generated/client-beta.js', () => ({ api: { endpoints: [] } }));

vi.mock('../../generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'stand-in-override-tool',
        method: 'patch',
        path: '/fake/stand-in/{id}',
        description: 'Fake declarative endpoint for suppression-mechanism testing.',
        parameters: [],
      },
      {
        alias: 'list-mail-messages',
        method: 'get',
        path: '/me/messages',
        description: 'List mail messages',
        parameters: [],
      },
    ],
  },
}));

// vi.mock factories are hoisted above module-level const declarations, so the
// fake overlay tool has to be defined inside vi.hoisted() to be visible here.
const { fakeOverlayTool } = vi.hoisted(() => ({
  fakeOverlayTool: {
    name: 'stand-in-override-tool',
    method: 'PATCH',
    path: 'tool:stand-in-override-tool',
    description: 'Synthetic overlay tool, purely for suppression-mechanism testing.',
    readOnlyHint: false,
    openWorldHint: true,
    buildSchema: () => ({}),
    execute: async () => ({ content: [{ type: 'text' as const, text: '{}' }] }),
  },
}));

vi.mock('../index.js', () => ({
  PLANNER_TOOLS: [fakeOverlayTool],
  OVERRIDDEN_TOOL_NAMES: new Set(['stand-in-override-tool']),
}));

describe('overlay suppression of a colliding declarative endpoint', () => {
  it('registers ONLY the overlay for a name in OVERRIDDEN_TOOL_NAMES, not the declarative endpoint', () => {
    const mockServer = { tool: vi.fn(), registerTool: vi.fn() };

    registerGraphTools(mockServer as never, {} as GraphClient, false);

    const declarativeNames = mockServer.registerTool.mock.calls.map((call) => call[0]);
    const utilityNames = mockServer.tool.mock.calls.map((call) => call[0]);

    // The declarative (generated-client) registration must NOT fire for the
    // overridden name - this is the assertion that fails against the old
    // endpointsData-only filter, because allEndpoints was never filtered there.
    expect(declarativeNames).not.toContain('stand-in-override-tool');

    // The overlay must be the sole registration for that name.
    expect(utilityNames.filter((n) => n === 'stand-in-override-tool')).toHaveLength(1);

    // An unrelated declarative endpoint is unaffected by the override.
    expect(declarativeNames).toContain('list-mail-messages');
  });

  // buildToolsRegistry backs discovery mode's search-tools/get-tool-schema/execute-tool
  // triad and reads the SAME allEndpoints array as registerGraphTools above, but is a
  // separate function - pin it directly so a future refactor that splits the two
  // registration paths can't silently regress discovery mode's suppression.
  it('excludes the overridden name from buildToolsRegistry (discovery mode) too', () => {
    const registry = buildToolsRegistry(false, false);
    expect(registry.has('stand-in-override-tool')).toBe(false);
    expect(registry.has('list-mail-messages')).toBe(true);
  });
});
