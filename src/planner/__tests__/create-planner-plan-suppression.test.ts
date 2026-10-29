import { describe, it, expect, vi } from 'vitest';
import { registerGraphTools, buildToolsRegistry } from '../../graph-tools.js';
import type GraphClient from '../../graph-client.js';

/**
 * Upstream added its own declarative create-planner-plan (a raw POST /planner/plans
 * passthrough, endpoints.json + generated client) after this fork shipped its
 * composite create-planner-plan. That turned a fork-only name into a third genuine
 * collision, alongside create-planner-task and update-planner-task. Unlike those two
 * files, this one does NOT mock the generated clients: it drives the REAL merged
 * client and endpoints.json, so it fails if a future upstream sync ever lets the
 * declarative entry win (or register twice) for this name.
 */

vi.mock('../../logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

describe('create-planner-plan: an upstream collision the overlay must win', () => {
  it('registers exactly once, and the registration is ours', () => {
    const mockServer = { tool: vi.fn(), registerTool: vi.fn() };

    registerGraphTools(mockServer as never, {} as GraphClient, { readOnly: false });

    const declarativeNames = mockServer.registerTool.mock.calls.map((call) => call[0]);
    const utilityCalls = mockServer.tool.mock.calls;

    expect(declarativeNames).not.toContain('create-planner-plan');
    const ours = utilityCalls.filter((call) => call[0] === 'create-planner-plan');
    expect(ours).toHaveLength(1);

    // Upstream's declarative tool takes a raw `body`; ours resolves a group by name
    // and can create buckets in the same call.
    const schemaShape = ours[0][2] as Record<string, unknown>;
    expect(Object.keys(schemaShape)).toEqual(
      expect.arrayContaining(['group', 'title', 'buckets', 'confirm'])
    );
  });

  it('excludes the declarative entry from buildToolsRegistry (discovery mode) too', () => {
    const registry = buildToolsRegistry(false, false);
    expect(registry.has('create-planner-plan')).toBe(false);
    // An upstream sibling Planner tool is unaffected.
    expect(registry.has('get-planner-plan')).toBe(true);
  });
});
