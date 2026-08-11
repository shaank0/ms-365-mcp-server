import { describe, it, expect } from 'vitest';
import { PLANNER_TOOLS } from '../index.js';
import { UTILITY_TOOLS } from '../../graph-tools.js';

describe('planner tool registration', () => {
  // Regression guard for the "endpointsData vs allEndpoints" suppression bug: assert
  // against the REAL registry graph-tools.ts builds, not against OVERRIDDEN_TOOL_NAMES
  // itself (which is derived from PLANNER_TOOLS and would trivially match any
  // implementation). See override-suppression.test.ts for the collision-suppression
  // mechanism itself.
  it('registers delete-planner-task in UTILITY_TOOLS exactly once', () => {
    const matches = UTILITY_TOOLS.filter((t) => t.name === 'delete-planner-task');
    expect(matches).toHaveLength(1);
  });

  it('gives every tool the required utility-tool shape', () => {
    for (const tool of PLANNER_TOOLS) {
      expect(tool.name).toMatch(/^[a-z-]+$/);
      expect(tool.path).toBe(`tool:${tool.name}`);
      expect(typeof tool.description).toBe('string');
      expect(tool.description.length).toBeGreaterThan(20);
      expect(typeof tool.buildSchema).toBe('function');
      expect(typeof tool.execute).toBe('function');
      expect(typeof tool.readOnlyHint).toBe('boolean');
    }
  });

  it('has no duplicate names anywhere in UTILITY_TOOLS', () => {
    const names = UTILITY_TOOLS.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });
});
