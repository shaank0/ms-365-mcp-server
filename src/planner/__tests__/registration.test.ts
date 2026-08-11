import { describe, it, expect } from 'vitest';
import { PLANNER_TOOLS, OVERRIDDEN_TOOL_NAMES } from '../index.js';

describe('planner tool registration', () => {
  it('exposes every tool name in the override set', () => {
    expect(OVERRIDDEN_TOOL_NAMES).toEqual(new Set(PLANNER_TOOLS.map((t) => t.name)));
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

  it('has no duplicate names', () => {
    const names = PLANNER_TOOLS.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });
});
