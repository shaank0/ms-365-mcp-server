import { deletePlannerTaskTool, type PlannerTool } from './tools/tasks.js';
import { createPlannerPlanTool, listPlannerPlansTool } from './tools/plans.js';

/**
 * Tools this fork adds to the upstream ms-365-mcp-server. Any name listed
 * here is filtered out of endpoints.json at the load site in graph-tools.ts,
 * so ours is the only registration ("overlay wins").
 */
export const PLANNER_TOOLS: PlannerTool[] = [
  createPlannerPlanTool,
  listPlannerPlansTool,
  deletePlannerTaskTool,
];

export const OVERRIDDEN_TOOL_NAMES = new Set(PLANNER_TOOLS.map((t) => t.name));

export type { PlannerTool };
