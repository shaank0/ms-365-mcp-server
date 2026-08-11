import { z } from 'zod';
import { getEtag, deleteWithEtag, withEtagRetry, PLANNER_ID, type GraphLike } from '../graph.js';
import { ok, toolError, type McpResult } from '../result.js';
import { checkConfirmGate } from '../confirm-gate.js';
import { requirePlannerId } from '../validate.js';
import { CONFIRM_PARAM_DESCRIPTION } from '../../lib/param-descriptions.js';

export interface PlannerTool {
  name: string;
  method: string;
  path: string;
  description: string;
  readOnlyHint: boolean;
  openWorldHint: boolean;
  buildSchema: (ctx?: unknown) => Record<string, z.ZodTypeAny>;
  // Intersected with an index signature (upstream's UtilityTool/CallToolResult
  // shape requires one) so PlannerTool structurally satisfies UtilityTool with
  // no cast at the UTILITY_TOOLS spread site in graph-tools.ts. McpResult itself
  // (result.ts) stays untouched — execute() below produces the index signature
  // for free by spreading ok()/toolError()'s return into a fresh object literal.
  execute: (
    params: any,
    deps: { graphClient: GraphLike }
  ) => Promise<McpResult & Record<string, unknown>>;
}

export const deletePlannerTaskTool: PlannerTool = {
  name: 'delete-planner-task',
  method: 'DELETE',
  path: 'tool:delete-planner-task',
  description:
    'Delete a Microsoft Planner task. The required If-Match ETag is fetched and applied automatically, so no ETag argument is needed. This is permanent — the task and its details, checklist and comments are removed.',
  readOnlyHint: false,
  openWorldHint: true,
  buildSchema: () => ({
    taskId: z
      .string()
      .regex(
        PLANNER_ID,
        'Must be a 28-character Planner task id (letters, digits, "_" or "-" only) — e.g. from list-plan-tasks. Not a path or URL.'
      )
      .describe('Planner task id (28-character string, e.g. from list-plan-tasks).'),
    confirm: z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional(),
  }),
  execute: async (params, { graphClient }) => {
    const refusal = checkConfirmGate('delete-planner-task', params);
    if (refusal) return refusal;
    try {
      // Second line of defense: Zod's schema regex only runs on the normal MCP
      // registration path. Discovery mode's execute-tool calls execute() directly
      // with raw, unparsed client input, so re-validate here too.
      const taskId = requirePlannerId(params.taskId, 'taskId');
      const endpoint = `/planner/tasks/${taskId}`;
      await withEtagRetry(
        () => getEtag(graphClient, endpoint),
        (etag) => deleteWithEtag(graphClient, endpoint, etag)
      );
      return { ...ok({ message: `Task ${taskId} deleted.` }) };
    } catch (err) {
      return { ...toolError(err) };
    }
  },
};
