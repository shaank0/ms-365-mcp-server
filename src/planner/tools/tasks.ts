import { z } from 'zod';
import { getEtag, deleteWithEtag, withEtagRetry, type GraphLike } from '../graph.js';
import { ok, toolError, type McpResult } from '../result.js';

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
      .describe('Planner task id (28-character string, e.g. from list-plan-tasks).'),
  }),
  execute: async (params, { graphClient }) => {
    try {
      const endpoint = `/planner/tasks/${params.taskId}`;
      await withEtagRetry(
        () => getEtag(graphClient, endpoint),
        (etag) => deleteWithEtag(graphClient, endpoint, etag)
      );
      return { ...ok({ message: `Task ${params.taskId} deleted.` }) };
    } catch (err) {
      return { ...toolError(err) };
    }
  },
};
