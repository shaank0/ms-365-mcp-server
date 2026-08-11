import { z } from 'zod';
import {
  getEtag,
  getJson,
  deleteWithEtag,
  patchWithEtag,
  postJson,
  withEtagRetry,
  PLANNER_ID,
  type GraphLike,
} from '../graph.js';
import { resolveBucket, resolveAssignees } from '../resolve.js';
import { toPriority, toChecklist, statusToPercent } from '../convert.js';
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

/**
 * Write description/checklist onto a task's details entity. Returns an error
 * message instead of throwing, because a details failure must not lose an
 * already-created task — the caller reports it as a warning alongside the
 * task id rather than losing the write silently.
 */
async function applyDetails(
  graphClient: GraphLike,
  taskId: string,
  body: Record<string, unknown>
): Promise<string | null> {
  const endpoint = `/planner/tasks/${taskId}/details`;
  try {
    await withEtagRetry(
      () => getEtag(graphClient, endpoint),
      (etag) => patchWithEtag(graphClient, endpoint, body, etag)
    );
    return null;
  } catch (err) {
    return err instanceof Error ? err.message : String(err);
  }
}

export const createPlannerTaskTool: PlannerTool = {
  name: 'create-planner-task',
  method: 'POST',
  path: 'tool:create-planner-task',
  description:
    'Create a Microsoft Planner task in one call, including its description, checklist and assignees. The bucket may be given by name, assignees by email, and priority as urgent/important/medium/low. Planner stores the description on a separate entity, which this handles internally along with the required ETags. This overrides the upstream declarative create-planner-task, which only takes a raw body and requires a separate call for the description.',
  readOnlyHint: false,
  openWorldHint: true,
  buildSchema: () => ({
    planId: z
      .string()
      .regex(
        PLANNER_ID,
        'Must be a 28-character Planner plan id (letters, digits, "_" or "-" only) — e.g. from list-planner-plans. Not a path or URL.'
      )
      .describe('Plan id (28-character string). Use list-planner-plans to find one.'),
    title: z.string().describe('Task title.'),
    bucket: z
      .string()
      .optional()
      .describe('Bucket name (e.g. "To Do") or bucket id. Must already exist in the plan.'),
    description: z.string().optional().describe('Long-form task description.'),
    assignees: z
      .array(z.string())
      .optional()
      .describe(
        "Email addresses (or user ids) to assign. They must be members of the plan's group."
      ),
    dueDateTime: z
      .string()
      .optional()
      .describe('Due date, ISO 8601 UTC, e.g. 2026-09-01T00:00:00Z.'),
    startDateTime: z.string().optional().describe('Start date, ISO 8601 UTC.'),
    priority: z
      .union([z.string(), z.number()])
      .optional()
      .describe('urgent, important, medium, low — or an integer 0-10.'),
    checklist: z
      .array(
        z.union([z.string(), z.object({ title: z.string(), checked: z.boolean().optional() })])
      )
      .optional()
      .describe('Checklist items, as plain strings or {title, checked} objects.'),
    confirm: z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional(),
  }),
  execute: async (params, { graphClient }) => {
    const refusal = checkConfirmGate('create-planner-task', params);
    if (refusal) return refusal;
    try {
      // Second line of defense: Zod's schema regex only runs on the normal MCP
      // registration path. Discovery mode's execute-tool calls execute() directly
      // with raw, unparsed client input, so re-validate here too.
      const planId = requirePlannerId(params.planId, 'planId');

      const body: Record<string, unknown> = { planId, title: params.title };
      if (params.bucket) {
        body.bucketId = await resolveBucket(graphClient, planId, params.bucket);
      }
      if (params.assignees?.length) {
        body.assignments = await resolveAssignees(graphClient, planId, params.assignees);
      }
      if (params.dueDateTime !== undefined) body.dueDateTime = params.dueDateTime;
      if (params.startDateTime !== undefined) body.startDateTime = params.startDateTime;
      if (params.priority !== undefined) body.priority = toPriority(params.priority);

      const task = await postJson<{ id: string; title: string }>(
        graphClient,
        '/planner/tasks',
        body
      );

      // The description and checklist live on a separate "details" entity with
      // its own ETag, so only round-trip to it when one of those was actually
      // supplied — the common case (no description, no checklist) stays a
      // single POST.
      const detailsBody: Record<string, unknown> = {};
      if (params.description) detailsBody.description = params.description;
      if (params.checklist?.length) detailsBody.checklist = toChecklist(params.checklist);

      if (Object.keys(detailsBody).length > 0) {
        const failure = await applyDetails(graphClient, task.id, detailsBody);
        if (failure) {
          // The task itself was created successfully — report that as success,
          // but never silently drop the description/checklist. Name both the
          // unapplied fields and the task id so the caller can recover.
          return {
            ...ok({
              id: task.id,
              title: task.title,
              planId,
              warning: `The task was created, but its ${Object.keys(detailsBody).join(' and ')} could NOT be saved: ${failure}. Retry with update-planner-task, using taskId ${task.id}.`,
            }),
          };
        }
      }

      return { ...ok({ id: task.id, title: task.title, planId }) };
    } catch (err) {
      return { ...toolError(err) };
    }
  },
};

export const updatePlannerTaskTool: PlannerTool = {
  name: 'update-planner-task',
  method: 'PATCH',
  path: 'tool:update-planner-task',
  description:
    'Update a Microsoft Planner task. Any subset of fields may be given; the required ETags are fetched and applied automatically, so a 412 cannot happen. Use status "complete" to finish a task. The description and checklist live on a separate entity, which this handles internally along with its own ETag. This overrides the upstream declarative update-planner-task, which only PATCHes the raw task entity, requires a manually supplied If-Match ETag, and cannot touch the description or checklist at all.',
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
    title: z.string().optional().describe('New task title.'),
    bucket: z
      .string()
      .optional()
      .describe(
        'Bucket name (e.g. "To Do") or bucket id to move the task into. Must already exist in the plan.'
      ),
    description: z.string().optional().describe('Replacement description.'),
    assignees: z
      .array(z.string())
      .optional()
      .describe(
        "Replacement assignee list, by email or user id. Members of the plan's group only. Pass an empty array to unassign everyone."
      ),
    dueDateTime: z
      .string()
      .optional()
      .describe('Due date, ISO 8601 UTC, e.g. 2026-09-01T00:00:00Z.'),
    startDateTime: z.string().optional().describe('Start date, ISO 8601 UTC.'),
    priority: z
      .union([z.string(), z.number()])
      .optional()
      .describe('urgent, important, medium, low — or an integer 0-10.'),
    status: z
      .string()
      .optional()
      .describe('complete, in-progress or not-started. Sets percentComplete to 100, 50 or 0.'),
    percentComplete: z
      .number()
      .optional()
      .describe('Explicit completion percentage, 0-100. Overrides status if both are given.'),
    checklist: z
      .array(
        z.union([z.string(), z.object({ title: z.string(), checked: z.boolean().optional() })])
      )
      .optional()
      .describe(
        'Replacement checklist items, as plain strings or {title, checked} objects. Pass an empty array to clear the checklist.'
      ),
    confirm: z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional(),
  }),
  execute: async (params, { graphClient }) => {
    const refusal = checkConfirmGate('update-planner-task', params);
    if (refusal) return refusal;
    try {
      // Second line of defense: Zod's schema regex only runs on the normal MCP
      // registration path. Discovery mode's execute-tool calls execute() directly
      // with raw, unparsed client input, so re-validate here too.
      const taskId = requirePlannerId(params.taskId, 'taskId');
      const taskEndpoint = `/planner/tasks/${taskId}`;
      const taskBody: Record<string, unknown> = {};

      // Every optional field below is tested with !== undefined, never
      // truthiness, so that legitimate falsy values - percentComplete: 0,
      // priority: 0, an explicit empty assignees/checklist array - are never
      // silently dropped in favor of a stale value already on the task.
      if (params.title !== undefined) taskBody.title = params.title;
      if (params.dueDateTime !== undefined) taskBody.dueDateTime = params.dueDateTime;
      if (params.startDateTime !== undefined) taskBody.startDateTime = params.startDateTime;
      if (params.priority !== undefined) taskBody.priority = toPriority(params.priority);
      if (params.status !== undefined) taskBody.percentComplete = statusToPercent(params.status);
      // percentComplete is checked after status so an explicit value always wins.
      if (params.percentComplete !== undefined) taskBody.percentComplete = params.percentComplete;

      // Bucket and assignee resolution both need the plan, which lives on the task.
      if (params.bucket !== undefined || params.assignees !== undefined) {
        const task = await getJson<{ planId: string }>(graphClient, taskEndpoint);
        if (params.bucket !== undefined) {
          taskBody.bucketId = await resolveBucket(graphClient, task.planId, params.bucket);
        }
        if (params.assignees !== undefined) {
          taskBody.assignments = await resolveAssignees(graphClient, task.planId, params.assignees);
        }
      }

      const detailsBody: Record<string, unknown> = {};
      if (params.description !== undefined) detailsBody.description = params.description;
      if (params.checklist !== undefined) detailsBody.checklist = toChecklist(params.checklist);

      if (Object.keys(taskBody).length === 0 && Object.keys(detailsBody).length === 0) {
        throw new Error(
          'Provide at least one field to update (title, bucket, description, assignees, dates, priority, status, percentComplete or checklist).'
        );
      }

      if (Object.keys(taskBody).length > 0) {
        await withEtagRetry(
          () => getEtag(graphClient, taskEndpoint),
          (etag) => patchWithEtag(graphClient, taskEndpoint, taskBody, etag)
        );
      }

      if (Object.keys(detailsBody).length > 0) {
        const failure = await applyDetails(graphClient, taskId, detailsBody);
        if (failure) {
          // The task fields (if any) were already saved successfully - report
          // that as success, but never silently drop the description/checklist.
          // Name both the unapplied fields and the task id so the caller can
          // recover with a straight retry, no manual ETag juggling required.
          return {
            ...ok({
              id: taskId,
              updated: Object.keys(taskBody),
              warning: `The task fields were saved, but its ${Object.keys(detailsBody).join(' and ')} could NOT be saved: ${failure}. Retry with update-planner-task, using taskId ${taskId}.`,
            }),
          };
        }
      }

      return {
        ...ok({ id: taskId, updated: [...Object.keys(taskBody), ...Object.keys(detailsBody)] }),
      };
    } catch (err) {
      return { ...toolError(err) };
    }
  },
};
