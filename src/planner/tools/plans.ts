import { z } from 'zod';
import { postJson, paginate, type GraphLike } from '../graph.js';
import { resolveGroup } from '../resolve.js';
import { ok, toolError } from '../result.js';
import type { PlannerTool } from './tasks.js';

/** Cap on groups scanned by list-planner-plans, to bound the request fan-out. */
const GROUP_SCAN_CAP = 60;

export const createPlannerPlanTool: PlannerTool = {
  name: 'create-planner-plan',
  method: 'POST',
  path: 'tool:create-planner-plan',
  description:
    'Create a Microsoft Planner plan inside a Microsoft 365 group. Pass the group by display name or id — names resolve against groups you are a member of. Optionally supply bucket names to create in order. Returns the new plan id, which every other Planner tool needs.',
  readOnlyHint: false,
  openWorldHint: true,
  buildSchema: () => ({
    group: z
      .string()
      .describe(
        'Microsoft 365 group that will own the plan: its display name (e.g. "Marketing") or its group id (GUID). You must be a member of it.'
      ),
    title: z.string().describe('Title of the new plan.'),
    buckets: z
      .array(z.string())
      .optional()
      .describe('Optional bucket (column) names to create in the plan, in order.'),
  }),
  execute: async (params, { graphClient }) => {
    try {
      const groupId = await resolveGroup(graphClient, params.group);
      const plan = await postJson<{ id: string; title: string }>(graphClient, '/planner/plans', {
        container: { url: `https://graph.microsoft.com/v1.0/groups/${groupId}` },
        title: params.title,
      });

      const created: string[] = [];
      const failed: string[] = [];
      for (const name of (params.buckets ?? []) as string[]) {
        try {
          await postJson(graphClient, '/planner/buckets', {
            name,
            planId: plan.id,
            orderHint: ' !',
          });
          created.push(name);
        } catch {
          failed.push(name);
        }
      }

      return ok({
        id: plan.id,
        title: plan.title,
        groupId,
        buckets: created,
        ...(failed.length
          ? {
              warning: `The plan was created, but these buckets were NOT created: ${failed.join(', ')}. Add them with create-planner-bucket using planId ${plan.id}.`,
            }
          : {}),
      });
    } catch (err) {
      return toolError(err);
    }
  },
};

export const listPlannerPlansTool: PlannerTool = {
  name: 'list-planner-plans',
  method: 'GET',
  path: 'tool:list-planner-plans',
  description:
    'List every Microsoft Planner plan you can reach, across all Microsoft 365 groups you belong to. Use this to discover a planId before calling list-plan-tasks, list-plan-buckets or create-planner-task. Groups whose plans cannot be read are skipped and reported.',
  readOnlyHint: true,
  openWorldHint: true,
  buildSchema: () => ({}),
  execute: async (_params, { graphClient }) => {
    try {
      const { items, truncated } = await paginate<{
        id: string;
        displayName?: string;
        groupTypes?: string[];
      }>(graphClient, '/me/memberOf?$select=id,displayName,groupTypes', GROUP_SCAN_CAP);

      const groups = items.filter((i) => (i.groupTypes ?? []).includes('Unified'));
      const plans: Array<{ id: string; title: string; group: string; groupId: string }> = [];
      const skippedGroups: string[] = [];

      for (const group of groups) {
        try {
          const page = await paginate<{ id: string; title?: string }>(
            graphClient,
            `/groups/${group.id}/planner/plans`,
            200
          );
          for (const plan of page.items) {
            plans.push({
              id: plan.id,
              title: plan.title ?? '(untitled)',
              group: group.displayName ?? '(unnamed group)',
              groupId: group.id,
            });
          }
        } catch {
          skippedGroups.push(group.displayName ?? group.id);
        }
      }

      return ok({
        plans,
        ...(skippedGroups.length ? { skippedGroups } : {}),
        ...(truncated
          ? {
              truncated: `Only the first ${GROUP_SCAN_CAP} groups were scanned; you belong to more. Plans in the remaining groups are not listed.`,
            }
          : {}),
      });
    } catch (err) {
      return toolError(err);
    }
  },
};
