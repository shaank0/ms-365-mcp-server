import { z } from 'zod';
import { postJson, paginate } from '../graph.js';
import { resolveGroup } from '../resolve.js';
import { ok, toolError } from '../result.js';
import { checkConfirmGate } from '../confirm-gate.js';
import type { PlannerTool } from './tasks.js';
import { CONFIRM_PARAM_DESCRIPTION } from '../../lib/param-descriptions.js';

/**
 * Cap on directory memberships read from /me/memberOf. This is cheap — one
 * paginated read — so it can be generous, matching resolveGroup's own cap.
 * NOT the cap on how many groups' plans get fetched: /me/memberOf returns
 * every directory object (security groups, directory roles, administrative
 * units), most of which are never Planner-capable, so capping HERE first
 * would let unrelated memberships crowd real Microsoft 365 groups out of the
 * scan. See GROUP_FANOUT_CAP below for the cap that actually matters.
 */
const MEMBERSHIP_SCAN_CAP = 500;

/**
 * Cap on Unified (Microsoft 365) groups whose plans are actually fetched —
 * applied AFTER filtering to Unified groups, not to the raw membership list.
 * This is the expensive part (one Graph request per group), so it is the cap
 * that bounds real cost.
 */
const GROUP_FANOUT_CAP = 60;

/** Cap on plans read per group. */
const PLANS_PER_GROUP_CAP = 200;

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
    confirm: z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional(),
  }),
  execute: async (params, { graphClient }) => {
    const refusal = checkConfirmGate('create-planner-plan', params);
    if (refusal) return refusal;
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

      return {
        ...ok({
          id: plan.id,
          title: plan.title,
          groupId,
          buckets: created,
          ...(failed.length
            ? {
                warning: `The plan was created, but these buckets were NOT created: ${failed.join(', ')}. Add them with create-planner-bucket using planId ${plan.id}.`,
              }
            : {}),
        }),
      };
    } catch (err) {
      return { ...toolError(err) };
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
      const { items, truncated: membershipTruncated } = await paginate<{
        id: string;
        displayName?: string;
        groupTypes?: string[];
      }>(graphClient, '/me/memberOf?$select=id,displayName,groupTypes', MEMBERSHIP_SCAN_CAP);

      const unifiedGroups = items.filter((i) => (i.groupTypes ?? []).includes('Unified'));
      const groups = unifiedGroups.slice(0, GROUP_FANOUT_CAP);
      const groupsTruncated = unifiedGroups.length > GROUP_FANOUT_CAP;

      const plans: Array<{ id: string; title: string; group: string; groupId: string }> = [];
      const skippedGroups: string[] = [];
      const groupsWithMorePlans: string[] = [];

      for (const group of groups) {
        try {
          const page = await paginate<{ id: string; title?: string }>(
            graphClient,
            `/groups/${group.id}/planner/plans`,
            PLANS_PER_GROUP_CAP
          );
          for (const plan of page.items) {
            plans.push({
              id: plan.id,
              title: plan.title ?? '(untitled)',
              group: group.displayName ?? '(unnamed group)',
              groupId: group.id,
            });
          }
          if (page.truncated) {
            groupsWithMorePlans.push(group.displayName ?? group.id);
          }
        } catch {
          skippedGroups.push(group.displayName ?? group.id);
        }
      }

      return {
        ...ok({
          plans,
          ...(skippedGroups.length ? { skippedGroups } : {}),
          ...(groupsWithMorePlans.length
            ? {
                truncatedGroupPlans: `These groups have more than ${PLANS_PER_GROUP_CAP} plans; only the first ${PLANS_PER_GROUP_CAP} are listed for each: ${groupsWithMorePlans.join(', ')}.`,
              }
            : {}),
          ...(groupsTruncated
            ? {
                truncated: `Only the first ${GROUP_FANOUT_CAP} of your ${unifiedGroups.length} Microsoft 365 groups were scanned; plans in the remaining groups are not listed.`,
              }
            : {}),
          ...(membershipTruncated
            ? {
                truncatedMemberships: `Only the first ${MEMBERSHIP_SCAN_CAP} directory memberships were checked; you may belong to Microsoft 365 groups beyond those, which were not considered.`,
              }
            : {}),
        }),
      };
    } catch (err) {
      return { ...toolError(err) };
    }
  },
};
