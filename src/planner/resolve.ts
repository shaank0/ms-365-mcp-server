import { getJson, paginate, GUID, PLANNER_ID, type GraphLike } from './graph.js';

/** A resolution failure the caller should surface verbatim — it names the valid options. */
export class PlannerResolveError extends Error {}

interface DirectoryGroup {
  id: string;
  displayName?: string;
  groupTypes?: string[];
}

/**
 * Resolve a Microsoft 365 group by display name or id. Names are matched
 * against the CALLER'S OWN memberships, so this can never resolve a group
 * the caller does not belong to.
 */
export async function resolveGroup(g: GraphLike, value: string): Promise<string> {
  const wanted = value.trim();
  if (GUID.test(wanted)) return wanted;

  const { items } = await paginate<DirectoryGroup>(
    g,
    '/me/memberOf?$select=id,displayName,groupTypes',
    500
  );
  const groups = items.filter((i) => (i.groupTypes ?? []).includes('Unified'));
  const target = wanted.toLowerCase();
  const exact = groups.filter((x) => (x.displayName ?? '').toLowerCase() === target);

  if (exact.length === 1) return exact[0].id;
  if (exact.length > 1) {
    const list = exact.map((x) => `${x.displayName} (${x.id})`).join(', ');
    throw new PlannerResolveError(
      `More than one group is named "${wanted}". Pass the group id instead — candidates: ${list}`
    );
  }

  const near = groups
    .filter((x) => (x.displayName ?? '').toLowerCase().includes(target))
    .slice(0, 10)
    .map((x) => x.displayName);
  throw new PlannerResolveError(
    near.length
      ? `You are not a member of any group named exactly "${wanted}". Did you mean: ${near.join(', ')}?`
      : `You are not a member of any group named "${wanted}".`
  );
}

/** Resolve a bucket by name within a plan. Never creates one on a miss. */
export async function resolveBucket(g: GraphLike, planId: string, value: string): Promise<string> {
  const wanted = value.trim();
  if (PLANNER_ID.test(wanted)) return wanted;

  const { items } = await paginate<{ id: string; name?: string }>(
    g,
    `/planner/plans/${planId}/buckets`,
    200
  );
  const target = wanted.toLowerCase();
  const exact = items.filter((b) => (b.name ?? '').toLowerCase() === target);

  if (exact.length === 1) return exact[0].id;
  if (exact.length > 1) {
    throw new PlannerResolveError(
      `More than one bucket in this plan is named "${wanted}". Pass the bucket id instead.`
    );
  }

  const available =
    items
      .map((b) => b.name)
      .filter(Boolean)
      .join(', ') || '(none)';
  throw new PlannerResolveError(
    `This plan has no bucket named "${wanted}". Available buckets: ${available}. Create one with create-planner-bucket first.`
  );
}

/**
 * Resolve assignees to the plannerAssignment dictionary Graph expects.
 * Names resolve against the plan's container group members — Planner only
 * permits assigning users who have access to the plan, and Group.Read.All
 * already covers this, so no directory-wide lookup is needed.
 */
export async function resolveAssignees(
  g: GraphLike,
  planId: string,
  values: string[]
): Promise<Record<string, unknown>> {
  const plan = await getJson<{ container?: { containerId?: string } }>(
    g,
    `/planner/plans/${planId}`
  );
  const groupId = plan.container?.containerId;
  if (!groupId) {
    throw new PlannerResolveError(
      `Plan ${planId} has no group container, so assignees cannot be resolved. Assign the task in the Planner UI.`
    );
  }

  const { items } = await paginate<{
    id: string;
    mail?: string | null;
    userPrincipalName?: string | null;
  }>(g, `/groups/${groupId}/members?$select=id,mail,userPrincipalName,displayName`, 500);

  const assignments: Record<string, unknown> = {};
  for (const raw of values) {
    const wanted = raw.trim();
    let id: string | undefined;

    if (GUID.test(wanted)) {
      id = wanted;
    } else {
      const target = wanted.toLowerCase();
      id = items.find(
        (m) =>
          (m.mail ?? '').toLowerCase() === target ||
          (m.userPrincipalName ?? '').toLowerCase() === target
      )?.id;
    }

    if (!id) {
      const members =
        items
          .map((m) => m.mail ?? m.userPrincipalName)
          .filter(Boolean)
          .join(', ') || '(none)';
      throw new PlannerResolveError(
        `"${raw}" is not a member of this plan's group, so they cannot be assigned. Members: ${members}`
      );
    }

    assignments[id] = {
      '@odata.type': '#microsoft.graph.plannerAssignment',
      orderHint: ' !',
    };
  }

  return assignments;
}
