import { PLANNER_ID } from './graph.js';

/**
 * Second line of defense against path-traversal-style ids, alongside the Zod
 * regex each tool's buildSchema already declares. Zod only runs on the
 * normal MCP registration path (server.tool parses input before the handler
 * runs) - discovery mode's execute-tool calls utility.execute(parameters,
 * ctx) directly with raw, unparsed client input (see graph-tools.ts's
 * execute-tool handler), so a tool that only trusted the schema is still
 * exploitable there. Every planner tool should re-validate id-shaped
 * parameters at the top of execute(), inside the same try/catch that already
 * converts thrown errors via toolError() - so this throws rather than
 * returning an McpResult itself, keeping every tool's execute() to one error
 * path instead of two.
 */
function requireMatch(value: unknown, paramName: string, pattern: RegExp, shape: string): string {
  if (typeof value !== 'string' || !pattern.test(value)) {
    throw new Error(
      `${paramName} must be ${shape} — not a path or URL. Got: ${JSON.stringify(value)}`
    );
  }
  return value;
}

/** Validates a plan/bucket/task id (28-character opaque string, see PLANNER_ID). */
export function requirePlannerId(value: unknown, paramName: string): string {
  return requireMatch(
    value,
    paramName,
    PLANNER_ID,
    'a 28-character Planner id (letters, digits, "_" or "-" only)'
  );
}
