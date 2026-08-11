import type { McpResult } from './result.js';

/**
 * Defense-in-depth for planner write tools, mirroring graph-tools.ts's confirm
 * gate for generated destructive tools (see executeGraphTool + its private
 * isConfirmGateEnabled there) so this fork's tools behave identically to
 * upstream's: opt in per-deployment via MS365_MCP_REQUIRE_CONFIRM=true
 * (default off), and once on, refuse to touch data unless the caller passes
 * confirm: true.
 *
 * Utility tools (which is what every src/planner/ tool registers as, via
 * UTILITY_TOOLS in graph-tools.ts) never go through executeGraphTool, so
 * without this they would be the only destructive tools in the server an LLM
 * could fire unconfirmed. Every planner write tool should call
 * checkConfirmGate() first and return its result if non-null.
 */
export function isConfirmGateEnabled(): boolean {
  return process.env.MS365_MCP_REQUIRE_CONFIRM === 'true';
}

/**
 * Returns a refusal result if the gate is enabled and the caller has not
 * passed confirm: true, or null if the call may proceed. The refusal shape
 * mirrors executeGraphTool's so both paths look identical to a client.
 */
export function checkConfirmGate(
  toolName: string,
  params: { confirm?: unknown }
): (McpResult & Record<string, unknown>) | null {
  if (!isConfirmGateEnabled() || params.confirm === true) {
    return null;
  }
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          error: 'confirmation_required',
          tool: toolName,
          destructive: true,
          message:
            'This tool modifies user data. Re-call with parameter "confirm": true after the user has explicitly approved the operation.',
        }),
      },
    ],
    isError: true,
  };
}
