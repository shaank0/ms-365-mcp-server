export interface McpResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}

/** Success payload, serialized as JSON text. */
export function ok(payload: unknown): McpResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload) }] };
}

/**
 * Error result. Graph errors arrive as "Microsoft Graph API error: <status>
 * <text> - <body>"; we surface an actionable sentence instead of the raw
 * body, matching how upstream tools report failures.
 */
export function toolError(err: unknown): McpResult {
  const raw = err instanceof Error ? err.message : String(err);
  let message = raw;
  if (/\b403\b/.test(raw)) {
    message =
      "You don't have access to this Planner item, or it is in a group you are not a member of.";
  } else if (/\b404\b/.test(raw)) {
    message = 'No such Planner item — check the id.';
  } else if (/\b412\b/.test(raw)) {
    message =
      'The item was modified by someone else while this change was being applied. Read it again and retry.';
  }
  return { content: [{ type: 'text', text: JSON.stringify({ error: message }) }], isError: true };
}
