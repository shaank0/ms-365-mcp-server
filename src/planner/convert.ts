import { randomUUID } from 'node:crypto';

const PRIORITY_WORDS: Record<string, number> = Object.create(null);
PRIORITY_WORDS.urgent = 1;
PRIORITY_WORDS.important = 3;
PRIORITY_WORDS.medium = 5;
PRIORITY_WORDS.low = 9;

const STATUS_PERCENT: Record<string, number> = Object.create(null);
STATUS_PERCENT.complete = 100;
STATUS_PERCENT['in-progress'] = 50;
STATUS_PERCENT['not-started'] = 0;

export function toPriority(value: string | number): number {
  if (typeof value === 'number') {
    if (!Number.isInteger(value) || value < 0 || value > 10) {
      throw new Error(`priority must be an integer 0-10, got ${value}`);
    }
    return value;
  }
  const mapped = PRIORITY_WORDS[value.trim().toLowerCase()];
  if (mapped === undefined) {
    throw new Error(
      `priority must be one of urgent, important, medium, low — or an integer 0-10. Got "${value}".`
    );
  }
  return mapped;
}

export function statusToPercent(status: string): number {
  const mapped = STATUS_PERCENT[status.trim().toLowerCase()];
  if (mapped === undefined) {
    throw new Error(`status must be one of complete, in-progress, not-started. Got "${status}".`);
  }
  return mapped;
}

export function toChecklist(
  items: Array<string | { title: string; checked?: boolean }>,
  makeId: () => string = randomUUID
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const item of items) {
    const title = typeof item === 'string' ? item : item.title;
    const isChecked = typeof item === 'string' ? false : Boolean(item.checked);
    out[makeId()] = {
      '@odata.type': '#microsoft.graph.plannerChecklistItem',
      title,
      isChecked,
    };
  }
  return out;
}
