import { describe, it, expect } from 'vitest';
import { toPriority, statusToPercent, toChecklist, toReplacement } from '../convert.js';

describe('toPriority', () => {
  it.each([
    ['urgent', 1],
    ['important', 3],
    ['medium', 5],
    ['low', 9],
    ['  URGENT  ', 1],
  ])('maps %s to %i', (word, expected) => {
    expect(toPriority(word as string)).toBe(expected);
  });

  it('passes integers 0-10 through', () => {
    expect(toPriority(0)).toBe(0);
    expect(toPriority(10)).toBe(10);
  });

  it('explicitly pins priority 0', () => {
    expect(toPriority(0)).toBe(0);
  });

  it('rejects out-of-range integers', () => {
    expect(() => toPriority(11)).toThrow(/0-10/);
  });

  it.each([[-1], [3.5], [NaN], [Infinity]])('rejects numeric edge case %s', (value) => {
    expect(() => toPriority(value)).toThrow(/0-10/);
  });

  it('rejects unknown words', () => {
    expect(() => toPriority('spicy')).toThrow(/urgent/);
  });

  it.each([['constructor'], ['toString'], ['hasOwnProperty']])(
    'rejects prototype property %s',
    (prop) => {
      expect(() => toPriority(prop)).toThrow(/urgent/);
    }
  );
});

describe('statusToPercent', () => {
  it.each([
    ['complete', 100],
    ['in-progress', 50],
    ['not-started', 0],
  ])('maps %s to %i', (status, expected) => {
    expect(statusToPercent(status as string)).toBe(expected);
  });

  it('trims and lowercases status strings', () => {
    expect(statusToPercent('  Complete  ')).toBe(100);
  });

  it('rejects an unknown status', () => {
    expect(() => statusToPercent('done-ish')).toThrow(/complete/);
  });

  it.each([['constructor'], ['valueOf']])('rejects prototype property %s', (prop) => {
    expect(() => statusToPercent(prop)).toThrow(/complete/);
  });
});

describe('toChecklist', () => {
  it('builds a GUID-keyed dict from plain strings', () => {
    let n = 0;
    const result = toChecklist(['a', 'b'], () => `id-${++n}`);
    expect(result).toEqual({
      'id-1': {
        '@odata.type': '#microsoft.graph.plannerChecklistItem',
        title: 'a',
        isChecked: false,
      },
      'id-2': {
        '@odata.type': '#microsoft.graph.plannerChecklistItem',
        title: 'b',
        isChecked: false,
      },
    });
  });

  it('honours the checked flag on object form', () => {
    const result = toChecklist([{ title: 'a', checked: true }], () => 'k');
    expect((result as any).k.isChecked).toBe(true);
  });

  it('generates distinct keys by default', () => {
    expect(Object.keys(toChecklist(['a', 'b', 'c'])).length).toBe(3);
  });

  it('returns empty dict for empty list', () => {
    expect(toChecklist([])).toEqual({});
  });
});

describe('toReplacement', () => {
  it('nulls every existing key when entries is empty (genuine clear)', () => {
    expect(toReplacement(['a', 'b'], {})).toEqual({ a: null, b: null });
  });

  it('nulls stale keys AND keeps/adds entries keys', () => {
    expect(toReplacement(['a', 'b'], { b: { v: 2 }, c: { v: 3 } })).toEqual({
      a: null,
      b: { v: 2 },
      c: { v: 3 },
    });
  });

  it('passes entries through unchanged when there is nothing existing to clear', () => {
    expect(toReplacement([], { c: { v: 3 } })).toEqual({ c: { v: 3 } });
  });

  it('returns an empty object when both existingKeys and entries are empty', () => {
    expect(toReplacement([], {})).toEqual({});
  });
});
