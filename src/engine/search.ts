import type { Story, TextSpan } from './document';

export interface SearchRange {
  spanId: string;
  from: number;
  to: number;
}

export interface StoryMatch {
  ranges: SearchRange[];
  before: string;
  editable: boolean;
  reason?: string;
}

/** Formatting is transparent to search; every other code is a hard boundary. */
export function searchStory(story: Story, expression: RegExp): StoryMatch[] {
  const result: StoryMatch[] = [];
  let group: { span: TextSpan; start: number; end: number }[] = [];
  let text = '';
  const flush = () => {
    let cursor = 0;
    for (const match of text.matchAll(expression)) {
      const from = match.index;
      const to = from + match[0].length;
      while (cursor < group.length && group[cursor]!.end <= from) cursor++;
      const ranges: SearchRange[] = [];
      let reason: string | undefined;
      for (let i = cursor; i < group.length && group[i]!.start < to; i++) {
        const { span, start, end } = group[i]!;
        if (end <= from || start === end) continue;
        ranges.push({
          spanId: span.id,
          from: Math.max(0, from - start),
          to: Math.min(end, to) - start,
        });
        if (!span.editable) reason ??= span.reason ?? 'Protected text';
      }
      result.push({ ranges, before: match[0], editable: !reason, reason });
    }
    group = [];
    text = '';
  };
  for (const token of story.tokens) {
    if (token.kind === 'text') {
      group.push({
        span: token.span,
        start: text.length,
        end: text.length + token.span.text.length,
      });
      text += token.span.text;
    } else if (token.category !== 'format') {
      flush();
    }
  }
  flush();
  return result;
}
