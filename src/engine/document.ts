import { DocxPackage, type Diagnostic } from '../package/docx';
import { child, descendants, isWord, wordAttr, type XmlNode } from '../package/xml';

export type Formatting = Record<string, string>;
export interface SourceBinding {
  part: string;
  nodeId: string;
  start: number;
  end: number;
}
export interface TextSpan {
  id: string;
  text: string;
  source: SourceBinding;
  editable: boolean;
  reason?: string;
  direct: Formatting;
  inherited: Formatting;
}
export interface CodeToken {
  kind: 'code';
  id: string;
  label: string;
  category: 'structure' | 'format' | 'review' | 'opaque';
  source: SourceBinding;
  details: Record<string, unknown>;
}
export interface TextToken {
  kind: 'text';
  span: TextSpan;
}
export type Token = CodeToken | TextToken;
export interface Story {
  id: string;
  part: string;
  kind: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'comment';
  label: string;
  tokens: Token[];
}
export interface DocumentModel {
  stories: Story[];
  diagnostics: Diagnostic[];
}

function properties(node: XmlNode | undefined): Formatting {
  const result: Formatting = {};
  for (const prop of node?.children ?? []) {
    if (!isWord(prop)) continue;
    result[prop.local] =
      wordAttr(prop, 'val') ?? (Object.keys(prop.attrs).length ? JSON.stringify(prop.attrs) : 'on');
  }
  return result;
}

/** Small, explicit style cascade; numbering/table/conditional styles are not resolved. */
function styleResolver(pkg: DocxPackage) {
  const link = pkg.relationships.find(
    (r) => r.source === pkg.mainPart && r.type.endsWith('/styles') && !r.external,
  );
  const root = link && pkg.names().includes(link.target) ? pkg.xml(link.target) : undefined;
  const styles = new Map(
    (root?.children.filter((n) => isWord(n, 'style')) ?? []).map((n) => [
      wordAttr(n, 'styleId'),
      n,
    ]),
  );
  const defaults = properties(
    root &&
      child(child(root, 'docDefaults') ?? root, 'rPrDefault')?.children.find((n) =>
        isWord(n, 'rPr'),
      ),
  );
  const defaultParagraph = [...styles.values()].find(
    (n) =>
      wordAttr(n, 'type') === 'paragraph' &&
      ['1', 'true', 'on'].includes(wordAttr(n, 'default') ?? ''),
  );
  function resolve(id: string | undefined, seen = new Set<string>()): Formatting {
    if (!id || seen.has(id)) return {};
    seen.add(id);
    const style = styles.get(id);
    if (!style) return {};
    return {
      ...resolve(child(style, 'basedOn') && wordAttr(child(style, 'basedOn')!, 'val'), seen),
      ...properties(child(style, 'rPr')),
    };
  }
  return (paragraphStyle?: string, runStyle?: string): Formatting => ({
    ...defaults,
    ...resolve(paragraphStyle ?? (defaultParagraph && wordAttr(defaultParagraph, 'styleId'))),
    ...resolve(runStyle),
  });
}

export function readDocument(pkg: DocxPackage): DocumentModel {
  const stories: Story[] = [];
  const diagnostics = [...pkg.diagnostics];
  const inherit = styleResolver(pkg);
  const globallyProtected = pkg.diagnostics.some(
    (d) => d.code === 'signed-package' || d.severity === 'error',
  );
  const settings = pkg.relationships.find(
    (r) => r.source === pkg.mainPart && r.type.endsWith('/settings') && !r.external,
  );
  const settingsNodes =
    settings && pkg.names().includes(settings.target) ? descendants(pkg.xml(settings.target)) : [];
  const restricted = settingsNodes.some(
    (n) =>
      isWord(n, 'documentProtection') &&
      ['1', 'true', 'on'].includes(wordAttr(n, 'enforcement') ?? ''),
  );
  const tracking = settingsNodes.some(
    (n) => isWord(n, 'trackRevisions') && !['0', 'false', 'off'].includes(wordAttr(n, 'val') ?? ''),
  );
  if (restricted || tracking)
    diagnostics.push({
      severity: 'warning',
      code: 'protected-editing',
      message: restricted
        ? 'Document protection is enforced; editing is disabled.'
        : 'Track Changes is enabled; editing is disabled until revision creation is supported.',
    });
  const candidates: { part: string; kind: Story['kind'] }[] = [
    { part: pkg.mainPart, kind: 'body' },
  ];
  for (const rel of pkg.relationships) {
    if (rel.external || rel.source !== pkg.mainPart || !pkg.names().includes(rel.target)) continue;
    const type = rel.type.split('/').at(-1);
    const storyKinds: Record<string, Story['kind']> = {
      header: 'header',
      footer: 'footer',
      footnotes: 'footnote',
      endnotes: 'endnote',
      comments: 'comment',
    };
    const kind = type ? storyKinds[type] : undefined;
    if (kind && !candidates.some((c) => c.part === rel.target))
      candidates.push({ part: rel.target, kind });
  }
  for (const { part, kind } of candidates) {
    const xml = pkg.text(part);
    const root = pkg.xml(part);
    const containers = ['footnote', 'endnote', 'comment'].includes(kind)
      ? root.children.filter((n) => isWord(n, kind))
      : [kind === 'body' ? (child(root, 'body') ?? root) : root];
    for (const container of containers) {
      const recordId = wordAttr(container, 'id');
      const story: Story = {
        id: `${part}:${recordId ?? 'main'}`,
        part,
        kind,
        label: `${kind}${recordId === undefined ? '' : ` ${recordId}`}`,
        tokens: [],
      };
      const binding = (node: XmlNode): SourceBinding => ({
        part,
        nodeId: node.id,
        start: node.start,
        end: node.end,
      });
      const code = (node: XmlNode, label: string, category: CodeToken['category'], extra = '') => {
        story.tokens.push({
          kind: 'code',
          id: `${node.id}:${extra || label}`,
          label,
          category,
          source: binding(node),
          details: {
            attributes: node.attrs,
            xml: xml.slice(node.start, Math.min(node.end, node.start + 4000)),
          },
        });
      };
      // Inspect the full XML, including opaque wrappers: a field can start inside
      // unsupported markup and end in an otherwise editable run several paragraphs later.
      const fieldText = new Set<string>();
      const unbalancedFields = (() => {
        let depth = 0;
        for (const n of descendants(container)) {
          if (isWord(n, 'fldChar')) {
            const type = wordAttr(n, 'fldCharType');
            if (type === 'begin') depth++;
            if (type === 'end' && --depth < 0) return true;
          }
          if (depth > 0 && isWord(n, 't')) fieldText.add(n.id);
        }
        return depth !== 0;
      })();
      function walk(
        node: XmlNode,
        context: {
          paragraphStyle?: string;
          direct?: Formatting;
          inherited?: Formatting;
          reason?: string;
        } = {},
      ): void {
        if (!isWord(node)) {
          code(node, `Opaque: ${node.name}`, 'opaque');
          return;
        }
        const name = node.local;
        if (['pPr', 'rPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr'].includes(name)) {
          if (node.children.length) code(node, name, 'format');
          return;
        }
        if (name === 'p') {
          const props = child(node, 'pPr');
          const style = props && child(props, 'pStyle');
          const paragraphStyle = style && wordAttr(style, 'val');
          code(
            node,
            paragraphStyle ? `Paragraph · ${paragraphStyle}` : 'Paragraph',
            'structure',
            'open',
          );
          for (const n of node.children) walk(n, { ...context, paragraphStyle });
          code(node, '¶', 'structure', 'close');
          return;
        }
        if (name === 'r') {
          const direct = properties(child(node, 'rPr'));
          const inherited = inherit(context.paragraphStyle, direct.rStyle);
          if (Object.keys(direct).length || Object.keys(inherited).length)
            code(
              node,
              `Run · ${
                Object.entries(direct)
                  .map(([k, v]) => `${k}=${v}`)
                  .join(', ') || 'inherited'
              }`,
              'format',
              'open',
            );
          const unsafeRun = node.children.some(
            (n) => isWord(n, 'rPr') && descendants(n).some((p) => isWord(p, 'rPrChange')),
          );
          for (const n of node.children.filter((n) => !isWord(n, 'rPr')))
            walk(n, {
              ...context,
              direct,
              inherited,
              reason: context.reason ?? (unsafeRun ? 'Tracked formatting is preserved' : undefined),
            });
          if (Object.keys(direct).length || Object.keys(inherited).length)
            code(node, '/Run', 'format', 'close');
          return;
        }
        if (name === 't' || name === 'delText') {
          const inner = xml.slice(node.openEnd, node.closeStart);
          const reason =
            context.reason ??
            (globallyProtected || restricted || tracking
              ? 'Package protection or diagnostics prevent editing'
              : kind === 'comment'
                ? 'Comment records are read-only'
                : unbalancedFields || fieldText.has(node.id)
                  ? 'Field content is preserved'
                  : name === 'delText'
                    ? 'Tracked deletion is preserved'
                    : inner.includes('<')
                      ? 'Text contains XML markup'
                      : undefined);
          story.tokens.push({
            kind: 'text',
            span: {
              id: node.id,
              text: node.text,
              source: binding(node),
              editable: !reason,
              reason,
              direct: context.direct ?? {},
              inherited: context.inherited ?? {},
            },
          });
          return;
        }
        if (name === 'fldChar') {
          const type = wordAttr(node, 'fldCharType');
          code(node, `Field ${type}`, 'opaque');
          return;
        }
        if (['ins', 'del', 'moveFrom', 'moveTo'].includes(name)) {
          code(node, `${name} · ${wordAttr(node, 'author') ?? ''}`, 'review', 'open');
          for (const n of node.children)
            walk(n, { ...context, reason: 'Tracked revisions are preserved' });
          code(node, `/${name}`, 'review', 'close');
          return;
        }
        if (['body', 'hdr', 'ftr', 'footnote', 'endnote', 'comment'].includes(name)) {
          for (const n of node.children) walk(n, context);
          return;
        }
        if (['tbl', 'tr', 'tc', 'hyperlink'].includes(name)) {
          code(node, name, 'structure', 'open');
          for (const n of node.children) walk(n, context);
          code(node, `/${name}`, 'structure', 'close');
          return;
        }
        if (
          [
            'commentRangeStart',
            'commentRangeEnd',
            'commentReference',
            'bookmarkStart',
            'bookmarkEnd',
          ].includes(name)
        ) {
          code(node, `${name} ${wordAttr(node, 'id') ?? ''}`, 'review');
          return;
        }
        if (
          [
            'tab',
            'br',
            'cr',
            'sectPr',
            'footnoteReference',
            'endnoteReference',
            'lastRenderedPageBreak',
          ].includes(name)
        ) {
          code(node, name, 'structure');
          return;
        }
        code(node, `Opaque: ${node.name}`, 'opaque');
        diagnostics.push({
          severity: 'info',
          code: 'opaque-content',
          part,
          message: `${node.name} preserved as an opaque object (${node.id})`,
        });
      }
      walk(container);
      stories.push(story);
    }
  }
  return { stories, diagnostics };
}

export function spans(model: DocumentModel): TextSpan[] {
  return model.stories.flatMap((s) => s.tokens.flatMap((t) => (t.kind === 'text' ? [t.span] : [])));
}

/** This is a review projection: revision text is included, opaque objects are marked. */
export function plainText(story: Story): string {
  return story.tokens
    .map((t) =>
      t.kind === 'text'
        ? t.span.text
        : t.label === '¶' || t.label === 'br' || t.label === 'cr'
          ? '\n'
          : t.label === 'tab'
            ? '\t'
            : t.category === 'opaque'
              ? `[${t.label}]`
              : '',
    )
    .join('');
}
