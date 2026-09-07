import {
  descendants,
  orderedPatches,
  parseXml,
  patchXml,
  type XmlNode,
  type XmlPatch,
} from './xml';

// New identities are session-local and never reused across branches or discarded staging.
let generation = 0;
export function freshXmlIds(part: string, count: number): string[] {
  const batch = ++generation;
  return Array.from({ length: count }, (_, index) => `${part}#g${batch}.${index}`);
}

export interface XmlIdentityChanges {
  retained: string[];
  created: string[];
  removed: string[];
}

/** Reports node continuity only; split/join commands must separately map text positions. */
export function compareXmlIdentities(before: XmlNode, after: XmlNode): XmlIdentityChanges {
  const left = new Set(descendants(before).map((node) => node.id));
  const right = new Set(descendants(after).map((node) => node.id));
  return {
    retained: [...left].filter((id) => right.has(id)),
    created: [...right].filter((id) => !left.has(id)),
    removed: [...left].filter((id) => !right.has(id)),
  };
}

/** No content or ordinal matching: continuity is proven by patch geometry or declared explicitly. */
export function remapXmlIdentities(
  source: string,
  part: string,
  before: XmlNode,
  patches: readonly XmlPatch[],
): { source: string; ids: readonly string[] } {
  const ordered = orderedPatches(source, patches);
  const updated = patchXml(source, ordered);
  const nodes = descendants(parseXml(updated, part));
  const oldNodes = descendants(before);
  const oldIds = new Map(oldNodes.map((node) => [node.id, node]));
  const newStarts = new Map(nodes.map((node, index) => [node.start, index]));
  const assigned = new Map<number, string>();
  const used = new Set<string>();
  const deltas = [0];
  for (const patch of ordered)
    deltas.push(deltas.at(-1)! + patch.replacement.length - (patch.end - patch.start));

  // Index of the first patch ending after offset. Insertions at offset precede the old tag.
  const after = (offset: number) => {
    let low = 0,
      high = ordered.length;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (ordered[mid]!.end <= offset) low = mid + 1;
      else high = mid;
    }
    return low;
  };
  const openingSurvives = (node: XmlNode) => {
    const patch = ordered[after(node.start)];
    return !patch || patch.start >= node.openEnd;
  };
  const assign = (old: XmlNode, start: number, explicit: boolean) => {
    const index = newStarts.get(start);
    const node = index === undefined ? undefined : nodes[index];
    if (!node || node.uri !== old.uri || node.local !== old.local) {
      if (explicit) throw new Error(`Invalid retained node target: ${old.id}`);
      return;
    }
    if (assigned.has(index!) || used.has(old.id))
      throw new Error(`Duplicate retained identity: ${old.id}`);
    assigned.set(index!, old.id);
    used.add(old.id);
  };
  for (const node of oldNodes) {
    if (openingSurvives(node)) assign(node, node.start + deltas[after(node.start)]!, false);
  }
  for (const [index, patch] of ordered.entries()) {
    for (const retained of patch.retain ?? []) {
      const node = oldIds.get(retained.nodeId);
      if (!node || openingSurvives(node))
        throw new Error(
          `Retained identity must name a rewritten or moved node: ${retained.nodeId}`,
        );
      if (
        !Number.isInteger(retained.offset) ||
        retained.offset < 0 ||
        retained.offset >= patch.replacement.length
      )
        throw new Error(`Invalid retained node offset: ${retained.nodeId}`);
      const start = patch.start + deltas[index]! + retained.offset;
      assign(node, start, true);
    }
  }
  const fresh = freshXmlIds(part, nodes.length - assigned.size);
  let next = 0;
  return { source: updated, ids: nodes.map((_, index) => assigned.get(index) ?? fresh[next++]!) };
}
