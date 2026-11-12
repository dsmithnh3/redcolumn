import type { Assembly, AssemblyVariable, CustomColumn } from '@nb/markup';
import { assemblyProposals, type AssemblyProposal, type ProposedItem } from '../assemblies/assemblyProposals';
import { assemblyProblems, MEASURED_NAMES } from '../assemblies/quantities';
import { evalFormula } from '../columns/formula';
import { BUILT_IN_COLUMNS } from '../columns/listColumns';
import { assemblyOf, findDocument, record, type AssistantDocument } from './assistant';

/**
 * The assembly-proposal tools. An AI client can only propose assemblies (a new one, or an edit
 * with `replaces`); each waits in the Review AI Proposals panel until a person approves or rejects
 * it. This is the boundary with an untrusted caller, so everything is checked here, all or
 * nothing: one bad assembly means none is kept.
 */
export const ASSEMBLY_PROPOSAL_METHODS = ['propose_assembly', 'list_assembly_proposals'] as const;

export type AssemblyProposalMethod = (typeof ASSEMBLY_PROPOSAL_METHODS)[number];

/** The most assemblies one call proposes; the most items and variables one assembly may have. */
const MAX_ASSEMBLIES = 10;
const MAX_ITEMS = 40;
const MAX_VARIABLES = 20;
/**
 * Length limits. Over-length is refused rather than cut, so a name that comes back in a later
 * `replaces` or item comparison is the name that was checked.
 */
const MAX_NAME = 200;
const MAX_UNIT = 40;
const MAX_DESCRIPTION = 2000;
const MAX_FORMULA = 1000;
/** Reasoning is cut, the way propose_takeoff cuts its reasoning. */
const MAX_REASONING = 2000;

const REVIEW_NOTE = 'Waiting for review: the person approves or rejects each one in Tools › Takeoff › Review AI Proposals. Nothing is added to the document until they approve it.';

/** Text, trimmed. Over-length is refused, naming the field and the limit. */
function bounded(raw: unknown, name: string, cap: number): string {
  if (typeof raw !== 'string') throw new Error(`${name} has to be text.`);
  const value = raw.trim();
  if (value.length > cap) throw new Error(`${name} has to be ${cap} characters or fewer.`);
  return value;
}

/** Characters that do not display: control, bidi override and isolate marks, zero-width. ZWJ and ZWNL are left out on purpose. */
const INVISIBLE = /[\u0000-\u001f\u007f-\u009f\u202a-\u202e\u2066-\u2069\u200b\u2060\ufeff]/;

/** Names and units cannot hide what a person reads; ZWJ and ZWNJ (U+200C, U+200D) are allowed, as some scripts need them. */
function plain(value: string, name: string): string {
  if (INVISIBLE.test(value)) throw new Error(`${name} cannot contain characters that do not display (control, bidi override or zero-width).`);
  return value;
}

/** Names compared as the app compares them: Unicode-normalised, trimmed, case-insensitive. */
const norm = (value: string) => value.trim().normalize('NFC').toLowerCase();

/** A number the caller sent: only a real finite number, nothing that could be coerced to one. */
function numberOrNone(raw: unknown, name: string): number | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw !== 'number' || !Number.isFinite(raw)) throw new Error(`${name} has to be a number.`);
  return raw;
}

/** Every name a formula may use besides the assembly's own variables: the measured values and the list's columns. */
function columnNames(columns: readonly CustomColumn[]): string[] {
  return [...BUILT_IN_COLUMNS.map((b) => b.label), ...columns.map((c) => c.name.trim())];
}

/** The assembly's variables: named, unique, and not the name of a measured value (which would win). */
function variablesOf(raw: unknown): AssemblyVariable[] {
  if (raw === undefined || raw === null) return [];
  if (!Array.isArray(raw)) throw new Error('variables is a list of { name, value }.');
  if (raw.length > MAX_VARIABLES) throw new Error(`at most ${MAX_VARIABLES} variables.`);
  const measured = new Set(MEASURED_NAMES.map((m) => norm(m.name)));
  const out: AssemblyVariable[] = [];
  const seen = new Map<string, string>();
  for (const v of raw) {
    const r = record(v);
    const name = plain(bounded(r.name, 'a variable name', MAX_NAME), 'a variable name');
    if (!name) throw new Error('name each variable.');
    const key = norm(name);
    if (measured.has(key)) throw new Error(`[${name}] is already a measured value a formula can use; name the variable something else.`);
    if (seen.has(key)) throw new Error(`two variables are named ${seen.get(key)}.`);
    seen.set(key, name);
    const value = numberOrNone(r.value, `${name}'s value`);
    if (value === undefined) throw new Error(`${name} needs a value.`);
    out.push({ name, value });
  }
  return out;
}

/**
 * The assembly's items: named and unique, with a formula that works out. The dry run gives each
 * variable its value and every measured value and list column 1, as `lookupFor` resolves them,
 * so a broken formula, an unknown name, a text result and a division by zero are caught here,
 * before anything is kept.
 */
function itemsOf(raw: unknown, variables: readonly AssemblyVariable[], columns: readonly CustomColumn[]): ProposedItem[] {
  if (!Array.isArray(raw) || !raw.length) throw new Error(`give the assembly's items: a list of { name, unit, formula }.`);
  if (raw.length > MAX_ITEMS) throw new Error(`at most ${MAX_ITEMS} items.`);
  const byName = new Set([...MEASURED_NAMES.map((m) => norm(m.name)), ...columnNames(columns).map(norm)]);
  const lookup = (name: string): number | undefined => {
    const key = norm(name);
    const v = variables.find((x) => norm(x.name) === key);
    if (v) return v.value;
    return byName.has(key) ? 1 : undefined;
  };
  const out: ProposedItem[] = [];
  const seen = new Map<string, string>();
  for (const i of raw) {
    const r = record(i);
    const name = plain(bounded(r.name, 'an item name', MAX_NAME), 'an item name');
    if (!name) throw new Error('name each item.');
    const key = norm(name);
    if (seen.has(key)) throw new Error(`two items are named ${seen.get(key)}.`);
    seen.set(key, name);
    const unit = plain(bounded(r.unit, `${name}'s unit`, MAX_UNIT), `${name}'s unit`);
    const formula = typeof r.formula === 'string' ? r.formula.trim() : '';
    if (!formula) throw new Error(`give ${name} a formula, such as [Length] / [Spacing] + 1.`);
    if (formula.length > MAX_FORMULA) throw new Error(`the formula for ${name} has to be ${MAX_FORMULA} characters or fewer.`);
    const checked = evalFormula(formula, lookup);
    if ('error' in checked) throw new Error(`the formula for ${name} does not work out: ${checked.error}`);
    if (typeof checked.value !== 'number') throw new Error(`the formula for ${name} has to work out to a number.`);
    const waste = numberOrNone(r.waste, `${name}'s waste`);
    if (waste !== undefined && (waste < 0 || waste > 100)) throw new Error(`${name}'s waste is a percentage from 0 to 100.`);
    if (r.roundUp !== undefined && r.roundUp !== null && typeof r.roundUp !== 'boolean') throw new Error(`${name}'s roundUp is true or false.`);
    const roundUp = r.roundUp === true;
    const roundTo = numberOrNone(r.roundTo, `${name}'s roundTo`);
    if (roundTo !== undefined && (roundTo < 1e-6 || roundTo > 1e6)) throw new Error(`${name}'s roundTo is the multiple to round up to, between 0.000001 and 1000000.`);
    const unitCost = numberOrNone(r.unitCost, `${name}'s unitCost`);
    if (unitCost !== undefined && (unitCost < 0 || unitCost > 1e9)) throw new Error(`${name}'s unitCost is a price from 0 to 1000000000.`);
    out.push({ name, unit, formula, ...(waste !== undefined && waste !== 0 ? { waste } : {}), ...(roundUp ? { roundUp: true, ...(roundTo !== undefined && roundTo !== 1 ? { roundTo } : {}) } : {}), ...(unitCost !== undefined ? { unitCost } : {}) });
  }
  return out;
}

/** One proposed assembly, checked against the document's own assemblies. Throws on anything bad. */
function proposedAssembly(doc: AssistantDocument, existing: readonly Assembly[], raw: unknown): Omit<AssemblyProposal, 'id' | 'source' | 'createdAt'> {
  const item = record(raw);
  const name = plain(bounded(item.name, 'name', MAX_NAME), 'name');
  if (!name) throw new Error('give the assembly a name.');
  // No description sent keeps the assembly's notes (an edit); an empty one clears them.
  const description = item.description === undefined || item.description === null ? undefined : bounded(item.description, 'description', MAX_DESCRIPTION);
  const reasoning = typeof item.reasoning === 'string' ? item.reasoning.trim().slice(0, MAX_REASONING) : '';
  if (!reasoning) throw new Error('say why (reasoning), so the person can check it.');
  const replaced = item.replaces === undefined || item.replaces === null ? undefined : assemblyOf(doc, typeof item.replaces === 'string' ? item.replaces.normalize('NFC') : item.replaces);
  const clash = existing.find((a) => norm(a.name) === norm(name) && a.id !== replaced?.id);
  if (clash) throw new Error(`${doc.name} already has an assembly "${clash.name}". To propose an edit of it, give replaces: "${clash.name}".`);
  const columns = doc.store.columnSet().columns;
  const variables = variablesOf(item.variables);
  const items = itemsOf(item.items, variables, columns);
  // Checked as the Assemblies window would check it, against the document's own assemblies and
  // columns: a variable named like a list column is refused, because the column wins at lookup
  // and the quantities would come out wrong.
  const candidate: Assembly = { id: replaced?.id ?? crypto.randomUUID(), name, ...(description !== undefined ? { description } : {}), variables, items: items.map((i) => ({ id: crypto.randomUUID(), ...i })) };
  const problem = [...assemblyProblems(candidate, replaced ? existing.filter((a) => a.id !== replaced.id) : existing, columns)][0];
  if (problem) {
    const about = problem[0].startsWith('item:') ? candidate.items.find((i) => `item:${i.id}` === problem[0])?.name : problem[0].startsWith('var:') ? candidate.variables[Number(problem[0].slice(4))]?.name : undefined;
    throw new Error(`${about ? `${about}: ` : ''}${problem[1]}`);
  }
  return { name, ...(description !== undefined ? { description } : {}), variables, items, ...(replaced ? { replacesId: replaced.id } : {}), reasoning };
}

function proposeAssemblies(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (!Array.isArray(args.assemblies) || !args.assemblies.length) throw new Error('Give assemblies: a list of { name, items, reasoning }, with replaces to propose an edit of one.');
  if (args.assemblies.length > MAX_ASSEMBLIES) throw new Error(`Propose at most ${MAX_ASSEMBLIES} assemblies at a time.`);
  const source = typeof args.source === 'string' && args.source.trim() ? args.source.trim().slice(0, 80) : 'AI assistant';
  const existing = doc.store.assemblies();
  const now = Date.now();
  // Every assembly is checked before any is kept.
  const proposals: AssemblyProposal[] = args.assemblies.map((raw: unknown, n: number) => {
    const rawName = record(raw).name;
    const label = `Assembly ${n + 1}${typeof rawName === 'string' && rawName.trim() ? ` (${rawName.trim().slice(0, MAX_NAME)})` : ''}`;
    try {
      return { id: crypto.randomUUID(), ...proposedAssembly(doc, existing, raw), source, createdAt: now };
    } catch (err) {
      throw new Error(`${label}: ${err instanceof Error ? err.message : String(err)}`);
    }
  });
  const names = new Map<string, number>();
  const edits = new Map<string, number>();
  proposals.forEach((p, n) => {
    const label = `Assembly ${n + 1} (${p.name})`;
    const first = names.get(norm(p.name));
    if (first !== undefined) throw new Error(`${label}: two assemblies in this call are named ${p.name}.`);
    names.set(norm(p.name), n);
    if (p.replacesId === undefined) return;
    const other = edits.get(p.replacesId);
    if (other !== undefined) {
      const target = existing.find((a) => a.id === p.replacesId);
      throw new Error(`${label} and Assembly ${other + 1} (${proposals[other]!.name}) both replace "${target?.name ?? p.replacesId}". Propose one edit of an assembly at a time.`);
    }
    edits.set(p.replacesId, n);
  });
  if (!assemblyProposals.add(doc.store.fileHash, proposals)) throw new Error("Could not keep the proposals: this browser's storage is full or unavailable.");
  return { document: doc.name, proposed: proposals.length, ids: proposals.map((p) => p.id), note: REVIEW_NOTE };
}

function listAssemblyProposals(docs: readonly AssistantDocument[], params: unknown) {
  const doc = findDocument(docs, record(params).document);
  const assemblies = doc.store.assemblies();
  return {
    document: doc.name,
    proposals: assemblyProposals.list(doc.store.fileHash).map((p) => ({
      id: p.id,
      name: p.name,
      ...(p.replacesId ? { replaces: assemblies.find((a) => a.id === p.replacesId)?.name ?? null } : {}),
      itemCount: p.items.length,
      state: p.rejected ? 'rejected' : 'pending',
      ...(p.rejected ? { rejectedBy: p.rejected.by } : {}),
    })),
  };
}

export async function assemblyProposalCall(docs: readonly AssistantDocument[], method: AssemblyProposalMethod, params: unknown): Promise<unknown> {
  if (method === 'propose_assembly') return proposeAssemblies(docs, params);
  return listAssemblyProposals(docs, params);
}
