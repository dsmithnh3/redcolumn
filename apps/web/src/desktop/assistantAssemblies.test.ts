import assert from 'node:assert/strict';
import test from 'node:test';
import { MarkupStore, type Assembly } from '@nb/markup';
import { assemblyProposals } from '../assemblies/assemblyProposals.ts';
import { assistantCall, type AssistantDocument } from './assistant.ts';

// The proposals are kept per document in localStorage; a stub lets the tests see what is kept.
const saved = new Map<string, string>();
(globalThis as { localStorage?: unknown }).localStorage = {
  getItem: (k: string) => saved.get(k) ?? null,
  setItem: (k: string, v: string) => void saved.set(k, v),
  removeItem: (k: string) => void saved.delete(k),
};

const wall: Assembly = {
  id: 'wall',
  name: 'Stud Wall',
  variables: [{ name: 'Spacing', value: 16 }],
  items: [{ id: 'studs', name: 'Studs', unit: 'ea', formula: '[Length] / [Spacing] + 1', unitCost: 3.5 }],
};
const slab: Assembly = { id: 'slab', name: 'Slab', variables: [], items: [{ id: 'conc', name: 'Concrete', unit: 'm³', formula: '[Area] * 0.1' }] };

async function drawing(assemblies: readonly Assembly[] = [wall]): Promise<AssistantDocument> {
  const store = await MarkupStore.open(`assembly-tool-${Math.random()}`, { persist: false });
  store.setAssemblies([...assemblies]);
  return { id: 'set', name: 'Set.pdf', pageCount: 1, pageIndex: 0, active: true, store, pageSize: () => ({ width: 1200, height: 800 }) };
}

/** A valid proposed assembly, with any field overridden. */
const ok = (over: Record<string, unknown> = {}) => ({ name: 'Pipe run', items: [{ name: 'Pipe', unit: 'm', formula: '[Length] * 1.05' }], reasoning: 'The detail shows pipe in 10 m lengths', ...over });

test('list_assemblies shows each item unit cost, and no item ids', async () => {
  const d = await drawing();
  const listed = (await assistantCall([d], 'list_assemblies', {})) as { assemblies: { items: Record<string, unknown>[] }[] };
  const item = listed.assemblies[0]!.items[0]!;
  assert.equal(item.unitCost, 3.5);
  assert.equal('id' in item, false);
  assert.equal('waste' in item, false, 'only what is set is listed');
});

test('an AI client proposes assemblies; they wait for a person and the document is unchanged', async () => {
  const d = await drawing();
  const made = (await assistantCall([d], 'propose_assembly', { source: 'Claude', assemblies: [ok({ variables: [{ name: 'Joint length', value: 10 }] })] })) as { proposed: number; ids: string[]; note: string };
  assert.equal(made.proposed, 1);
  assert.match(made.note, /Nothing is added to the document until they approve it/);
  assert.deepEqual(d.store.assemblies().map((a) => a.name), ['Stud Wall'], 'the document is unchanged');
  const pending = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: Record<string, unknown>[] };
  assert.deepEqual(pending.proposals.map((p) => [p.id, p.name, p.itemCount, p.state]), [[made.ids[0], 'Pipe run', 1, 'pending']]);
  assert.equal('items' in pending.proposals[0]!, false, 'the payload is not repeated in the list');
  assert.equal('replaces' in pending.proposals[0]!, false);
  assemblyProposals.reject(d.store.fileHash, made.ids, 'Hana');
  const decided = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: { state: string; rejectedBy?: string }[] };
  assert.deepEqual(decided.proposals.map((p) => [p.state, p.rejectedBy]), [['rejected', 'Hana']]);
});

test('every assembly is checked first: one bad one keeps them all out', async () => {
  const d = await drawing();
  await assert.rejects(
    () => assistantCall([d], 'propose_assembly', { assemblies: [ok(), ok({ name: 'Second', items: [{ name: 'Pipe', unit: 'm', formula: '1 +' }] })] }),
    /Assembly 2 \(Second\): the formula for Pipe does not work out/,
  );
  assert.deepEqual(((await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: unknown[] }).proposals, [], 'nothing was kept');
});

test('each rule says what is wrong', async () => {
  const d = await drawing();
  const refuses = (assemblies: unknown, message: RegExp) => assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies }), message);
  await refuses([], /Give assemblies/);
  await refuses([ok({ name: '' })], /give the assembly a name/);
  await refuses([ok({ reasoning: '  ' })], /say why/);
  await refuses([ok({ items: [] })], /give the assembly's items/);
  await refuses([ok({ items: Array.from({ length: 41 }, () => ({ name: `Item ${Math.random()}`, unit: 'ea', formula: '1' })) })], /at most 40 items/);
  await refuses([ok({ variables: Array.from({ length: 21 }, (_, i) => ({ name: `V ${i}`, value: 1 })) })], /at most 20 variables/);
  await refuses([ok({ items: [{ name: ' ', unit: 'm', formula: '[Length]' }] })], /name each item/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length]' }, { name: 'pipe', unit: 'm', formula: '[Length]' }] })], /two items are named Pipe/);
  await refuses([ok({ variables: [{ name: 'Spacing', value: 1 }, { name: 'spacing', value: 2 }] })], /two variables are named Spacing/);
  await refuses([ok({ variables: [{ name: 'Length', value: 1 }] })], /\[Length\] is already a measured value/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length]', waste: 101 }] })], /Pipe's waste is a percentage from 0 to 100/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length]', waste: -1 }] })], /Pipe's waste is a percentage from 0 to 100/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length]', roundUp: true, roundTo: 0 }] })], /Pipe's roundTo is the multiple to round up to, between 0\.000001 and 1000000/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length]', unitCost: -1 }] })], /Pipe's unitCost is a price from 0 to 1000000000/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length] + 1 *' }] })], /the formula for Pipe does not work out/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '[Length] / [Spacing]' }] })], /the formula for Pipe does not work out: No column named \[Spacing\]/);
  await refuses([ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1 / ([Count] - 1)' }] })], /the formula for Pipe does not work out: #DIV\/0/);
  await refuses([ok(), ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1' }] })], /two assemblies in this call are named Pipe run/);
  await refuses(Array.from({ length: 11 }, () => ok({ name: `A ${Math.random()}`, items: [{ name: 'Pipe', unit: 'm', formula: '1' }] })), /at most 10 assemblies/);
});

test('a variable named like a column never wins at lookup, so it is refused', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ variables: [{ name: 'Slope', value: 1 }] })] }), /Slope: \[Slope\] is already a measured value or column/);
  d.store.setColumnSet({ columns: [...d.store.columnSet().columns, { id: 'spacing', name: 'Spacing', type: 'number' }], statuses: d.store.columnSet().statuses });
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ variables: [{ name: 'Spacing', value: 1 }] })] }), /Spacing: \[Spacing\] is already a measured value or column/);
});

test('a formula that works out to text is refused', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '"N/A"' }] })] }), /the formula for Pipe has to work out to a number/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: 'CONCAT("a", "b")' }] })] }), /the formula for Pipe has to work out to a number/);
});

test('roundTo and unitCost have bounds; nothing that merely coerces to a number is taken', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1', roundUp: true, roundTo: 1e-320 }] })] }), /Pipe's roundTo is the multiple to round up to, between 0\.000001 and 1000000/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1', unitCost: 1e308 }] })] }), /Pipe's unitCost is a price from 0 to 1000000000/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1', waste: true }] })] }), /Pipe's waste has to be a number/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1', roundTo: '0.5' }] })] }), /Pipe's roundTo has to be a number/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1', unitCost: [5] }] })] }), /Pipe's unitCost has to be a number/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ variables: [{ name: 'Spacing', value: [] }] })] }), /Spacing's value has to be a number/);
  // The bounds themselves are fine.
  const made = (await assistantCall([d], 'propose_assembly', {
    assemblies: [ok({ items: [
      { name: 'Fine', unit: 'm', formula: '1', roundUp: true, roundTo: 1e-6 },
      { name: 'Coarse', unit: 'm', formula: '1', roundUp: true, roundTo: 1e6 },
      { name: 'Priced', unit: 'm', formula: '1', unitCost: 1e9 },
    ] })],
  })) as { proposed: number };
  assert.equal(made.proposed, 1);
});

test('a proposal whose target was deleted lists replaces: null', async () => {
  const d = await drawing();
  await assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Stud Wall', replaces: 'wall' })] });
  d.store.setAssemblies([]);
  const listed = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: { replaces: string | null }[] };
  assert.equal(listed.proposals[0]!.replaces, null);
});

test('a name the drawing already uses needs replaces; a replaces that matches nothing is refused', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'stud wall' })] }), /Set\.pdf already has an assembly "Stud Wall"\. To propose an edit of it, give replaces: "Stud Wall"\./);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ replaces: 'Slab' })] }), /Set\.pdf has no assembly "Slab"/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ replaces: '' })] }), /Say which assembly/);
  // An edit of that assembly may keep its name (or take a free one).
  await assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Stud Wall', replaces: 'stud wall' })] });
  const [edit] = assemblyProposals.list(d.store.fileHash);
  assert.equal(edit!.replacesId, 'wall');
  const listed = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: { replaces: string }[] };
  assert.equal(listed.proposals[0]!.replaces, 'Stud Wall');
});

test('an edit cannot take the name of another assembly', async () => {
  const d = await drawing([wall, slab]);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Slab', replaces: 'Stud Wall' })] }), /Set\.pdf already has an assembly "Slab"/);
  // Renaming onto a free name is fine.
  await assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Wall 2', replaces: 'wall' })] });
  const [edit] = assemblyProposals.list(d.store.fileHash);
  assert.equal(edit!.replacesId, 'wall');
});

test("a formula may use the list's columns, and list_assemblies names them", async () => {
  const d = await drawing();
  d.store.setColumnSet({ columns: [{ id: 'pitch', name: 'Pitch factor', type: 'number' }], statuses: d.store.columnSet().statuses });
  const listed = (await assistantCall([d], 'list_assemblies', {})) as { columnNames: string[]; formulaNames: string[] };
  assert.ok(listed.columnNames.includes('Pitch factor'), 'the document column');
  assert.ok(listed.columnNames.includes('Slope'), 'a built-in column label');
  assert.ok(!listed.formulaNames.includes('Slope'), 'formulaNames stays the measured values');
  const made = (await assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Sheet', unit: 'm²', formula: '[Area] * [Pitch factor] / 100' }] })] })) as { proposed: number };
  assert.equal(made.proposed, 1, 'a column formula is accepted, as the real lookup resolves columns');
  // An edit that renames the assembly and keeps the column formula round-trips.
  const edit = (await assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Roof sheet', replaces: 'wall', items: [{ name: 'Sheet', unit: 'm²', formula: '[Area] * [Pitch factor] / 100' }] })] })) as { proposed: number };
  assert.equal(edit.proposed, 1);
  const p = assemblyProposals.list(d.store.fileHash).find((x) => x.name === 'Roof sheet')!;
  assert.equal(p.replacesId, 'wall');
});

test('over-length fields are refused, not cut', async () => {
  const d = await drawing();
  const long = 'x'.repeat(201);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: long })] }), /Assembly 1 .*: name has to be 200 characters or fewer/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm'.repeat(41), formula: '1' }] })] }), /Pipe's unit has to be 40 characters or fewer/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ description: 'd'.repeat(2001) })] }), /description has to be 2000 characters or fewer/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ variables: [{ name: long, value: 1 }] })] }), /a variable name has to be 200 characters or fewer/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: long, unit: 'ea', formula: '1' }] })] }), /an item name has to be 200 characters or fewer/);
});

test('a formula has a length limit', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm', formula: '1'.repeat(1001) }] })] }), /the formula for Pipe has to be 1000 characters or fewer/);
});

test("when the browser will not keep the proposals, the call says so and keeps nothing", async () => {
  const d = await drawing();
  const original = (globalThis as { localStorage?: unknown }).localStorage;
  (globalThis as { localStorage?: unknown }).localStorage = {
    getItem: () => null,
    setItem: () => {
      throw new Error('full');
    },
    removeItem: () => {},
  };
  try {
    await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok()] }), /Could not keep the proposals: this browser's storage is full or unavailable/);
  } finally {
    (globalThis as { localStorage?: unknown }).localStorage = original;
  }
  const none = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: unknown[] };
  assert.deepEqual(none.proposals, [], 'nothing was kept in memory either');
});

test('two edits of the same assembly in one call are refused, naming both', async () => {
  const d = await drawing();
  const item = { name: 'Studs', unit: 'ea', formula: '[Length] * 2' };
  await assert.rejects(
    () => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'A', replaces: 'wall', items: [item] }), ok({ name: 'B', replaces: 'Stud Wall', items: [item] })] }),
    /Assembly 2 \(B\) and Assembly 1 \(A\) both replace "Stud Wall"/,
  );
  const none = (await assistantCall([d], 'list_assembly_proposals', {})) as { proposals: unknown[] };
  assert.deepEqual(none.proposals, [], 'nothing was kept');
});

test('names cannot hide characters that do not display; ZWJ is fine', async () => {
  const d = await drawing();
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Pipe\u202Erun' })] }), /name cannot contain characters that do not display/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Pi\u0007pe' })] }), /name cannot contain characters that do not display/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ items: [{ name: 'Pipe', unit: 'm\u200B', formula: '1' }] })] }), /Pipe's unit cannot contain characters that do not display/);
  await assert.rejects(() => assistantCall([d], 'propose_assembly', { assemblies: [ok({ variables: [{ name: 'Spacing\u2066', value: 1 }] })] }), /a variable name cannot contain characters that do not display/);
  const made = (await assistantCall([d], 'propose_assembly', { assemblies: [ok({ name: 'Na\u200Dm' })] })) as { proposed: number };
  assert.equal(made.proposed, 1, 'ZWJ is allowed, as some scripts need it');
});
