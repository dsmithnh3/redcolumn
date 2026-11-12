import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MarkupStore, type Assembly } from '@nb/markup';
import { approveAssemblyProposals, assemblyProposals, newlyWaiting, type AssemblyProposal } from './assemblyProposals.ts';
import { loadLibrary, writeLibrary } from './library.ts';

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
  description: 'Partition wall',
  variables: [{ name: 'Spacing', value: 16 }],
  items: [
    { id: 'studs', name: 'Studs', unit: 'ea', formula: '[Length] / [Spacing] + 1', roundUp: true },
    { id: 'board', name: 'Board', unit: 'sheets', formula: '[Wall Area] / 32' },
  ],
};

const proposal = (over: Partial<AssemblyProposal> = {}): AssemblyProposal => ({
  id: `p${Math.random()}`,
  name: 'Pipe run',
  variables: [{ name: 'Length per pipe', value: 10 }],
  items: [{ name: 'Pipe', unit: 'lf', formula: '[Length] * 1.05', waste: 5, roundUp: true, roundTo: 10, unitCost: 4.2 }],
  reasoning: 'The detail shows pipe in 10 ft lengths with 5% waste',
  source: 'Claude',
  createdAt: 1,
  ...over,
});

const slab: Assembly = { id: 'slab', name: 'Slab', variables: [], items: [{ id: 'conc', name: 'Concrete', unit: 'm³', formula: '[Area] * 0.1' }] };

async function setup(assemblies: Assembly[] = [wall]) {
  const store = await MarkupStore.open(`assembly-proposals-${Math.random()}`, { persist: false });
  if (assemblies.length) store.setAssemblies(assemblies);
  return store;
}

test('proposals wait on this device, per document; a new one for the same assembly replaces the older pending one', async () => {
  const a = await setup();
  const b = await setup();
  const p1 = proposal();
  const p2 = proposal({ id: 'p2', items: [{ name: 'Pipe', unit: 'lf', formula: '[Length]' }] });
  assemblyProposals.add(a.fileHash, [p1]);
  assemblyProposals.add(b.fileHash, [proposal({ id: 'other' })]);
  assert.deepEqual(assemblyProposals.list(a.fileHash).map((p) => p.id), [p1.id]);
  assert.deepEqual(assemblyProposals.list(b.fileHash).map((p) => p.id), ['other'], 'another document has its own list');
  assert.ok(JSON.parse(saved.get(`nb.assemblyProposals.${a.fileHash}`)!).some((p: { id: string }) => p.id === p1.id), 'kept in storage');
  assert.equal(saved.has(`nb.assemblyProposals.${b.fileHash}`), true);
  assemblyProposals.add(a.fileHash, [p2]);
  assert.deepEqual(assemblyProposals.list(a.fileHash).map((p) => p.id), [p2.id], 'the same assembly again replaces the pending one');
  assert.equal(a.assemblies().length, 1, 'the document is untouched');
});

test('corrupt storage is ignored, and a rejection stays and can be reviewed again', () => {
  const junk = `junk-${Math.random()}`;
  saved.set(`nb.assemblyProposals.${junk}`, '{not json');
  assert.deepEqual(assemblyProposals.list(junk), []);
  const half = `half-${Math.random()}`;
  const p = proposal();
  saved.set(`nb.assemblyProposals.${half}`, JSON.stringify([{ id: 7 }, { name: 'not a proposal' }, p]));
  assert.deepEqual(assemblyProposals.list(half).map((x) => x.name), ['Pipe run'], 'entries that are not proposals are dropped');
  assemblyProposals.reject(half, [p.id], 'Hana');
  assert.equal(assemblyProposals.list(half)[0]!.rejected!.by, 'Hana');
  assemblyProposals.reopen(half, [p.id]);
  assert.equal(assemblyProposals.list(half)[0]!.rejected, undefined);
});

test('approving writes the assemblies as one undo step, and the proposal leaves the list', async () => {
  const store = await setup();
  const p = proposal();
  assemblyProposals.add(store.fileHash, [p]);
  const { approved, skipped } = approveAssemblyProposals(store, [p]);
  assert.deepEqual([approved, skipped], [[p.id], []]);
  assert.deepEqual(assemblyProposals.list(store.fileHash), [], 'an approved proposal is written, not kept');
  const written = store.assemblies();
  assert.equal(written.length, 2);
  assert.equal(written[1]!.name, 'Pipe run');
  assert.deepEqual(written[1]!.items[0], { id: written[1]!.items[0]!.id, name: 'Pipe', unit: 'lf', formula: '[Length] * 1.05', waste: 5, roundUp: true, roundTo: 10, unitCost: 4.2 });
  store.undo();
  assert.deepEqual(store.assemblies().map((a) => a.id), ['wall'], 'one undo step takes the approval back');
});

test('an edit keeps the assembly id and the item ids its item names match', async () => {
  const store = await setup();
  const edit = proposal({
    name: 'Stud Wall',
    replacesId: 'wall',
    variables: [{ name: 'Spacing', value: 24 }],
    items: [
      { name: 'Studs', unit: 'ea', formula: '[Length] / [Spacing] + 1', waste: 5 },
      { name: 'Top track', unit: 'lf', formula: '[Length] * 2' },
    ],
  });
  assemblyProposals.add(store.fileHash, [edit]);
  approveAssemblyProposals(store, [edit]);
  const [replaced] = store.assemblies();
  assert.equal(replaced!.id, 'wall', 'the assembly id stays, so measurements stay linked');
  assert.equal(replaced!.variables[0]!.value, 24);
  assert.deepEqual(
    replaced!.items.map((i) => [i.name, i.id]),
    [
      ['Studs', 'studs'],
      ['Top track', replaced!.items[1]!.id],
    ],
    'a matching item name keeps its id; a new one gets a fresh id',
  );
  assert.notEqual(replaced!.items[1]!.id, 'board');
});

test('the My Assemblies box merges by name, without duplicates', async () => {
  const store = await setup();
  saved.set('nb.assemblyLibrary', JSON.stringify([{ id: 'lib1', name: 'pipe run', variables: [], items: [{ id: 'i', name: 'Pipe', unit: 'lf', formula: '[Length]' }] }]));
  const p = proposal();
  assemblyProposals.add(store.fileHash, [p]);
  approveAssemblyProposals(store, [p], true);
  const library = loadLibrary();
  assert.equal(library.length, 1, 'a saved assembly of the same name is replaced, not duplicated');
  assert.equal(library[0]!.id, 'lib1');
  assert.equal(library[0]!.items[0]!.formula, '[Length] * 1.05');
  // Without the box, the library is left alone.
  const p2 = proposal({ name: 'Slab', items: [{ name: 'Concrete', unit: 'm³', formula: '[Area] * 0.1' }] });
  assemblyProposals.add(store.fileHash, [p2]);
  approveAssemblyProposals(store, [p2], false);
  assert.deepEqual(loadLibrary().map((a) => [a.name, a.id]), [['Pipe run', 'lib1']], 'nothing more is written');
  writeLibrary([]);
});

test('a read-only document refuses the approval, and a proposal whose target is gone is skipped', async () => {
  const store = await setup();
  const p = proposal();
  assemblyProposals.add(store.fileHash, [p]);
  store.setReadOnly(true);
  assert.throws(() => approveAssemblyProposals(store, [p]), /read-only/);
  assert.deepEqual(assemblyProposals.list(store.fileHash).map((x) => x.id), [p.id], 'the proposal waits');
  assert.equal(store.assemblies().length, 1);
  store.setReadOnly(false);

  const gone = proposal({ replacesId: 'deleted' });
  const taken = proposal({ name: 'Stud Wall' });
  assemblyProposals.add(store.fileHash, [gone, taken]);
  const { approved, skipped } = approveAssemblyProposals(store, [gone, taken]);
  assert.deepEqual(approved, []);
  assert.deepEqual(skipped, [gone.id, taken.id], 'a replaced assembly that is gone, or a name another assembly took, is skipped');
  assert.equal(store.assemblies().length, 1);
  assert.deepEqual(assemblyProposals.list(store.fileHash).map((x) => x.id), [p.id, gone.id, taken.id], 'skipped proposals wait');
});

test('a proposal is checked again at approve time against what the document has since become', async () => {
  const store = await setup();
  const p = proposal({ variables: [{ name: 'Bay width', value: 3 }], items: [{ name: 'Boards', unit: 'ea', formula: '[Length] / [Bay width]' }] });
  assemblyProposals.add(store.fileHash, [p]);
  // A column named like the proposal's variable appeared after proposing: the column would win
  // at lookup, so the quantities would come out wrong.
  store.setColumnSet({ columns: [{ id: 'bay', name: 'Bay width', type: 'number' }], statuses: store.columnSet().statuses });
  const { approved, skipped } = approveAssemblyProposals(store, [p]);
  assert.deepEqual([approved, skipped], [[], [p.id]]);
  assert.deepEqual(assemblyProposals.list(store.fileHash).map((x) => x.id), [p.id], 'it waits');
  assert.equal(store.assemblies().length, 1);
});

test('a name taken at approve time is skipped, for edits and for two proposals that end with the same name', async () => {
  const both = await setup([wall, slab]);
  const rename = proposal({ name: 'Slab', replacesId: 'wall', variables: [], items: [{ name: 'Studs', unit: 'ea', formula: '[Length]' }] });
  assemblyProposals.add(both.fileHash, [rename]);
  const r = approveAssemblyProposals(both, [rename]);
  assert.deepEqual([r.approved, r.skipped], [[], [rename.id]], "an edit cannot take another assembly's name");
  assert.deepEqual(both.assemblies().map((a) => a.name), ['Stud Wall', 'Slab']);

  const seq = await setup();
  const edit = proposal({ name: 'Wall 2', replacesId: 'wall', variables: [], items: [{ name: 'Studs', unit: 'ea', formula: '[Length]' }] });
  const fresh = proposal({ name: 'Wall 2', variables: [], items: [{ name: 'Studs', unit: 'ea', formula: '[Length]' }] });
  assemblyProposals.add(seq.fileHash, [edit, fresh]);
  const s = approveAssemblyProposals(seq, [edit, fresh]);
  assert.deepEqual(s.approved, [edit.id]);
  assert.deepEqual(s.skipped, [fresh.id], 'the second one to arrive at the same name is skipped');
  assert.deepEqual(seq.assemblies().map((a) => [a.name, a.id]), [['Wall 2', 'wall']]);
});

test('an item keeps its id when the name differs only by case or spaces', async () => {
  const store = await setup();
  const edit = proposal({
    name: 'Stud Wall',
    replacesId: 'wall',
    variables: [],
    items: [
      { name: 'studs ', unit: 'ea', formula: '[Length] / 2' },
      { name: ' BOARD', unit: 'sheets', formula: '[Area]' },
    ],
  });
  assemblyProposals.add(store.fileHash, [edit]);
  approveAssemblyProposals(store, [edit]);
  const [a] = store.assemblies();
  assert.deepEqual(
    a!.items.map((i) => [i.name, i.id]),
    [
      ['studs', 'studs'],
      ['BOARD', 'board'],
    ],
    'the document trims the names; the ids still carried over by name',
  );
});

test('a rejected proposal for an assembly stays beside the new pending one', () => {
  const hash = `keep-${Math.random()}`;
  const p1 = proposal();
  assemblyProposals.add(hash, [p1]);
  assemblyProposals.reject(hash, [p1.id], 'Hana');
  const p2 = proposal({ items: [{ name: 'Pipe', unit: 'lf', formula: '[Length]' }] });
  assemblyProposals.add(hash, [p2]);
  assert.deepEqual(
    assemblyProposals.list(hash).map((p) => [p.id, !!p.rejected]),
    [
      [p1.id, true],
      [p2.id, false],
    ],
  );
});

test('when the browser will not keep My Assemblies, the document is still written and librarySaved is false', async () => {
  const store = await setup();
  const p = proposal();
  assemblyProposals.add(store.fileHash, [p]);
  const original = (globalThis as { localStorage: unknown }).localStorage;
  (globalThis as { localStorage: unknown }).localStorage = {
    getItem: (k: string) => saved.get(k) ?? null,
    setItem: () => {
      throw new Error('full');
    },
    removeItem: (k: string) => void saved.delete(k),
  };
  const outcome = approveAssemblyProposals(store, [p], true);
  (globalThis as { localStorage: unknown }).localStorage = original;
  assert.deepEqual(outcome.approved, [p.id]);
  assert.equal(outcome.librarySaved, false);
  assert.equal(store.assemblies().length, 2, 'the document was still written');
});

test('approving an edit with the box merges the edited assembly into My Assemblies by name', async () => {
  const store = await setup();
  saved.set('nb.assemblyLibrary', JSON.stringify([{ id: 'lib1', name: 'Stud Wall', variables: [], items: [{ id: 'li', name: 'Studs', unit: 'ea', formula: '[Length]' }] }]));
  const edit = proposal({ name: 'Stud Wall', replacesId: 'wall', variables: [], items: [{ name: 'Studs', unit: 'ea', formula: '[Length] / 2' }] });
  assemblyProposals.add(store.fileHash, [edit]);
  const { librarySaved } = approveAssemblyProposals(store, [edit], true);
  assert.equal(librarySaved, true);
  assert.deepEqual(loadLibrary().map((a) => [a.id, a.items[0]!.formula]), [['lib1', '[Length] / 2']], 'the library copy is replaced by name');
  writeLibrary([]);
});

test('approving the same proposal twice writes once', async () => {
  const store = await setup();
  const p = proposal();
  assemblyProposals.add(store.fileHash, [p]);
  const first = approveAssemblyProposals(store, [p]);
  assert.deepEqual(first.approved, [p.id]);
  const again = approveAssemblyProposals(store, [p], true);
  assert.deepEqual(again, { approved: [], skipped: [], librarySaved: false, libraryReplaced: [] }, 'a proposal already written is ignored');
  assert.equal(store.assemblies().length, 2);
  store.undo();
  assert.deepEqual(store.assemblies().map((a) => a.name), ['Stud Wall'], 'one undo step, so one write');
});

test('an edit keeps the existing notes unless it sends new ones', async () => {
  const store = await setup();
  const items = [{ name: 'Studs', unit: 'ea', formula: '[Length] / [Spacing] + 1' }];
  const keep = proposal({ name: 'Stud Wall', replacesId: 'wall', variables: [{ name: 'Spacing', value: 16 }], items });
  assemblyProposals.add(store.fileHash, [keep]);
  approveAssemblyProposals(store, [keep]);
  assert.equal(store.assemblies()[0]!.description, 'Partition wall', 'no description sent keeps the notes');
  const clear = proposal({ name: 'Stud Wall', replacesId: 'wall', variables: [{ name: 'Spacing', value: 16 }], items, description: '' });
  assemblyProposals.add(store.fileHash, [clear]);
  approveAssemblyProposals(store, [clear]);
  assert.equal(store.assemblies()[0]!.description, undefined, 'an empty one clears them');
  const change = proposal({ name: 'Stud Wall', replacesId: 'wall', variables: [{ name: 'Spacing', value: 16 }], items, description: 'Partition wall, boarded one side' });
  assemblyProposals.add(store.fileHash, [change]);
  approveAssemblyProposals(store, [change]);
  assert.equal(store.assemblies()[0]!.description, 'Partition wall, boarded one side');
});

test('at most 200 proposals are kept for a document, the oldest rejected going first', () => {
  const hash = `cap-${Math.random()}`;
  const rejected = [proposal({ id: 'r1', name: 'Old 1' }), proposal({ id: 'r2', name: 'Old 2' })];
  assemblyProposals.add(hash, rejected);
  assemblyProposals.reject(hash, rejected.map((p) => p.id), 'Hana');
  assemblyProposals.add(hash, Array.from({ length: 198 }, (_, i) => proposal({ id: `p${i}`, name: `Run ${i}` })));
  assert.equal(assemblyProposals.list(hash).length, 200);
  // Two more would be 202: the two rejected ones are dropped to make room.
  assemblyProposals.add(hash, [proposal({ id: 'new1', name: 'New 1' }), proposal({ id: 'new2', name: 'New 2' })]);
  const list = assemblyProposals.list(hash);
  assert.equal(list.length, 200);
  assert.ok(!list.some((p) => p.rejected), 'the rejected ones were dropped first');
  assert.ok(list.some((p) => p.id === 'new2'));
  // With nothing rejected left to drop, the call is refused and nothing changes.
  assert.throws(() => assemblyProposals.add(hash, [proposal({ id: 'new3', name: 'New 3' })]), /Review or reject some first/);
  assert.equal(assemblyProposals.list(hash).length, 200);
});

test('the library merge is by name, so a rename edit cannot leave two of one name, and a replace is reported', async () => {
  const store = await setup();
  // The library holds this assembly under its old name, and another entry already has the new one.
  saved.set('nb.assemblyLibrary', JSON.stringify([
    { id: 'wall', name: 'Stud Wall', variables: [], items: [{ id: 's', name: 'Studs', unit: 'ea', formula: '[Length]' }] },
    { id: 'lib1', name: 'wall 2', variables: [], items: [{ id: 'j', name: 'Track', unit: 'lf', formula: '[Length]' }] },
  ]));
  const rename = proposal({ name: 'Wall 2', replacesId: 'wall', variables: [], items: [{ name: 'Studs', unit: 'ea', formula: '[Length] / 2' }] });
  assemblyProposals.add(store.fileHash, [rename]);
  const out = approveAssemblyProposals(store, [rename], true);
  const library = loadLibrary();
  const ofNewName = library.filter((a) => a.name.trim().toLowerCase() === 'wall 2');
  assert.equal(ofNewName.length, 1, 'no two entries of one name');
  assert.deepEqual([ofNewName[0]!.id, ofNewName[0]!.items[0]!.formula], ['lib1', '[Length] / 2'], "the library entry's own id, the proposal's content");
  assert.deepEqual(out.libraryReplaced, ['wall 2'], 'the overwritten library entry, by its own name');
  writeLibrary([]);
});

test('a library entry of the same name but different content is replaced, and said so', async () => {
  const store = await setup();
  saved.set('nb.assemblyLibrary', JSON.stringify([{ id: 'lib2', name: 'Pipe run', variables: [], items: [{ id: 'j', name: 'Pipe', unit: 'lf', formula: '[Length]' }] }]));
  const fresh = proposal({ name: 'Pipe run', variables: [], items: [{ name: 'Pipe', unit: 'lf', formula: '[Length] * 3' }] });
  assemblyProposals.add(store.fileHash, [fresh]);
  assert.deepEqual(approveAssemblyProposals(store, [fresh], true).libraryReplaced, ['Pipe run']);
  assert.deepEqual(loadLibrary().map((a) => [a.id, a.items[0]!.formula]), [['lib2', '[Length] * 3']]);
  writeLibrary([]);
});

test('corrupt entries that would crash the review card are dropped', () => {
  const hash = `strict-${Math.random()}`;
  const good = proposal();
  const junk = [
    { ...good, id: 'a', items: [{ name: 'No unit or formula' }] },
    { ...good, id: 'b', variables: [{ name: 'V', value: 'six' }] },
    { ...good, id: 'c', description: 7 },
    good,
  ];
  saved.set(`nb.assemblyProposals.${hash}`, JSON.stringify(junk));
  assert.deepEqual(assemblyProposals.list(hash).map((p) => p.id), [good.id]);
});

test('a proposal that starts waiting while others were decided is still noticed', () => {
  const before = new Set(['t1', 't2', 'a1']);
  const after = new Set(['t2', 'a1', 'a2']);
  assert.deepEqual(newlyWaiting(before, after), ['a2'], 'ids, not counts: an approval removed t1 in the same moment');
  assert.deepEqual(newlyWaiting(before, new Set(['t1'])), []);
});
