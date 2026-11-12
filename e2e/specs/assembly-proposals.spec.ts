import { expect, test, type App } from '../support/app.ts';

// Assembly proposals: an AI client proposes a takeoff assembly through the assistant
// (propose_assembly), and it waits in Tools › Takeoff › Review AI Proposals until a person approves
// or rejects it. Same pattern as the takeoff-proposal test in measure.spec.ts.

interface Item {
  id: string;
  name: string;
  unit: string;
  formula: string;
  roundUp?: boolean;
}
interface Asm {
  id: string;
  name: string;
  variables: { name: string; value: number }[];
  items: Item[];
}
type Store = { assemblies(): Asm[]; setAssemblies(a: unknown[]): void };

/** The active document's assemblies, straight from the store. */
const assemblies = (app: App) => app.page.evaluate(() => JSON.parse(JSON.stringify((window.__nb.tools as unknown as { store: Store }).store.assemblies())) as Asm[]);
const setAssemblies = (app: App, list: Asm[]) => app.page.evaluate((l) => (window.__nb.tools as unknown as { store: Store }).store.setAssemblies(l), list);
/** My Assemblies, as this browser keeps it. */
const library = (app: App) => app.page.evaluate(() => JSON.parse(localStorage.getItem('nb.assemblyLibrary') ?? '[]') as Asm[]);
const proposalStates = async (app: App) => (await app.assistant<{ proposals: { name: string; state: string }[] }>('list_assembly_proposals')).proposals.map((p) => [p.name, p.state]);

const PIPE_RUN = {
  name: 'Pipe run',
  reasoning: 'Acme Co hangs its pipe every 8 ft, with one hanger at each end.',
  variables: [{ name: 'Spacing', value: 8 }],
  items: [
    { name: 'Pipe', unit: 'ft', formula: '[Length]' },
    { name: 'Hangers', unit: 'ea', formula: '[Length] / [Spacing] + 1', roundUp: true },
  ],
};

test.describe('assembly proposals', () => {
  test.beforeEach(async ({ app }) => {
    await app.open('plan.pdf');
  });

  test('a proposed assembly waits for review with its items and reasoning; the document is unchanged', async ({ app }) => {
    const { page } = app;
    const made = await app.assistant<{ proposed: number }>('propose_assembly', { source: 'Claude', assemblies: [PIPE_RUN] });
    expect(made.proposed).toBe(1);
    // The review opens by itself.
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    await expect(review).toBeVisible();
    const section = review.getByRole('region', { name: 'Assembly proposals' });
    const card = section.getByRole('article', { name: 'Assembly proposal for Pipe run' });
    await expect(card).toHaveCount(1);
    await expect(card.locator('.tp-head')).toContainText('New assembly');
    await expect(card.locator('.tp-state')).toHaveText('Waiting');
    await expect(card.locator('.tp-vars')).toHaveText(/^Variables: Spacing = 8(\.0+)?$/);
    const rows = card.locator('tbody tr');
    await expect(rows).toHaveCount(2);
    await expect(rows.nth(0).locator('td').nth(0)).toHaveText('Pipe ft');
    await expect(rows.nth(0).locator('code')).toHaveText('[Length]');
    await expect(rows.nth(1).locator('td').nth(0)).toHaveText('Hangers ea');
    await expect(rows.nth(1).locator('code')).toHaveText('[Length] / [Spacing] + 1');
    await expect(rows.nth(1).locator('td').nth(3)).toHaveText('up');
    await expect(card.locator('.tp-reason')).toContainText('Claude');
    await expect(card.locator('.tp-reason')).toContainText(PIPE_RUN.reasoning);
    await expect(review.locator('.takeoff-review-head')).toContainText('1 waiting');
    expect(await proposalStates(app)).toEqual([['Pipe run', 'pending']]);
    // Nothing reaches the document until a person approves.
    expect(await assemblies(app)).toEqual([]);
  });

  test('Approve adds the assembly to the document; Edit › Undo takes it out again', async ({ app }) => {
    const { page } = app;
    await app.assistant('propose_assembly', { source: 'Claude', assemblies: [PIPE_RUN] });
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    const card = review.getByRole('article', { name: 'Assembly proposal for Pipe run' });
    await card.getByRole('button', { name: 'Approve', exact: true }).click();
    await expect.poll(async () => (await assemblies(app)).map((a) => a.name)).toEqual(['Pipe run']);
    await expect(app.notice()).toContainText('Approved 1 assembly proposal');
    await expect(card).toHaveCount(0);
    expect(await proposalStates(app)).toEqual([]);
    const [added] = await assemblies(app);
    expect(added!.variables).toEqual([{ name: 'Spacing', value: 8 }]);
    expect(added!.items.map((i) => [i.name, i.unit, i.formula, !!i.roundUp])).toEqual([
      ['Pipe', 'ft', '[Length]', false],
      ['Hangers', 'ea', '[Length] / [Spacing] + 1', true],
    ]);
    // The Assemblies window lists it.
    await review.getByRole('button', { name: 'Close takeoff review' }).click();
    await app.cmd('tools.assemblies');
    const dialog = page.getByRole('dialog', { name: 'Assemblies' });
    const list = dialog.getByRole('list', { name: 'Assemblies in this document' });
    await expect(list.locator('.col-name')).toHaveText(['Pipe run']);
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(dialog).toBeHidden();
    // Approving was one undo step.
    await app.cmd('edit.undo');
    await expect.poll(async () => (await assemblies(app)).map((a) => a.name)).toEqual([]);
    await app.cmd('tools.assemblies');
    await expect(list.locator('.col-name')).toHaveCount(0);
    await expect(list).toContainText('No assemblies yet');
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  });

  test('Reject takes it off the waiting list; it shows as rejected under all, and Review again brings it back', async ({ app }) => {
    const { page } = app;
    await app.assistant('propose_assembly', { source: 'Claude', assemblies: [PIPE_RUN] });
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    const card = review.getByRole('article', { name: 'Assembly proposal for Pipe run' });
    await expect(card).toHaveCount(1);
    await card.getByRole('button', { name: 'Reject', exact: true }).click();
    await expect(card).toHaveCount(0);
    await expect(review.locator('.takeoff-review-list .empty')).toHaveText('Nothing waiting for review.');
    await expect(review.locator('.takeoff-review-head')).toContainText('0 waiting');
    expect(await proposalStates(app)).toEqual([['Pipe run', 'rejected']]);
    // Shown with the decisions: rejected, by whom, and a way back.
    await review.getByLabel('Show').selectOption('all');
    await expect(card).toHaveCount(1);
    await expect(card.locator('.tp-state')).toHaveText('Rejected');
    await expect(card).toContainText('Rejected by E2E Tester');
    await expect(card.getByRole('button', { name: 'Approve', exact: true })).toHaveCount(0);
    await card.getByRole('button', { name: 'Review again' }).click();
    await expect(card.locator('.tp-state')).toHaveText('Waiting');
    await review.getByLabel('Show').selectOption('pending');
    await expect(card).toHaveCount(1);
    await expect(card.getByRole('button', { name: 'Approve', exact: true })).toBeVisible();
    expect(await proposalStates(app)).toEqual([['Pipe run', 'pending']]);
    expect(await assemblies(app)).toEqual([]);
  });

  test('an edit (replaces) shows what it adds, removes and changes; Approve keeps the assembly and its item ids', async ({ app }) => {
    const { page } = app;
    await setAssemblies(app, [
      {
        id: 'pipe',
        name: 'Pipe run',
        variables: [{ name: 'Spacing', value: 8 }],
        items: [
          { id: 'pipe-ft', name: 'Pipe', unit: 'ft', formula: '[Length]' },
          { id: 'hangers', name: 'Hangers', unit: 'ea', formula: '[Length] / [Spacing] + 1', roundUp: true },
          { id: 'paint', name: 'Paint', unit: 'gal', formula: '[Length] / 100' },
        ],
      },
    ]);
    await app.assistant('propose_assembly', {
      source: 'Claude',
      assemblies: [
        {
          name: 'Pipe run',
          replaces: 'Pipe run',
          reasoning: 'Acme Co insulates its pipe and does not paint it; hangers every 6 ft.',
          variables: [{ name: 'Spacing', value: 6 }],
          items: [
            { name: 'Pipe', unit: 'ft', formula: '[Length]' },
            { name: 'Hangers', unit: 'ea', formula: '[Length] / [Spacing] + 2', roundUp: true },
            { name: 'Insulation', unit: 'ft', formula: '[Length]' },
          ],
        },
      ],
    });
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    const card = review.getByRole('article', { name: 'Assembly proposal for Pipe run' });
    await expect(card.locator('.tp-head')).toContainText('Replaces Pipe run');
    await expect(card).toContainText('added: Insulation · removed: Paint · changed: Hangers');
    await expect(card.locator('.tp-vars')).toHaveText(/^Variables: Spacing = 6(\.0+)?$/);
    expect((await assemblies(app))[0]!.items.map((i) => i.name)).toEqual(['Pipe', 'Hangers', 'Paint']);
    await card.getByRole('button', { name: 'Approve', exact: true }).click();
    await expect.poll(async () => (await assemblies(app))[0]!.items.map((i) => i.name)).toEqual(['Pipe', 'Hangers', 'Insulation']);
    const after = await assemblies(app);
    expect(after).toHaveLength(1);
    expect(after[0]!.id).toBe('pipe');
    expect(after[0]!.variables).toEqual([{ name: 'Spacing', value: 6 }]);
    // Items with the same name keep their ids, so measurements keep their per-item overrides.
    expect(after[0]!.items.map((i) => [i.id, i.formula])).toEqual([
      ['pipe-ft', '[Length]'],
      ['hangers', '[Length] / [Spacing] + 2'],
      [expect.any(String), '[Length]'],
    ]);
    expect(['pipe-ft', 'hangers', 'paint']).not.toContain(after[0]!.items[2]!.id);
  });

  test('Also save to My Assemblies keeps a copy in this browser, merged by name', async ({ app }) => {
    const { page } = app;
    expect(await library(app)).toEqual([]);
    await app.assistant('propose_assembly', { source: 'Claude', assemblies: [PIPE_RUN] });
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    const card = review.getByRole('article', { name: 'Assembly proposal for Pipe run' });
    await card.getByLabel('Also save to My Assemblies').check();
    await card.getByRole('button', { name: 'Approve', exact: true }).click();
    await expect(app.notice()).toContainText('saved to My Assemblies');
    await expect.poll(async () => (await library(app)).map((a) => a.name)).toEqual(['Pipe run']);
    const [first] = await library(app);
    expect(first!.items.map((i) => i.formula)).toEqual(['[Length]', '[Length] / [Spacing] + 1']);

    // Take it out of the document and approve a revised one under the same name: the library
    // keeps one Pipe run (the first copy's id), with the new items.
    await app.cmd('edit.undo');
    await expect.poll(async () => (await assemblies(app)).length).toBe(0);
    const revised = { ...PIPE_RUN, items: [PIPE_RUN.items[0]!, { ...PIPE_RUN.items[1]!, formula: '[Length] / [Spacing] + 2' }] };
    await app.assistant('propose_assembly', { source: 'Claude', assemblies: [revised] });
    await expect(card).toHaveCount(1);
    await card.getByLabel('Also save to My Assemblies').check();
    await card.getByRole('button', { name: 'Approve', exact: true }).click();
    await expect.poll(async () => (await assemblies(app)).map((a) => a.name)).toEqual(['Pipe run']);
    await expect.poll(async () => (await library(app)).map((a) => [a.id, a.items.map((i) => i.formula)])).toEqual([[first!.id, ['[Length]', '[Length] / [Spacing] + 2']]]);
  });

  test('a variable named like a list column (Slope) is refused, all or nothing, and nothing waits for review', async ({ app }) => {
    const { page } = app;
    const sloped = { ...PIPE_RUN, name: 'Roof drain', variables: [{ name: 'Slope', value: 2 }], items: [{ name: 'Drain pipe', unit: 'ft', formula: '[Length] * [Slope]' }] };
    await expect(app.assistant('propose_assembly', { source: 'Claude', assemblies: [PIPE_RUN, sloped] })).rejects.toThrow(/Roof drain.*Slope.*already a measured value or column/);
    expect(await proposalStates(app)).toEqual([]);
    expect(await assemblies(app)).toEqual([]);
    const review = page.getByRole('complementary', { name: 'Takeoff review' });
    await expect(review).toBeHidden();
    await app.cmd('tools.takeoffReview');
    await expect(review).toBeVisible();
    await expect(review.getByRole('region', { name: 'Assembly proposals' })).toHaveCount(0);
    await expect(review.locator('.takeoff-review-list .empty')).toContainText('No proposals yet');
  });

  test.afterEach(async ({ app }) => {
    expect(app.errors).toEqual([]);
  });
});
