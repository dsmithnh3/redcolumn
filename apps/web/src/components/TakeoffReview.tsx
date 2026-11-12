import { useState } from 'react';
import { MARKUP_LABELS, type Assembly, type Markup, type MarkupStore } from '@nb/markup';
import { itemQuantities } from '../assemblies/quantities';
import { approveProposals, stateOf, takeoffProposals, useTakeoffProposals, type ApprovalEdit, type TakeoffProposal } from '../assemblies/proposals';
import { approveAssemblyProposals, assemblyProposals, useAssemblyProposals, type AssemblyProposal } from '../assemblies/assemblyProposals';
import { cellsFor, formatNumber, type CellContext } from '../columns/listColumns';

interface Props {
  store: MarkupStore;
  /** Who approves: the name markups are authored under. */
  me: string;
  ctx: CellContext;
  pageLabel: (pageIndex: number) => string;
  /** Selects the measurement on the sheet and brings it into view. */
  onShow: (m: Markup) => void;
  onMessage: (text: string) => void;
  onClose: () => void;
}

const n = (v: number | null | undefined) => (v === null || v === undefined ? '—' : formatNumber(v, 2));
const when = (at: number) => new Date(at).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });

/**
 * Tools › Takeoff › Review AI Proposals: each assembly an AI client proposed for a measurement,
 * with its formula, variables, quantities and reasoning, to approve, edit or reject. Nothing is
 * approved except by a person clicking Approve; approving several at once is its own button.
 * A panel beside the sheet (not a dialog), so the measurement stays in view.
 */
export function TakeoffReview({ store, me, ctx, pageLabel, onShow, onMessage, onClose }: Props) {
  const proposals = useTakeoffProposals(store.fileHash);
  const assemblies = store.assemblies();
  const [show, setShow] = useState<'pending' | 'all'>('pending');
  const [picked, setPicked] = useState<Set<string>>(new Set());
  const [edits, setEdits] = useState<Record<string, ApprovalEdit>>({});
  const rows = proposals.map((p) => ({ p, m: store.get(p.markupId), state: stateOf(p, store.get(p.markupId)) }));
  const shown = rows.filter((r) => show === 'all' || r.state === 'pending');
  const pending = rows.filter((r) => r.state === 'pending');
  const pickedPending = pending.filter((r) => picked.has(r.p.id));
  const assemblyRows = useAssemblyProposals(store.fileHash).map((p) => ({ p, state: p.rejected ? ('rejected' as const) : ('pending' as const) }));
  const shownAssemblies = assemblyRows.filter((r) => show === 'all' || r.state === 'pending');
  const pendingAssemblies = assemblyRows.filter((r) => r.state === 'pending');
  const readOnly = store.readOnly;

  const approve = (list: TakeoffProposal[]) => {
    const { approved, skipped } = approveProposals(store, list, me, edits);
    setPicked((s) => new Set([...s].filter((id) => !approved.includes(id))));
    onMessage(`Approved ${approved.length} takeoff proposal${approved.length === 1 ? '' : 's'}${skipped.length ? `; ${skipped.length} could not be (the measurement is gone or locked)` : ''}. Undo takes it back.`);
  };
  const reject = (list: TakeoffProposal[]) => {
    takeoffProposals.reject(store.fileHash, list.map((p) => p.id), me);
    setPicked((s) => new Set([...s].filter((id) => !list.some((p) => p.id === id))));
  };
  const edit = (id: string, patch: (e: ApprovalEdit) => ApprovalEdit) => setEdits((all) => ({ ...all, [id]: patch(all[id] ?? {}) }));
  const approveAssembly = (list: AssemblyProposal[], alsoLibrary: boolean) => {
    try {
      const { approved, skipped, librarySaved } = approveAssemblyProposals(store, list, alsoLibrary);
      const parts = [`Approved ${approved.length} assembly proposal${approved.length === 1 ? '' : 's'}`];
      if (alsoLibrary && approved.length) parts.push(librarySaved ? 'saved to My Assemblies' : 'but My Assemblies could not be saved (this browser would not keep them: storage full or unavailable)');
      if (skipped.length) parts.push(`${skipped.length} could not be (the assembly it replaces is gone, the name is taken, or a column now clashes with a variable)`);
      onMessage(`${parts.join('; ')}. Undo takes it back.`);
    } catch (err) {
      onMessage(err instanceof Error ? err.message : String(err));
    }
  };
  const rejectAssembly = (list: AssemblyProposal[]) => assemblyProposals.reject(store.fileHash, list.map((p) => p.id), me);

  return (
    <aside className="takeoff-review" role="complementary" aria-label="Takeoff review">
      <header className="takeoff-review-head">
        <h3>Review AI Proposals</h3>
        <span className="hint-text">{pending.length + pendingAssemblies.length} waiting</span>
        <span className="grow" />
        <select aria-label="Show" value={show} onChange={(e) => setShow(e.target.value as 'pending' | 'all')}>
          <option value="pending">Waiting for review</option>
          <option value="all">All, with decisions</option>
        </select>
        <button type="button" className="link-preview-btn" aria-label="Close takeoff review" onClick={onClose}>
          ×
        </button>
      </header>
      <p className="hint-text takeoff-review-note">An AI client proposed these through the local assistant. Check each one against the drawing: nothing is added to the takeoff until you approve it.</p>
      <div className="takeoff-review-list">
        {!shown.length && !shownAssemblies.length && <p className="empty">{proposals.length || assemblyRows.length ? 'Nothing waiting for review.' : 'No proposals yet. Ask your AI client (connected to redcolumn) to propose a takeoff or an assembly.'}</p>}
        {!!shownAssemblies.length && (
          <section className="ap-section" aria-label="Assembly proposals">
            <h4>Assembly proposals</h4>
            <p className="hint-text ap-note">Assemblies an AI client proposed for this document. Approving writes them into its assemblies; nothing is added until you approve it.</p>
            {shownAssemblies.map(({ p, state }) => (
              <AssemblyCard
                key={p.id}
                p={p}
                state={state}
                assemblies={assemblies}
                readOnly={readOnly}
                onApprove={(alsoLibrary) => approveAssembly([p], alsoLibrary)}
                onReject={() => rejectAssembly([p])}
                onReopen={() => assemblyProposals.reopen(store.fileHash, [p.id])}
              />
            ))}
          </section>
        )}
        {shown.map(({ p, m, state }) => (
          <Proposal
            key={p.id}
            p={p}
            m={m}
            state={state}
            assemblies={assemblies}
            edit={edits[p.id] ?? {}}
            picked={picked.has(p.id)}
            readOnly={readOnly}
            ctx={ctx}
            pageLabel={pageLabel}
            onPick={(on) => setPicked((s) => (on ? new Set([...s, p.id]) : new Set([...s].filter((x) => x !== p.id))))}
            onEdit={(patch) => edit(p.id, patch)}
            onApprove={() => approve([p])}
            onReject={() => reject([p])}
            onReopen={() => takeoffProposals.reopen(store.fileHash, [p.id])}
            onShow={() => m && onShow(m)}
          />
        ))}
      </div>
      <footer className="takeoff-review-foot">
        <label className="tp-select-all">
          <input type="checkbox" checked={!!pending.length && pickedPending.length === pending.length} onChange={(e) => setPicked(e.target.checked ? new Set(pending.map((r) => r.p.id)) : new Set())} disabled={!pending.length} />
          <span>Select all waiting</span>
        </label>
        <span className="grow" />
        <button type="button" className="btn small" disabled={!pickedPending.length} onClick={() => reject(pickedPending.map((r) => r.p))}>
          Reject Selected ({pickedPending.length})
        </button>
        <button type="button" className="btn small primary" disabled={readOnly || !pickedPending.length} onClick={() => approve(pickedPending.map((r) => r.p))}>
          Approve Selected ({pickedPending.length})
        </button>
      </footer>
    </aside>
  );
}

function Proposal({
  p,
  m,
  state,
  assemblies,
  edit,
  picked,
  readOnly,
  ctx,
  pageLabel,
  onPick,
  onEdit,
  onApprove,
  onReject,
  onReopen,
  onShow,
}: {
  p: TakeoffProposal;
  m: Markup | undefined;
  state: 'pending' | 'approved' | 'rejected';
  assemblies: readonly Assembly[];
  edit: ApprovalEdit;
  picked: boolean;
  readOnly: boolean;
  ctx: CellContext;
  pageLabel: (pageIndex: number) => string;
  onPick: (on: boolean) => void;
  onEdit: (patch: (e: ApprovalEdit) => ApprovalEdit) => void;
  onApprove: () => void;
  onReject: () => void;
  onReopen: () => void;
  onShow: () => void;
}) {
  const assemblyId = state === 'approved' ? (m?.assemblyId ?? p.assemblyId) : (edit.assemblyId ?? p.assemblyId);
  const assembly = assemblies.find((a) => a.id === assemblyId);
  const proposed = assemblies.find((a) => a.id === p.assemblyId);
  const cells = m ? cellsFor(m, ctx) : null;
  const overrides = state === 'approved' ? m?.assemblyOverrides : edit.overrides;
  const items = m && assembly && cells ? itemQuantities({ ...m, assemblyOverrides: overrides }, assembly, cells, ctx.columns) : [];
  const expectedOf = (name: string) => p.expected?.find((e) => e.item.trim().toLowerCase() === name.trim().toLowerCase());
  const pending = state === 'pending';

  return (
    <article className={`takeoff-proposal ${state}`} aria-label={`Proposal for ${m ? (m.subject ?? MARKUP_LABELS[m.type]) : 'a deleted measurement'}`}>
      <div className="tp-head">
        {pending && <input type="checkbox" aria-label="Select proposal" checked={picked} onChange={(e) => onPick(e.target.checked)} />}
        <button type="button" className="tp-measure" title="Show it on the sheet" disabled={!m} onClick={onShow}>
          {m ? `${m.subject ?? MARKUP_LABELS[m.type]} · ${pageLabel(m.pageIndex)}` : 'Measurement deleted'}
        </button>
        <span className="tp-value">{cells?.measurement?.text ?? ''}</span>
        <span className="grow" />
        <span className={`tp-state ${state}`}>{state === 'pending' ? 'Waiting' : state === 'approved' ? 'Approved' : 'Rejected'}</span>
      </div>
      <div className="tp-assembly">
        <span className="hint-text">Assembly</span>
        {pending ? (
          <select aria-label="Assembly" value={assemblyId} onChange={(e) => onEdit((x) => ({ ...x, assemblyId: e.target.value === p.assemblyId ? undefined : e.target.value, overrides: undefined }))}>
            {assemblies.map((a) => (
              <option key={a.id} value={a.id}>
                {a.name}
                {a.id === p.assemblyId ? ' (proposed)' : ''}
              </option>
            ))}
          </select>
        ) : (
          <b>{assembly?.name ?? 'deleted assembly'}</b>
        )}
        {assemblyId !== p.assemblyId && <span className="warn">changed from {proposed?.name ?? 'the proposal'}</span>}
      </div>
      {assembly && assembly.variables.length > 0 && <div className="hint-text tp-vars">Variables: {assembly.variables.map((v) => `${v.name} = ${formatNumber(v.value, 4)}`).join(', ')}</div>}
      {items.length > 0 && (
        <table className="qty-table tp-items">
          <thead>
            <tr>
              <th>Item</th>
              <th>Formula</th>
              <th className="num">Quantity</th>
              <th className="num">With waste</th>
              <th className="num">AI expected</th>
            </tr>
          </thead>
          <tbody>
            {items.map((q) => {
              const exp = expectedOf(q.item.name);
              const differs = exp && q.rounded !== null && Math.abs(exp.quantity - q.rounded) > Math.max(0.01, Math.abs(q.rounded) * 0.01);
              const set = overrides?.[q.item.id];
              return (
                <tr key={q.item.id}>
                  <td>
                    {q.item.name} <span className="hint-text">{q.item.unit}</span>
                  </td>
                  <td>
                    <code>{q.item.formula}</code>
                    {q.error && <span className="warn"> {q.error}</span>}
                  </td>
                  <td className="num">
                    {pending ? (
                      <input
                        type="number"
                        step="any"
                        aria-label={`${q.item.name} quantity`}
                        className={set !== undefined ? 'tp-set' : undefined}
                        title={set !== undefined ? 'Set by hand (clear it to use the formula)' : 'From the formula; type to set it by hand'}
                        value={set ?? (q.base !== null ? Number(q.base.toFixed(4)) : '')}
                        onChange={(e) => {
                          const v = e.target.value === '' ? undefined : Number(e.target.value);
                          onEdit((x) => {
                            const next = { ...x.overrides };
                            if (v === undefined || !Number.isFinite(v)) delete next[q.item.id];
                            else next[q.item.id] = v;
                            return { ...x, overrides: next };
                          });
                        }}
                      />
                    ) : (
                      <>
                        {n(q.base)}
                        {set !== undefined && <span className="hint-text"> (set)</span>}
                      </>
                    )}
                  </td>
                  <td className="num">
                    <b>{n(q.rounded)}</b>
                  </td>
                  <td className={`num${differs ? ' warn' : ''}`} title={differs ? 'Differs from what the assembly works out' : undefined}>
                    {exp ? `${formatNumber(exp.quantity, 2)}${differs ? ' ⚠' : ''}` : '—'}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      <blockquote className="tp-reason">
        <span className="hint-text">
          {p.source}, {when(p.createdAt)}:
        </span>{' '}
        {p.reasoning}
      </blockquote>
      {state === 'approved' && m?.takeoffApproval && (
        <p className="hint-text">
          Approved by {m.takeoffApproval.by}, {when(m.takeoffApproval.at)}
          {m.takeoffApproval.edited ? ' (edited before approving)' : ''}.
        </p>
      )}
      {state === 'rejected' && p.rejected && (
        <p className="hint-text">
          Rejected by {p.rejected.by}, {when(p.rejected.at)}.{' '}
          <button type="button" className="rc-linkish" onClick={onReopen}>
            Review again
          </button>
        </p>
      )}
      {pending && (
        <div className="tp-actions">
          <button type="button" className="btn small" onClick={onReject}>
            Reject
          </button>
          <button type="button" className="btn small primary" disabled={readOnly || !m || !assembly} onClick={onApprove}>
            Approve
          </button>
        </div>
      )}
    </article>
  );
}

/** What an edit changes in the assembly's items, as names: added, removed and changed. */
function itemChanges(p: AssemblyProposal, existing: Assembly | undefined): { added: string[]; removed: string[]; changed: string[] } {
  const key = (name: string) => name.trim().toLowerCase();
  const oldItems = new Map(existing?.items.map((i) => [key(i.name), i]));
  const newNames = new Set(p.items.map((i) => key(i.name)));
  const added = p.items.filter((i) => !oldItems.has(key(i.name))).map((i) => i.name);
  const removed = existing?.items.filter((i) => !newNames.has(key(i.name))).map((i) => i.name) ?? [];
  const changed = p.items
    .filter((i) => {
      const old = oldItems.get(key(i.name));
      return old && (old.unit !== i.unit || old.formula !== i.formula || old.waste !== i.waste || old.roundUp !== i.roundUp || old.roundTo !== i.roundTo || old.unitCost !== i.unitCost);
    })
    .map((i) => i.name);
  return { added, removed, changed };
}

/** One proposed assembly: its items and variables, what an edit changes, and the AI's reasoning. */
function AssemblyCard({
  p,
  state,
  assemblies,
  readOnly,
  onApprove,
  onReject,
  onReopen,
}: {
  p: AssemblyProposal;
  state: 'pending' | 'rejected';
  assemblies: readonly Assembly[];
  readOnly: boolean;
  onApprove: (alsoLibrary: boolean) => void;
  onReject: () => void;
  onReopen: () => void;
}) {
  const [alsoLibrary, setAlsoLibrary] = useState(false);
  const existing = p.replacesId ? assemblies.find((a) => a.id === p.replacesId) : undefined;
  const pending = state === 'pending';
  const { added, removed, changed } = itemChanges(p, existing);
  const cost = (v: number | undefined) => (v === undefined ? '—' : formatNumber(v, 2));

  return (
    <article className={`takeoff-proposal ${state}`} aria-label={`Assembly proposal for ${p.name}`}>
      <div className="tp-head">
        <span className="ap-name">{p.name}</span>
        <span className="hint-text">{existing ? `Replaces ${existing.name}` : p.replacesId ? 'Replaces a deleted assembly' : 'New assembly'}</span>
        <span className="grow" />
        <span className={`tp-state ${state}`}>{pending ? 'Waiting' : 'Rejected'}</span>
      </div>
      {p.description && <p className="hint-text ap-note">{p.description}</p>}
      {existing && (
        <p className="hint-text ap-note">
          {[added.length ? `added: ${added.join(', ')}` : '', removed.length ? `removed: ${removed.join(', ')}` : '', changed.length ? `changed: ${changed.join(', ')}` : ''].filter(Boolean).join(' · ') || 'no changes to items'}
        </p>
      )}
      {p.variables.length > 0 && <div className="hint-text tp-vars">Variables: {p.variables.map((v) => `${v.name} = ${formatNumber(v.value, 4)}`).join(', ')}</div>}
      <table className="qty-table tp-items">
        <thead>
          <tr>
            <th>Item</th>
            <th>Formula</th>
            <th className="num">Waste</th>
            <th>Rounding</th>
            <th className="num">Unit cost</th>
          </tr>
        </thead>
        <tbody>
          {p.items.map((i) => (
            <tr key={i.name}>
              <td>
                {i.name} <span className="hint-text">{i.unit}</span>
              </td>
              <td>
                <code>{i.formula}</code>
              </td>
              <td className="num">{i.waste ? `${formatNumber(i.waste, 1)}%` : '—'}</td>
              <td>{i.roundUp ? `up${i.roundTo && i.roundTo !== 1 ? ` to ${formatNumber(i.roundTo, 2)}` : ''}` : '—'}</td>
              <td className="num">{cost(i.unitCost)}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <blockquote className="tp-reason">
        <span className="hint-text">
          {p.source}, {when(p.createdAt)}:
        </span>{' '}
        {p.reasoning}
      </blockquote>
      {state === 'rejected' && p.rejected && (
        <p className="hint-text">
          Rejected by {p.rejected.by}, {when(p.rejected.at)}.{' '}
          <button type="button" className="rc-linkish" onClick={onReopen}>
            Review again
          </button>
        </p>
      )}
      {pending && (
        <div className="tp-actions">
          <label className="tp-select-all" title="Keep a copy in this browser's My Assemblies, for use in any document">
            <input type="checkbox" checked={alsoLibrary} onChange={(e) => setAlsoLibrary(e.target.checked)} disabled={readOnly} />
            <span>Also save to My Assemblies</span>
          </label>
          <span className="grow" />
          <button type="button" className="btn small" onClick={onReject}>
            Reject
          </button>
          <button type="button" className="btn small primary" disabled={readOnly} onClick={() => onApprove(alsoLibrary)}>
            Approve
          </button>
        </div>
      )}
    </article>
  );
}
