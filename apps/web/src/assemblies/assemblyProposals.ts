import { useSyncExternalStore } from 'react';
import type { Assembly, AssemblyVariable, MarkupStore } from '@nb/markup';
import { assemblyProblems } from './quantities';
import { loadLibrary, writeLibrary } from './library';

/**
 * Assembly proposals: an AI client (through the local MCP server) proposes a new takeoff assembly,
 * or an edit to one, with its reasoning, and a person approves or rejects each one in the Review
 * AI Proposals panel. Proposals are kept apart from the document, on this device, until approved:
 * nothing reaches the document until a person approves it. Approving writes the document's
 * assemblies the way the Assemblies window does (one undo step); the proposals then leave the
 * store. Rejections stay here.
 */

/** One quantity the proposed assembly produces. No id: one is made when a person approves. */
export interface ProposedItem {
  name: string;
  unit: string;
  formula: string;
  waste?: number;
  roundUp?: boolean;
  roundTo?: number;
  unitCost?: number;
}

/**
 * An assembly as proposed, without ids: approving may be replacing an assembly that has them. A
 * missing `description` means "leave the notes as they are" (an edit); an empty one clears them.
 */
export interface ProposedAssembly {
  name: string;
  description?: string;
  variables: AssemblyVariable[];
  items: ProposedItem[];
}

export interface AssemblyProposal extends ProposedAssembly {
  id: string;
  /** The id of the assembly this proposes an edit of. */
  replacesId?: string;
  /** Why the AI proposed this, for the person checking it. */
  reasoning: string;
  /** Who proposed it: the AI client's name. */
  source: string;
  createdAt: number;
  /** Set when a person rejects it (approvals leave the store, written into the document). */
  rejected?: { by: string; at: number };
}

/** The most proposals kept for one document, so a forgotten flood cannot swamp the review panel. */
const MAX_STORED = 200;

const KEY = (fileHash: string) => `nb.assemblyProposals.${fileHash}`;

/** A proposal as stored has to have this shape; anything else in the storage is dropped. */
function shaped(raw: unknown): raw is AssemblyProposal {
  if (!raw || typeof raw !== 'object') return false;
  const p = raw as Record<string, unknown>;
  if (typeof p.id !== 'string' || typeof p.name !== 'string' || typeof p.reasoning !== 'string' || typeof p.source !== 'string' || typeof p.createdAt !== 'number') return false;
  if (p.replacesId !== undefined && typeof p.replacesId !== 'string') return false;
  if (p.description !== undefined && typeof p.description !== 'string') return false;
  if (!Array.isArray(p.variables) || !p.variables.every((v) => !!v && typeof v === 'object' && typeof (v as Record<string, unknown>).name === 'string' && typeof (v as Record<string, unknown>).value === 'number')) return false;
  return Array.isArray(p.items) && p.items.every((i) => !!i && typeof i === 'object' && typeof (i as Record<string, unknown>).name === 'string' && typeof (i as Record<string, unknown>).unit === 'string' && typeof (i as Record<string, unknown>).formula === 'string');
}

function load(fileHash: string): AssemblyProposal[] {
  try {
    const raw = localStorage.getItem(KEY(fileHash));
    const list = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(list) ? (list.filter(shaped) as AssemblyProposal[]) : [];
  } catch {
    return [];
  }
}

/** What a proposal targets: the assembly it would replace, else its own name. */
const targetOf = (p: Pick<AssemblyProposal, 'replacesId' | 'name'>) => p.replacesId ?? p.name.trim().toLowerCase();

/** Assembly proposals by document (its file hash), kept in this browser's storage. */
class AssemblyProposalStore {
  private cache = new Map<string, AssemblyProposal[]>();
  private listeners = new Set<() => void>();

  list(fileHash: string): AssemblyProposal[] {
    let list = this.cache.get(fileHash);
    if (!list) this.cache.set(fileHash, (list = load(fileHash)));
    return list;
  }

  /** Keeps the list; false when the browser would not keep it (the list stays for this visit). */
  private save(fileHash: string, list: AssemblyProposal[]): boolean {
    let persisted = true;
    try {
      localStorage.setItem(KEY(fileHash), JSON.stringify(list));
    } catch {
      persisted = false;
    }
    this.cache.set(fileHash, list);
    this.listeners.forEach((l) => l());
    return persisted;
  }

  /**
   * Adds proposals. A new one for the same assembly replaces its older pending one. The document
   * keeps at most 200: the oldest rejected ones are dropped first, and a call that would still
   * leave more is refused with nothing kept. False when the browser would not keep them, in which
   * case nothing is kept in memory either, so the caller is not told it succeeded.
   */
  add(fileHash: string, proposals: AssemblyProposal[]): boolean {
    const before = this.list(fileHash);
    const fresh = new Set(proposals.map(targetOf));
    let kept = [...before.filter((p) => p.rejected || !fresh.has(targetOf(p))), ...proposals];
    const overflow = kept.length - MAX_STORED;
    if (overflow > 0) {
      // The oldest rejected proposals go first; pending ones are never dropped quietly.
      const drop = kept.flatMap((p, i) => (p.rejected ? [i] : [])).slice(0, overflow);
      if (drop.length < overflow) throw new Error(`That is more than the ${MAX_STORED} assembly proposals kept for one drawing. Review or reject some first (Tools › Takeoff › Review AI Proposals).`);
      kept = kept.filter((_, i) => !drop.includes(i));
    }
    if (!this.save(fileHash, kept)) {
      this.cache.set(fileHash, before);
      this.listeners.forEach((l) => l());
      return false;
    }
    return true;
  }

  reject(fileHash: string, ids: readonly string[], by: string) {
    const at = Date.now();
    this.save(
      fileHash,
      this.list(fileHash).map((p) => (ids.includes(p.id) ? { ...p, rejected: { by, at } } : p)),
    );
  }

  /** Takes a rejection back (it is pending again). */
  reopen(fileHash: string, ids: readonly string[]) {
    this.save(
      fileHash,
      this.list(fileHash).map((p) => (ids.includes(p.id) ? { ...p, rejected: undefined } : p)),
    );
  }

  /** Forgets proposals once they are written into the document. */
  remove(fileHash: string, ids: readonly string[]) {
    this.save(
      fileHash,
      this.list(fileHash).filter((p) => !ids.includes(p.id)),
    );
  }

  subscribe = (l: () => void) => {
    this.listeners.add(l);
    return () => void this.listeners.delete(l);
  };
}

export const assemblyProposals = new AssemblyProposalStore();

export function useAssemblyProposals(fileHash: string | null): AssemblyProposal[] {
  return useSyncExternalStore(assemblyProposals.subscribe, () => (fileHash ? assemblyProposals.list(fileHash) : NONE));
}
const NONE: AssemblyProposal[] = [];

/**
 * The ids in `now` that were not in `before`. The review notice compares ids, not counts: an
 * approval takes ids away without the proposal store saying anything, so a count alone can hide a
 * proposal that arrived in the same moment.
 */
export function newlyWaiting(before: ReadonlySet<string>, now: ReadonlySet<string>): string[] {
  return [...now].filter((id) => !before.has(id));
}

/**
 * The proposal as an assembly: an edit keeps the assembly's id and its notes (unless the proposal
 * sends new ones), and an item with the same name as an existing one keeps that item's id, so
 * measurements keep their per-item overrides.
 */
function asAssembly(p: AssemblyProposal, existing: Assembly | undefined): Assembly {
  const oldItems = new Map(existing?.items.map((i) => [i.name.trim().toLowerCase(), i.id]));
  const description = p.description !== undefined ? p.description : existing?.description;
  return {
    id: existing?.id ?? crypto.randomUUID(),
    name: p.name,
    ...(description ? { description } : {}),
    variables: p.variables.map((v) => ({ ...v })),
    items: p.items.map((item) => ({ id: oldItems.get(item.name.trim().toLowerCase()) ?? crypto.randomUUID(), ...item })),
  };
}

/**
 * Approves assembly proposals as one undo step: each is written into the document's assemblies
 * the way the Assemblies window saves them, and the proposals leave the store. Ones whose target
 * is gone (the assembly it replaces was deleted, or another assembly has taken the name) or that
 * no longer pass the Assemblies window's checks (a column or assembly named like one of the
 * proposal's variables has been added since) are skipped and returned. A proposal no longer in
 * the store (approved already, say by a second click) is ignored. `alsoLibrary` saves the
 * approved assemblies into My Assemblies too, merged by name alone (`libraryReplaced` names the
 * entries it replaced); `librarySaved` says whether the browser kept them.
 */
export function approveAssemblyProposals(store: MarkupStore, proposals: readonly AssemblyProposal[], alsoLibrary = false): { approved: string[]; skipped: string[]; librarySaved: boolean; libraryReplaced: string[] } {
  if (store.readOnly) throw new Error('This document is read-only, so assembly proposals cannot be approved into it.');
  const waiting = new Set(assemblyProposals.list(store.fileHash).map((p) => p.id));
  const columns = store.columnSet().columns;
  const out = [...store.assemblies()];
  const approved: string[] = [];
  const skipped: string[] = [];
  const written: Assembly[] = [];
  let librarySaved = false;
  let libraryReplaced: string[] = [];
  for (const p of proposals) {
    if (!waiting.has(p.id)) continue;
    const at = p.replacesId ? out.findIndex((a) => a.id === p.replacesId) : -1;
    const taken = out.some((a) => a.id !== p.replacesId && a.name.trim().toLowerCase() === p.name.trim().toLowerCase());
    if (taken || (p.replacesId && at < 0)) {
      skipped.push(p.id);
      continue;
    }
    const assembly = asAssembly(p, at >= 0 ? out[at] : undefined);
    if (assemblyProblems(assembly, out.filter((a) => a.id !== assembly.id), columns).size) {
      skipped.push(p.id);
      continue;
    }
    if (at >= 0) out[at] = assembly;
    else out.push(assembly);
    written.push(assembly);
    approved.push(p.id);
  }
  if (approved.length) {
    store.setAssemblies(out);
    assemblyProposals.remove(store.fileHash, approved);
    if (alsoLibrary) {
      // Merged by name alone: an id from another document would dodge a rename.
      const library = [...loadLibrary()];
      libraryReplaced = [];
      for (const a of written) {
        const at = library.findIndex((l) => l.name.trim().toLowerCase() === a.name.trim().toLowerCase());
        if (at >= 0) {
          libraryReplaced.push(library[at]!.name);
          library[at] = { ...a, id: library[at]!.id };
        } else library.push({ ...a });
      }
      librarySaved = writeLibrary(library);
    }
  }
  return { approved, skipped, librarySaved, libraryReplaced };
}
