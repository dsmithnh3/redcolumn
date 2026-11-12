import { useSyncExternalStore } from 'react';
import { mergeAssemblies, type Assembly, type AssemblyVariable, type MarkupStore } from '@nb/markup';
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

/** An assembly as proposed, without ids: approving may be replacing an assembly that has them. */
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

const KEY = (fileHash: string) => `nb.assemblyProposals.${fileHash}`;

/** A proposal as stored has to have this shape; anything else in the storage is dropped. */
function shaped(raw: unknown): raw is AssemblyProposal {
  if (!raw || typeof raw !== 'object') return false;
  const p = raw as Record<string, unknown>;
  return typeof p.id === 'string' && typeof p.name === 'string' && typeof p.reasoning === 'string' && typeof p.source === 'string' && typeof p.createdAt === 'number' && Array.isArray(p.items) && p.items.every((i) => !!i && typeof i === 'object' && typeof (i as Record<string, unknown>).name === 'string');
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

  private save(fileHash: string, list: AssemblyProposal[]) {
    this.cache.set(fileHash, list);
    try {
      localStorage.setItem(KEY(fileHash), JSON.stringify(list));
    } catch {
      // Kept for this visit only.
    }
    this.listeners.forEach((l) => l());
  }

  /** Adds proposals. A new one for the same assembly replaces its older pending one. */
  add(fileHash: string, proposals: AssemblyProposal[]) {
    const fresh = new Set(proposals.map(targetOf));
    this.save(fileHash, [...this.list(fileHash).filter((p) => p.rejected || !fresh.has(targetOf(p))), ...proposals]);
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
 * The proposal as an assembly: an edit keeps the assembly's id, and an item with the same name as
 * an existing one keeps that item's id, so measurements keep their per-item overrides.
 */
function asAssembly(p: AssemblyProposal, existing: Assembly | undefined): Assembly {
  const oldItems = new Map(existing?.items.map((i) => [i.name.trim().toLowerCase(), i.id]));
  return {
    id: existing?.id ?? crypto.randomUUID(),
    name: p.name,
    ...(p.description ? { description: p.description } : {}),
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
 * approved assemblies into My Assemblies too, merged by name; `librarySaved` says whether the
 * browser kept them.
 */
export function approveAssemblyProposals(store: MarkupStore, proposals: readonly AssemblyProposal[], alsoLibrary = false): { approved: string[]; skipped: string[]; librarySaved: boolean } {
  if (store.readOnly) throw new Error('This document is read-only, so assembly proposals cannot be approved into it.');
  const waiting = new Set(assemblyProposals.list(store.fileHash).map((p) => p.id));
  const columns = store.columnSet().columns;
  const out = [...store.assemblies()];
  const approved: string[] = [];
  const skipped: string[] = [];
  const written: Assembly[] = [];
  let librarySaved = false;
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
    if (alsoLibrary) librarySaved = writeLibrary(mergeAssemblies(loadLibrary(), written));
  }
  return { approved, skipped, librarySaved };
}
