import { DEFAULT_STYLES, MarkupStore, measureProps, TYPE_INFO, type Markup, type MarkupStyle, type MarkupType, type Point, type StoredStitchGroup } from '@nb/markup';
import { STITCH_METHODS, stitchCall } from './assistantStitch';
import { VIEWPORT_METHODS, viewportCall } from './assistantViewports';
import { requireScaleFor, SCALE_METHODS, scaleCall, type ScaleMethod } from './assistantScales';
import { STANDARD_METHODS, standardCall, type AssistantStandards, type StandardMethod } from './assistantStandards';
import { ASSEMBLY_PROPOSAL_METHODS, assemblyProposalCall, type AssemblyProposalMethod } from './assistantAssemblies';
import { isMeasureKind, measureLabel, scaleEquation } from '@nb/measure';
import type { PageRegion, TextWord } from '@nb/pdf-core';
import { stateOf as proposalState, takeoffProposals, type TakeoffProposal } from '../assemblies/proposals';
import { MEASURED_NAMES } from '../assemblies/quantities';
import { documentParagraphs, exactMatches, type AiDecision, type ChangeKind, type Paragraph } from '../compare/semantic';
import { semanticCompare, type SemanticState } from '../compare/semanticStore';

/** A picture of part of a page, as `get_page_image` returns it. */
export interface PageImage {
  /** PNG bytes, base64. */
  png: string;
  width: number;
  height: number;
  /** Pixels per PDF point. */
  scale: number;
  /** The part of the page drawn, in PDF points, after clipping to the page. */
  region: PageRegion;
}

/** One open drawing the assistant can see. Page indexes are 0-based, as the viewer stores them. */
export interface AssistantDocument {
  id: string;
  name: string;
  pageCount: number;
  pageIndex: number;
  active: boolean;
  store: MarkupStore;
  /** Displayed page size in PDF points. Absent in tests that never ask for a page. */
  pageSize?: (pageIndex: number) => { width: number; height: number } | null;
  /**
   * Part of a page as a picture at `scale` pixels per point, with its markups when asked. Null
   * when the region is off the page. Absent in tests, and in a browser without a renderer.
   */
  render?: (pageIndex: number, region: PageRegion, scale: number, withMarkups: boolean) => Promise<PageImage | null>;
  /** The words on a page, in page space. Absent in tests. */
  text?: (pageIndex: number) => Promise<TextWord[]>;
  /** A page's line segments, `[x1, y1, x2, y2, ...]` in page space. Absent in tests. */
  geometry?: (pageIndex: number) => Promise<{ segments: ArrayLike<number> }>;
  /**
   * A stitched group drawn as one picture: `bounds` (the group's world space) at `scale` pixels
   * per world unit, each sheet cut at its region and match lines. Absent in tests.
   */
  renderStitch?: (group: StoredStitchGroup, bounds: PageRegion, scale: number, withMarkups: boolean) => Promise<{ png: string; width: number; height: number } | null>;
}

/** The stitched view and Stitch to New Page, for a drawing on screen. */
export interface AssistantStitching {
  /** Shows `group` as the stitched view, or the normal view with null. False when the drawing is not on screen. */
  show(documentId: string, group: StoredStitchGroup | null, focusPage: number): boolean;
  /** Stitch to New Page. Null when the drawing is not on screen or is read-only. */
  toPage(documentId: string, group: StoredStitchGroup, withMarkups: boolean): Promise<{ pageIndex: number; markups: number } | null>;
}

/** A menu command the assistant can list and run. */
export interface AssistantCommand {
  id: string;
  label: string;
  category: string;
  enabled: boolean;
  checked?: boolean;
  run: () => void;
}

/** Text and subject written onto a markup as it is placed. */
export interface PlaceExtras {
  text?: string;
  comment?: string;
  subject?: string;
}

/**
 * The drawing tools for one pane. `place` uses the armed tool's defaults; `select` is what
 * align, group, and copy act on.
 */
export interface AssistantTools {
  owns(store: MarkupStore): boolean;
  place(type: MarkupType, pageIndex: number, points: Point[], extras?: PlaceExtras): Markup;
  /** Cloud+: a cloud and the callout grouped with it. */
  placeCloudPlus?(pageIndex: number, points: Point[], extras?: PlaceExtras): { cloud: Markup; callout: Markup };
  /** Smart Fill: the markup made from the region enclosed around a point. */
  fill?(pageIndex: number, at: Point, extras?: PlaceExtras): Promise<Markup>;
  select(ids: Iterable<string>): void;
}

/** What the renderer supplies beyond the open drawings. */
export interface AssistantEnv {
  commands?: () => ReadonlyMap<string, AssistantCommand>;
  toolsFor?: (documentId: string) => AssistantTools | null;
  confirm?: (message: string) => Promise<boolean>;
  stitch?: AssistantStitching;
  /** Brings the AI panel into view, so a person sees a comparison the assistant started. */
  revealComparison?: () => void;
  /** The profile's drawing standards (get_drawing_standard, propose_standard_hints). */
  standards?: AssistantStandards;
}

/**
 * Commands that delete, discard, import, write the file, or change pages. `run_command` asks
 * before these. Arming a tool does not: the gesture still happens on the page.
 */
const DESTRUCTIVE_COMMANDS = new Set([
  'edit.delete',
  'edit.cut',
  'file.close',
  'file.closeAll',
  'file.revert',
  'file.importXfdf',
  'file.importFdf',
  'file.importBax',
  'file.importPdfMarkups',
  'file.save',
  'file.saveAll',
  'document.rotatePages',
  'document.rotateCounterclockwise',
  'document.insertPages',
  'document.bulkInsertPages',
  'document.insertBlank',
  'document.deletePages',
  'document.replacePages',
  'document.cropPages',
  'document.pageSetup',
  'document.headerFooter',
  'document.numberPages',
  'document.flatten',
  'document.unflatten',
  'document.reduceSize',
  'document.repair',
  'document.colour',
  'document.ocr',
  'document.slipSheet',
  'document.labelRegions',
  'document.scaleRegions',
  'tools.applyRedactions',
  'tools.sign',
  'tools.autoFields',
  'document.security',
  'document.pdfa',
]);

/** Open, Save As, and Print always hand the dialog to the person, even if a label loses its ellipsis. */
const NEEDS_YOU = new Set(['file.open', 'file.saveAs', 'file.print']);

const FILE_TYPES = new Set<MarkupType>(['image', 'stamp', 'signature', 'attachment']);

export function isDestructiveCommand(id: string): boolean {
  return id.startsWith('batch.') || DESTRUCTIVE_COMMANDS.has(id);
}

/** The active drawing wins when the same file is open in both panes. */
export function assistantDocuments(tabs: readonly AssistantDocument[]): AssistantDocument[] {
  const active = tabs.find((t) => t.active);
  const ordered = active ? [active, ...tabs] : tabs;
  const out: AssistantDocument[] = [];
  const seen = new Set<string>();
  for (const tab of ordered) {
    if (seen.has(tab.id)) continue;
    seen.add(tab.id);
    out.push(tab);
  }
  return out;
}

function sheetOf(doc: AssistantDocument) {
  const info = doc.store.allSheets()[doc.pageIndex];
  return { sheetNumber: info?.number ?? null, sheetTitle: info?.title ?? null };
}

function publicDocument(doc: AssistantDocument) {
  return {
    id: doc.id,
    name: doc.name,
    pageCount: doc.pageCount,
    page: doc.pageIndex + 1,
    active: doc.active,
    ...sheetOf(doc),
  };
}

function measurement(doc: AssistantDocument, markup: Markup): string | undefined {
  if (!isMeasureKind(markup.type) || markup.points.length === 0) return undefined;
  return measureLabel(markup.type, markup.points, doc.store.scaleOf(markup), measureProps(markup));
}

/** `brief` leaves out the geometry, style, and author: what a list needs to pick markups, at a fraction of the size. */
function publicMarkup(doc: AssistantDocument, markup: Markup, brief = false) {
  const columns = doc.store.columnSet();
  const status = columns.statuses.find((s) => s.id === markup.status)?.name ?? markup.status;
  const fields: Record<string, string> = {};
  for (const column of columns.columns) {
    const value = markup.fields?.[column.id];
    if (value) fields[column.name] = value;
  }
  const measured = measurement(doc, markup);
  return {
    id: markup.id,
    ...(markup.seq !== undefined ? { seq: markup.seq } : {}),
    type: markup.type,
    page: markup.pageIndex + 1,
    ...(brief ? {} : { points: markup.points, style: { stroke: markup.style.stroke, fill: markup.style.fill, width: markup.style.width } }),
    ...(markup.text ? { text: markup.text } : {}),
    ...(markup.subject ? { subject: markup.subject } : {}),
    ...(markup.comment ? { comment: markup.comment } : {}),
    status,
    ...(brief ? {} : { author: markup.author }),
    ...(measured ? { measurement: measured } : {}),
    ...(Object.keys(fields).length ? { fields } : {}),
    ...(markup.hidden ? { hidden: true } : {}),
    ...(markup.locked ? { locked: true } : {}),
  };
}

export function findDocument(docs: readonly AssistantDocument[], name: unknown): AssistantDocument {
  const open = assistantDocuments(docs);
  if (!open.length) throw new Error('No document is open.');
  if (typeof name !== 'string' || !name.trim()) {
    const active = open.find((d) => d.active) ?? open[0];
    if (!active) throw new Error('No document is open.');
    return active;
  }
  const wanted = name.trim().toLowerCase();
  const matches = open.filter((d) => d.id === name || d.name.toLowerCase() === wanted);
  const found = matches.find((d) => d.active) ?? matches[0];
  if (!found) throw new Error(`No open document matches "${name}".`);
  return found;
}

export function record(params: unknown): Record<string, unknown> {
  if (!params || typeof params !== 'object' || Array.isArray(params)) return {};
  return params as Record<string, unknown>;
}

export function pageIndexOf(doc: AssistantDocument, raw: unknown): number {
  const page = typeof raw === 'number' ? raw : Number(raw);
  if (!Number.isInteger(page) || page < 1 || page > doc.pageCount) throw new Error(`Page ${String(raw)} is not in ${doc.name} (1–${doc.pageCount}).`);
  return page - 1;
}

/** A list without a limit stops here, and says where to carry on from. */
const DEFAULT_LIST_LIMIT = 500;
const MAX_LIST_LIMIT = 2000;

/** The most markups, or items, one batch call changes. A bigger job is several calls. */
const MAX_BATCH = 500;

function wholeNumber(raw: unknown, name: string, min: number): number {
  const n = typeof raw === 'number' ? raw : Number(raw);
  if (!Number.isInteger(n) || n < min) throw new Error(`${name} has to be a whole number, ${min} or more.`);
  return n;
}

function statusOf(doc: AssistantDocument, raw: string) {
  const wanted = raw.trim().toLowerCase();
  const status = doc.store.columnSet().statuses.find((s) => s.id.toLowerCase() === wanted || s.name.toLowerCase() === wanted);
  if (!status) throw new Error(`"${raw}" is not a status on ${doc.name}.`);
  return status;
}

function tally(counts: Record<string, number>, key: string) {
  counts[key] = (counts[key] ?? 0) + 1;
}

function listMarkups(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  let markups = doc.store.all();
  if (args.page !== undefined) {
    const index = pageIndexOf(doc, args.page);
    markups = markups.filter((m) => m.pageIndex === index);
  }
  if (args.type !== undefined) {
    const type = markupType(args.type);
    markups = markups.filter((m) => m.type === type);
  }
  if (typeof args.status === 'string' && args.status.trim()) {
    const status = statusOf(doc, args.status);
    markups = markups.filter((m) => m.status === status.id);
  }
  if (typeof args.search === 'string' && args.search.trim()) {
    const wanted = args.search.trim().toLowerCase();
    markups = markups.filter((m) => [m.text, m.subject, m.comment].some((value) => value?.toLowerCase().includes(wanted)));
  }
  if (args.summary === true) {
    const statuses = doc.store.columnSet().statuses;
    const byType: Record<string, number> = {};
    const byStatus: Record<string, number> = {};
    const byPage: Record<string, number> = {};
    for (const m of markups) {
      tally(byType, m.type);
      tally(byStatus, statuses.find((s) => s.id === m.status)?.name ?? m.status);
      tally(byPage, String(m.pageIndex + 1));
    }
    return { document: doc.name, id: doc.id, total: markups.length, byType, byStatus, byPage };
  }
  const limit = args.limit === undefined ? DEFAULT_LIST_LIMIT : Math.min(wholeNumber(args.limit, 'Limit', 1), MAX_LIST_LIMIT);
  const offset = args.offset === undefined ? 0 : wholeNumber(args.offset, 'Offset', 0);
  const brief = args.detail === 'brief';
  const shown = markups.slice(offset, offset + limit);
  const next = offset + shown.length;
  return {
    document: doc.name,
    id: doc.id,
    total: markups.length,
    offset,
    ...(next < markups.length ? { nextOffset: next } : {}),
    markups: shown.map((m) => publicMarkup(doc, m, brief)),
  };
}

function findMarkup(doc: AssistantDocument, args: Record<string, unknown>): Markup {
  if (typeof args.id === 'string' && args.id) {
    const found = doc.store.get(args.id);
    if (!found) throw new Error(`No markup ${args.id} is on ${doc.name}.`);
    return found;
  }
  const seq = typeof args.seq === 'number' ? args.seq : typeof args.seq === 'string' && args.seq ? Number(args.seq) : undefined;
  if (seq !== undefined && Number.isInteger(seq)) {
    const found = doc.store.all().find((m) => m.seq === seq);
    if (!found) throw new Error(`No markup ${seq} is on ${doc.name}.`);
    return found;
  }
  throw new Error('Say which markup to change: its id or its number (seq) from list_markups.');
}

/** The change `set_markup` makes to one markup. Throws before anything is touched. */
function setPatch(doc: AssistantDocument, markup: Markup, args: Record<string, unknown>): Partial<Omit<Markup, 'id'>> {
  const patch: Partial<Omit<Markup, 'id'>> = {};
  if (typeof args.status === 'string') patch.status = statusOf(doc, args.status).id;
  if (typeof args.comment === 'string') patch.comment = args.comment ? args.comment : undefined;
  if (typeof args.subject === 'string') patch.subject = args.subject ? args.subject : undefined;
  if (typeof args.field === 'string' && args.field) {
    const wanted = args.field.trim().toLowerCase();
    const column = doc.store.columnSet().columns.find((c) => c.id.toLowerCase() === wanted || c.name.toLowerCase() === wanted);
    if (!column) throw new Error(`${doc.name} has no column named "${args.field}".`);
    if (column.type === 'formula') throw new Error(`${column.name} is calculated, so it cannot be set.`);
    if (typeof args.value !== 'string') throw new Error(`Give a value for ${column.name}.`);
    patch.fields = { ...markup.fields, [column.id]: args.value };
  }
  if (!('status' in patch) && !('comment' in patch) && !('subject' in patch) && !patch.fields) throw new Error('Say what to change: status, comment, subject, or a column.');
  return patch;
}

/** The markup after `set_markup` wrote `patch`, or why the store refused it (someone else's markup, say). */
function appliedSet(doc: AssistantDocument, id: string, patch: Partial<Omit<Markup, 'id'>>): Markup {
  const next = doc.store.get(id);
  if (!next) throw new Error(`No markup ${id} is on ${doc.name}.`);
  if (patch.status !== undefined && next.status !== patch.status) throw new Error('That markup cannot be changed.');
  if ('comment' in patch && (next.comment ?? '') !== (patch.comment ?? '')) throw new Error('That markup cannot be changed.');
  if ('subject' in patch && (next.subject ?? '') !== (patch.subject ?? '')) throw new Error('That markup cannot be changed.');
  if (patch.fields) {
    for (const [fieldId, value] of Object.entries(patch.fields)) {
      if ((next.fields?.[fieldId] ?? '') !== value) throw new Error('That markup cannot be changed.');
    }
  }
  return next;
}

function setMarkup(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const markup = findMarkup(doc, args);
  const patch = setPatch(doc, markup, args);
  doc.store.checkpoint();
  doc.store.update(markup.id, patch);
  return { document: doc.name, markup: publicMarkup(doc, appliedSet(doc, markup.id, patch)) };
}

function errorText(err: unknown): string {
  return err instanceof Error && err.message ? err.message : 'That markup cannot be changed.';
}

function batchItems(raw: unknown, empty: string): Record<string, unknown>[] {
  if (!Array.isArray(raw) || raw.length === 0 || raw.some((item) => !item || typeof item !== 'object' || Array.isArray(item))) throw new Error(empty);
  if (raw.length > MAX_BATCH) throw new Error(`That is ${raw.length} at once. The most is ${MAX_BATCH}: send the rest in another call.`);
  return raw as Record<string, unknown>[];
}

/** Markups named by id or by number, each once, in the order given. */
function pickMarkups(doc: AssistantDocument, args: Record<string, unknown>): Markup[] {
  const picked: Markup[] = [];
  const seen = new Set<string>();
  const add = (markup: Markup) => {
    if (seen.has(markup.id)) return;
    seen.add(markup.id);
    picked.push(markup);
  };
  const named = (Array.isArray(args.ids) ? args.ids.length : 0) + (Array.isArray(args.seqs) ? args.seqs.length : 0);
  if (named > MAX_BATCH) throw new Error(`That is ${named} markups at once. The most is ${MAX_BATCH}: send the rest in another call.`);
  if (args.ids !== undefined) for (const id of idsOf(args.ids, 'ids has to be a list of markup ids from list_markups.')) add(findMarkup(doc, { id }));
  if (args.seqs !== undefined) {
    if (!Array.isArray(args.seqs) || args.seqs.some((seq) => typeof seq !== 'number' || !Number.isInteger(seq))) throw new Error('seqs has to be a list of markup numbers from list_markups.');
    for (const seq of args.seqs as number[]) add(findMarkup(doc, { seq }));
  }
  if (!picked.length) throw new Error('Say which markups to change: ids or seqs from list_markups.');
  return picked;
}

/**
 * One status, comment, subject, or column value for many markups, as one undo step. Every markup
 * is checked first, so a bad name changes nothing. Only what the store refuses afterwards is
 * listed in `failed`.
 */
function setMarkups(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const plan = pickMarkups(doc, args).map((markup) => ({ markup, patch: setPatch(doc, markup, args) }));
  doc.store.checkpoint();
  doc.store.batch(() => {
    for (const { markup, patch } of plan) doc.store.update(markup.id, patch);
  });
  const updated: string[] = [];
  const failed: { id: string; error: string }[] = [];
  for (const { markup, patch } of plan) {
    try {
      appliedSet(doc, markup.id, patch);
      updated.push(markup.id);
    } catch (err) {
      failed.push({ id: markup.id, error: errorText(err) });
    }
  }
  return { document: doc.name, updated, ...(failed.length ? { failed } : {}) };
}

function markupType(raw: unknown): MarkupType {
  if (typeof raw !== 'string' || !(raw in TYPE_INFO)) throw new Error(`Unknown markup type "${String(raw)}".`);
  return raw as MarkupType;
}

export function pointsOf(raw: unknown): Point[] {
  if (!Array.isArray(raw) || raw.length === 0) throw new Error('Give the points to draw.');
  return raw.map((point, i) => {
    if (!Array.isArray(point) || point.length < 2 || typeof point[0] !== 'number' || typeof point[1] !== 'number' || !Number.isFinite(point[0]) || !Number.isFinite(point[1])) {
      throw new Error(`Point ${i + 1} has to be an x and a y.`);
    }
    return [point[0], point[1]];
  });
}

/** A note, flag, or typewriter placed with one click gets a box, the same way the tool does at zoom 1. */
function expandPlaced(type: MarkupType, points: Point[]): Point[] {
  if (points.length !== 1) return points;
  const a = points[0]!;
  if (type === 'note' || type === 'flag') {
    const s = 22;
    return [a, [a[0] + (type === 'flag' ? s * 0.8 : s), a[1] + s]];
  }
  if (type === 'typewriter') return [a, [a[0] + 48, a[1] + 22.4]];
  return points;
}

function requirePoints(type: MarkupType, points: Point[]) {
  const info = TYPE_INFO[type];
  if (type === 'callout') {
    if (points.length < 4) throw new Error("Callout needs a tip, a knee, and the box's two corners.");
    return;
  }
  if (info.draw === 'text') {
    if (points.length < 2 || points.length % 2 !== 0) throw new Error(`${info.label} needs a pair of corners for each line of text.`);
    return;
  }
  if (info.click) {
    if (info.click.fixed !== undefined && points.length !== info.click.fixed) throw new Error(`${info.label} needs ${info.click.fixed} points.`);
    else if (points.length < info.click.min) throw new Error(`${info.label} needs at least ${info.click.min} point${info.click.min === 1 ? '' : 's'}.`);
    return;
  }
  if (info.draw === 'freehand') {
    if (points.length < 2) throw new Error(`${info.label} needs at least 2 points.`);
    return;
  }
  if (info.draw === 'place') {
    if (points.length < 1 || points.length > 2) throw new Error(`${info.label} needs one point, or two corners.`);
    return;
  }
  if (points.length !== 2) throw new Error(`${info.label} needs two opposite corners.`);
}

function refuseFileType(type: MarkupType) {
  if (!FILE_TYPES.has(type)) return;
  const command = type === 'signature' ? 'tools.sign' : `tool.${type}`;
  throw new Error(`${TYPE_INFO[type].label} needs a file. Run ${command} with run_command, then place it in redcolumn.`);
}

function extrasOf(args: Record<string, unknown>): PlaceExtras {
  const extras: PlaceExtras = {};
  if (typeof args.text === 'string' && args.text) extras.text = args.text;
  if (typeof args.comment === 'string' && args.comment) extras.comment = args.comment;
  if (typeof args.subject === 'string' && args.subject) extras.subject = args.subject;
  return extras;
}

function placeOnStore(doc: AssistantDocument, type: MarkupType, pageIndex: number, points: Point[], extras: PlaceExtras): Markup {
  const now = Date.now();
  const markup: Markup = {
    id: crypto.randomUUID(),
    type,
    pageIndex,
    points,
    style: { ...DEFAULT_STYLES[type] },
    status: 'none',
    author: MarkupStore.author || 'redcolumn',
    createdAt: now,
    modifiedAt: now,
    ...(extras.text ? { text: extras.text } : {}),
    ...(extras.comment ? { comment: extras.comment } : {}),
    ...(extras.subject ? { subject: extras.subject } : {}),
  };
  doc.store.checkpoint();
  doc.store.add(markup);
  return doc.store.get(markup.id) ?? markup;
}

/** Cloud+ and Smart Fill are tools, not markup types: they make markups from a gesture. */
const GESTURE_TOOLS = new Set(['cloudPlus', 'dynamicFill']);

async function placeGesture(doc: AssistantDocument, tool: string, pageIndex: number, points: Point[], extras: PlaceExtras, env: AssistantEnv) {
  const label = tool === 'cloudPlus' ? 'Cloud+' : 'Smart Fill';
  const tools = env.toolsFor?.(doc.id) ?? null;
  if (!tools?.owns(doc.store)) throw new Error(`${doc.name} has to be the drawing on screen for ${label}.`);
  if (tool === 'cloudPlus') {
    if (!tools.placeCloudPlus) throw new Error('Cloud+ is not available.');
    requirePoints('cloud', points);
    const { cloud, callout } = tools.placeCloudPlus(pageIndex, points, extras);
    return { document: doc.name, markup: publicMarkup(doc, cloud), callout: publicMarkup(doc, callout) };
  }
  if (!tools.fill) throw new Error('Smart Fill is not available.');
  if (points.length !== 1) throw new Error('Smart Fill needs one point inside a closed outline.');
  requireScaleFor(doc, 'area', pageIndex, points[0]);
  const markup = await tools.fill(pageIndex, points[0]!, extras);
  return { document: doc.name, markup: publicMarkup(doc, markup) };
}

async function placeMarkup(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  if (args.page === undefined) throw new Error('Say which page to draw on.');
  const pageIndex = pageIndexOf(doc, args.page);
  if (typeof args.type === 'string' && GESTURE_TOOLS.has(args.type)) return placeGesture(doc, args.type, pageIndex, pointsOf(args.points), extrasOf(args), env);
  const type = markupType(args.type);
  refuseFileType(type);
  const points = pointsOf(args.points);
  requirePoints(type, points);
  requireScaleFor(doc, type, pageIndex, points[0]);
  const extras = extrasOf(args);
  const tools = env.toolsFor?.(doc.id) ?? null;
  const markup = tools?.owns(doc.store) ? tools.place(type, pageIndex, points, extras) : placeOnStore(doc, type, pageIndex, expandPlaced(type, points), extras);
  return { document: doc.name, markup: publicMarkup(doc, markup) };
}

/**
 * Several markups drawn as one undo step. Each item is what `place_markup` takes, with `page`
 * optional when the call gives one for all. Every item is checked first, so one bad item
 * draws nothing. Cloud+ and Smart Fill are gestures on the screen, so they stay in `place_markup`.
 */
function placeMarkups(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const items = batchItems(args.markups, 'Give markups: a list of what to draw, each with a type, a page, and points.');
  const plan = items.map((item, i) => {
    const where = `Markup ${i + 1}`;
    try {
      const page = item.page ?? args.page;
      if (page === undefined) throw new Error('Say which page to draw on.');
      const pageIndex = pageIndexOf(doc, page);
      if (typeof item.type === 'string' && GESTURE_TOOLS.has(item.type)) throw new Error(`${item.type} is one gesture on the screen. Use place_markup for it.`);
      const type = markupType(item.type);
      refuseFileType(type);
      const points = pointsOf(item.points);
      requirePoints(type, points);
      requireScaleFor(doc, type, pageIndex, points[0]);
      return { type, pageIndex, points, extras: extrasOf(item) };
    } catch (err) {
      throw new Error(`${where}: ${errorText(err)}`);
    }
  });
  const tools = env.toolsFor?.(doc.id) ?? null;
  const owned = tools?.owns(doc.store) ? tools : null;
  const placed: Markup[] = [];
  doc.store.checkpoint();
  doc.store.batch(() => {
    for (const { type, pageIndex, points, extras } of plan) {
      placed.push(owned ? owned.place(type, pageIndex, points, extras) : placeOnStore(doc, type, pageIndex, expandPlaced(type, points), extras));
    }
  });
  return {
    document: doc.name,
    placed: placed.map((m) => {
      const seq = doc.store.get(m.id)?.seq;
      return { id: m.id, ...(seq !== undefined ? { seq } : {}), type: m.type, page: m.pageIndex + 1 };
    }),
  };
}

function colorOf(value: unknown, name: string): string {
  if (typeof value !== 'string' || !/^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(value)) throw new Error(`${name} has to be a color like #ff0000.`);
  return value;
}

function styleOf(raw: unknown): Partial<MarkupStyle> {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('Style has to be stroke, fill, width, opacity, or font size.');
  const style = raw as Record<string, unknown>;
  if (!('stroke' in style) && !('fill' in style) && !('width' in style) && !('opacity' in style) && !('fontSize' in style)) {
    throw new Error('Style has to be stroke, fill, width, opacity, or font size.');
  }
  const patch: Partial<MarkupStyle> = {};
  if ('stroke' in style) patch.stroke = colorOf(style.stroke, 'Stroke');
  if ('fill' in style) patch.fill = style.fill === null ? null : colorOf(style.fill, 'Fill');
  if ('width' in style) {
    if (typeof style.width !== 'number' || !Number.isFinite(style.width) || style.width < 0) throw new Error('Width has to be a number of points.');
    patch.width = style.width;
  }
  if ('opacity' in style) {
    if (typeof style.opacity !== 'number' || !Number.isFinite(style.opacity) || style.opacity < 0 || style.opacity > 1) throw new Error('Opacity has to be between 0 and 1.');
    patch.opacity = style.opacity;
  }
  if ('fontSize' in style) {
    if (typeof style.fontSize !== 'number' || !Number.isFinite(style.fontSize) || style.fontSize <= 0) throw new Error('Font size has to be a number of points.');
    patch.fontSize = style.fontSize;
  }
  return patch;
}

function samePoints(a: readonly Point[], b: readonly Point[]): boolean {
  return a.length === b.length && a.every((p, i) => p[0] === b[i]?.[0] && p[1] === b[i]?.[1]);
}

/** The change `update_markup` makes to one markup. Throws before anything is touched. */
function updatePatch(doc: AssistantDocument, markup: Markup, args: Record<string, unknown>): Partial<Omit<Markup, 'id'>> {
  if (markup.locked) throw new Error('That markup is locked.');
  if (!doc.store.mayEdit(markup)) throw new Error('That markup cannot be changed.');
  const patch: Partial<Omit<Markup, 'id'>> = {};
  if (args.points !== undefined) {
    const points = pointsOf(args.points);
    requirePoints(markup.type, points);
    patch.points = points;
  }
  if (typeof args.text === 'string') patch.text = args.text ? args.text : undefined;
  if (args.style !== undefined) patch.style = { ...markup.style, ...styleOf(args.style) };
  if (!('points' in patch) && !('text' in patch) && !patch.style) throw new Error('Say what to change: points, text, or style.');
  return patch;
}

/** The markup after `update_markup` wrote `patch`, or why the store refused it. */
function appliedUpdate(doc: AssistantDocument, id: string, patch: Partial<Omit<Markup, 'id'>>): Markup {
  const next = doc.store.get(id);
  if (!next) throw new Error(`No markup ${id} is on ${doc.name}.`);
  if (patch.points && !samePoints(next.points, patch.points)) throw new Error('That markup cannot be changed.');
  if ('text' in patch && (next.text ?? '') !== (patch.text ?? '')) throw new Error('That markup cannot be changed.');
  if (patch.style) {
    if (next.style.stroke !== patch.style.stroke || next.style.fill !== patch.style.fill || next.style.width !== patch.style.width) throw new Error('That markup cannot be changed.');
    if (patch.style.opacity !== undefined && next.style.opacity !== patch.style.opacity) throw new Error('That markup cannot be changed.');
    if (patch.style.fontSize !== undefined && next.style.fontSize !== patch.style.fontSize) throw new Error('That markup cannot be changed.');
  }
  return next;
}

function updateMarkup(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const markup = findMarkup(doc, args);
  const patch = updatePatch(doc, markup, args);
  doc.store.checkpoint();
  doc.store.update(markup.id, patch);
  return { document: doc.name, markup: publicMarkup(doc, appliedUpdate(doc, markup.id, patch)) };
}

/** Several markups changed as one undo step: each item is what `update_markup` takes. Checked first, like `set_markups`. */
function updateMarkups(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const items = batchItems(args.markups, 'Give markups: a list of changes, each with an id or seq and points, text, or style.');
  const seen = new Set<string>();
  const plan = items.map((item) => {
    const markup = findMarkup(doc, item);
    if (seen.has(markup.id)) throw new Error(`Markup ${markup.seq ?? markup.id} is in the list twice.`);
    seen.add(markup.id);
    return { markup, patch: updatePatch(doc, markup, item) };
  });
  doc.store.checkpoint();
  doc.store.batch(() => {
    for (const { markup, patch } of plan) doc.store.update(markup.id, patch);
  });
  const updated: string[] = [];
  const failed: { id: string; error: string }[] = [];
  for (const { markup, patch } of plan) {
    try {
      appliedUpdate(doc, markup.id, patch);
      updated.push(markup.id);
    } catch (err) {
      failed.push({ id: markup.id, error: errorText(err) });
    }
  }
  return { document: doc.name, updated, ...(failed.length ? { failed } : {}) };
}

function idsOf(raw: unknown, empty: string): string[] {
  if (!Array.isArray(raw) || raw.some((id) => typeof id !== 'string' || !id)) throw new Error(empty);
  return raw as string[];
}

function markupsById(doc: AssistantDocument, ids: readonly string[]): Markup[] {
  return ids.map((id) => {
    const markup = doc.store.get(id);
    if (!markup) throw new Error(`No markup ${id} is on ${doc.name}.`);
    return markup;
  });
}

export async function ask(env: AssistantEnv, message: string) {
  const ok = env.confirm ? await env.confirm(message) : false;
  if (!ok) throw new Error('Not confirmed.');
}

async function deleteMarkups(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (doc.store.readOnly) throw new Error(`${doc.name} is read-only.`);
  const ids = idsOf(args.ids, 'Say which markups to delete: their ids from list_markups.');
  if (!ids.length) throw new Error('Say which markups to delete: their ids from list_markups.');
  const found = markupsById(doc, ids);
  if (found.some((m) => m.locked)) throw new Error(found.length === 1 ? 'That markup is locked.' : 'Those markups are locked.');
  if (found.some((m) => !doc.store.mayEdit(m))) throw new Error('That markup cannot be changed.');
  const labels = found.map((m) => (m.seq !== undefined ? String(m.seq) : m.id));
  const what = found.length === 1 ? `markup ${labels[0]}` : `${found.length} markups (${labels.join(', ')})`;
  await ask(env, `Delete ${what} on ${doc.name}? You can still Undo in redcolumn.`);
  doc.store.checkpoint();
  doc.store.remove(ids);
  if (ids.some((id) => doc.store.get(id))) throw new Error('That markup cannot be changed.');
  return { document: doc.name, deleted: ids };
}

function selectMarkups(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  const ids = idsOf(args.ids, 'Say which markups to select: their ids from list_markups.');
  markupsById(doc, ids);
  const tools = env.toolsFor?.(doc.id) ?? null;
  if (!tools?.owns(doc.store)) throw new Error(`${doc.name} is not the drawing on screen.`);
  tools.select(ids);
  return { document: doc.name, selected: ids };
}

function getPage(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  const pageIndex = args.page === undefined ? doc.pageIndex : pageIndexOf(doc, args.page);
  if (pageIndex < 0 || pageIndex >= doc.pageCount) throw new Error(`Page ${pageIndex + 1} is not in ${doc.name} (1–${doc.pageCount}).`);
  const size = doc.pageSize?.(pageIndex) ?? null;
  if (!size) throw new Error(`Page size is not available for page ${pageIndex + 1} of ${doc.name}.`);
  const scale = doc.store.hasScale(pageIndex) ? scaleEquation(doc.store.scaleFor(pageIndex)) : undefined;
  const viewports = doc.store.allViewports().filter((v) => v.pageIndex === pageIndex).length;
  return { document: doc.name, id: doc.id, page: pageIndex + 1, width: size.width, height: size.height, ...(scale ? { scale } : {}), ...(viewports ? { viewports } : {}) };
}

/** A whole page scaled to fit this many pixels on its longer side, unless the call says otherwise. */
const DEFAULT_IMAGE_SIZE = 1600;
const MAX_IMAGE_SIZE = 4096;
const DEFAULT_IMAGE_DPI = 200;
const MIN_IMAGE_DPI = 18;
const MAX_IMAGE_DPI = 600;

const DEFAULT_TEXT_LIMIT = 1000;
const MAX_TEXT_LIMIT = 5000;

export function finite(raw: unknown, name: string): number {
  const n = typeof raw === 'number' ? raw : Number(raw);
  if (!Number.isFinite(n)) throw new Error(`${name} has to be a number.`);
  return n;
}

/** The page the call names, the one on screen when it names none. */
export function pageOf(doc: AssistantDocument, args: Record<string, unknown>): { pageIndex: number; size: { width: number; height: number } } {
  const pageIndex = args.page === undefined ? doc.pageIndex : pageIndexOf(doc, args.page);
  if (pageIndex < 0 || pageIndex >= doc.pageCount) throw new Error(`Page ${pageIndex + 1} is not in ${doc.name} (1–${doc.pageCount}).`);
  const size = doc.pageSize?.(pageIndex) ?? null;
  if (!size) throw new Error(`Page size is not available for page ${pageIndex + 1} of ${doc.name}.`);
  return { pageIndex, size };
}

/** The region asked for, clipped to the page; the whole page when the call gives none. */
export function regionOf(raw: unknown, size: { width: number; height: number }): PageRegion {
  if (raw === undefined || raw === null) return { x: 0, y: 0, w: size.width, h: size.height };
  if (typeof raw !== 'object' || Array.isArray(raw)) throw new Error('Region has to be x, y, w, and h in PDF points.');
  const r = raw as Record<string, unknown>;
  const x = finite(r.x, 'Region x');
  const y = finite(r.y, 'Region y');
  const w = finite(r.w, 'Region w');
  const h = finite(r.h, 'Region h');
  if (w <= 0 || h <= 0) throw new Error('Region w and h have to be more than 0.');
  const x0 = Math.max(0, x);
  const y0 = Math.max(0, y);
  const x1 = Math.min(size.width, x + w);
  const y1 = Math.min(size.height, y + h);
  if (x1 <= x0 || y1 <= y0) throw new Error(`That region is off the page (${Math.round(size.width)} × ${Math.round(size.height)} points).`);
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

export function bounded(raw: unknown, name: string, min: number, max: number, fallback: number): number {
  if (raw === undefined) return fallback;
  const n = finite(raw, name);
  if (n < min || n > max) throw new Error(`${name} has to be between ${min} and ${max}.`);
  return n;
}

export const round1 = (n: number) => Math.round(n * 10) / 10;

async function getPageImage(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (!doc.render) throw new Error(`${doc.name} cannot be drawn here.`);
  const { pageIndex, size } = pageOf(doc, args);
  const region = regionOf(args.region, size);
  const dpi = bounded(args.dpi, 'dpi', MIN_IMAGE_DPI, MAX_IMAGE_DPI, DEFAULT_IMAGE_DPI);
  const maxSize = bounded(args.maxSize, 'maxSize', 256, MAX_IMAGE_SIZE, DEFAULT_IMAGE_SIZE);
  const scale = Math.min(dpi / 72, maxSize / Math.max(region.w, region.h));
  const withMarkups = args.markups !== false;
  const image = await doc.render(pageIndex, region, scale, withMarkups);
  if (!image) throw new Error(`That region is off page ${pageIndex + 1} of ${doc.name}.`);
  return {
    document: doc.name,
    id: doc.id,
    page: pageIndex + 1,
    region: { x: round1(image.region.x), y: round1(image.region.y), w: round1(image.region.w), h: round1(image.region.h) },
    width: image.width,
    height: image.height,
    scale: image.scale,
    dpi: round1(image.scale * 72),
    markups: withMarkups,
    image: { mimeType: 'image/png', data: image.png },
  };
}

function inRegion(word: TextWord, region: PageRegion): boolean {
  const cx = (word.x0 + word.x1) / 2;
  const cy = (word.y0 + word.y1) / 2;
  return cx >= region.x && cx <= region.x + region.w && cy >= region.y && cy <= region.y + region.h;
}

/** Reading order for a drawing: by line (top to bottom), then left to right within one. */
function readingOrder(a: TextWord, b: TextWord): number {
  const tolerance = Math.min(a.y1 - a.y0, b.y1 - b.y0) / 2;
  if (Math.abs(a.y0 - b.y0) > tolerance) return a.y0 - b.y0;
  return a.x0 - b.x0;
}

async function getPageText(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  if (!doc.text) throw new Error(`${doc.name} has no text here.`);
  const { pageIndex, size } = pageOf(doc, args);
  let words = await doc.text(pageIndex);
  if (args.region !== undefined && args.region !== null) {
    const region = regionOf(args.region, size);
    words = words.filter((w) => inRegion(w, region));
  }
  if (typeof args.search === 'string' && args.search.trim()) {
    const wanted = args.search.trim().toLowerCase();
    words = words.filter((w) => w.text.toLowerCase().includes(wanted));
  }
  words = [...words].sort(readingOrder);
  const limit = args.limit === undefined ? DEFAULT_TEXT_LIMIT : Math.min(wholeNumber(args.limit, 'Limit', 1), MAX_TEXT_LIMIT);
  const offset = args.offset === undefined ? 0 : wholeNumber(args.offset, 'Offset', 0);
  const shown = words.slice(offset, offset + limit);
  const next = offset + shown.length;
  return {
    document: doc.name,
    id: doc.id,
    page: pageIndex + 1,
    total: words.length,
    offset,
    ...(next < words.length ? { nextOffset: next } : {}),
    words: shown.map((w) => ({ text: w.text, x0: round1(w.x0), y0: round1(w.y0), x1: round1(w.x1), y1: round1(w.y1), size: round1(w.size) })),
  };
}

function listCommands(params: unknown, env: AssistantEnv) {
  const commands = env.commands?.();
  if (!commands) throw new Error('Commands are not available.');
  const categoryRaw = record(params).category;
  const category = typeof categoryRaw === 'string' ? categoryRaw.trim().toLowerCase() : '';
  const queryRaw = record(params).query;
  const query = typeof queryRaw === 'string' ? queryRaw.trim().toLowerCase() : '';
  const onlyAvailable = record(params).available === true;
  const list = [...commands.values()].filter(
    (command) =>
      (!category || command.category.toLowerCase() === category) &&
      (!onlyAvailable || command.enabled) &&
      (!query || command.id.toLowerCase().includes(query) || command.label.toLowerCase().includes(query)),
  );
  return {
    commands: list.map((command) => ({
      id: command.id,
      label: command.label,
      category: command.category,
      enabled: command.enabled,
      ...(command.checked === undefined ? {} : { checked: command.checked }),
    })),
  };
}

function commandNeedsYou(command: AssistantCommand): boolean {
  return command.label.endsWith('…') || command.label.endsWith('...') || NEEDS_YOU.has(command.id);
}

/** While a command runs, its own confirm is already answered by Allow. A confirm after an await still appears. */
function acceptConfirms(run: () => void) {
  const host = globalThis as { confirm?: (message?: string) => boolean };
  const previous = host.confirm;
  host.confirm = () => true;
  try {
    run();
  } finally {
    if (previous) host.confirm = previous;
    else delete host.confirm;
  }
}

function activeName(docs: readonly AssistantDocument[]): string | undefined {
  const open = assistantDocuments(docs);
  return (open.find((doc) => doc.active) ?? open[0])?.name;
}

async function runCommand(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const commands = env.commands?.();
  if (!commands) throw new Error('Commands are not available.');
  const id = record(params).id;
  if (typeof id !== 'string' || !id) throw new Error('Say which command to run. list_commands has the ids.');
  const command = commands.get(id);
  if (!command) throw new Error(`No command ${id}.`);
  const action = command.label.replace(/…$/, '').replace(/\.\.\.$/, '');
  if (!command.enabled) throw new Error(`${action} is not available.`);
  if (isDestructiveCommand(command.id)) {
    const name = activeName(docs);
    await ask(env, name ? `Allow the assistant to ${action} on ${name}?` : `Allow the assistant to ${action}?`);
  }
  acceptConfirms(() => command.run());
  if (commandNeedsYou(command)) return { started: true, needsYou: true };
  return { ok: true, id: command.id };
}

/** Runs one assistant tool against the drawings that are open right now. */
// ---- Takeoff proposals: the AI suggests, a person approves (see assemblies/proposals.ts) ----------

function listAssemblies(docs: readonly AssistantDocument[], params: unknown) {
  const doc = findDocument(docs, record(params).document);
  return {
    document: doc.name,
    assemblies: doc.store.assemblies().map((a) => ({
      id: a.id,
      name: a.name,
      ...(a.description ? { description: a.description } : {}),
      variables: a.variables,
      items: a.items.map((i) => ({ name: i.name, unit: i.unit, formula: i.formula, ...(i.waste ? { waste: i.waste } : {}), ...(i.roundUp ? { roundUp: true } : {}), ...(i.roundUp && i.roundTo ? { roundTo: i.roundTo } : {}), ...(i.unitCost !== undefined ? { unitCost: i.unitCost } : {}) })),
    })),
    formulaNames: MEASURED_NAMES.map((n) => n.name),
  };
}

export function assemblyOf(doc: AssistantDocument, raw: unknown) {
  if (typeof raw !== 'string' || !raw.trim()) throw new Error('Say which assembly: its name or id from list_assemblies.');
  const wanted = raw.trim().toLowerCase();
  const found = doc.store.assemblies().find((a) => a.id === raw || a.name.trim().toLowerCase() === wanted);
  if (!found) throw new Error(`${doc.name} has no assembly "${raw}". list_assemblies lists them.`);
  return found;
}

/** Proposals go to Takeoff Review; nothing reaches the takeoff until a person approves each one. */
function proposeTakeoff(docs: readonly AssistantDocument[], params: unknown) {
  const args = record(params);
  const doc = findDocument(docs, args.document);
  const items = batchItems(args.proposals, 'Give proposals: a list of { id or seq, assembly, reasoning }.');
  const source = typeof args.source === 'string' && args.source.trim() ? args.source.trim().slice(0, 80) : 'AI assistant';
  const now = Date.now();
  // Every proposal is checked before any is kept.
  const proposals: TakeoffProposal[] = items.map((item, n) => {
    try {
      const markup = findMarkup(doc, item);
      if (!isMeasureKind(markup.type)) throw new Error(`markup ${markup.seq ?? markup.id} is a ${markup.type}, not a measurement`);
      const assembly = assemblyOf(doc, item.assembly);
      const reasoning = typeof item.reasoning === 'string' ? item.reasoning.trim().slice(0, 2000) : '';
      if (!reasoning) throw new Error('say why (reasoning), so the person can check it');
      const expected = Array.isArray(item.expected)
        ? item.expected.flatMap((e: unknown) => {
            const r = record(e);
            const quantity = typeof r.quantity === 'number' ? r.quantity : Number(r.quantity);
            return typeof r.item === 'string' && Number.isFinite(quantity) ? [{ item: r.item, quantity, ...(typeof r.unit === 'string' ? { unit: r.unit } : {}) }] : [];
          })
        : undefined;
      return { id: crypto.randomUUID(), markupId: markup.id, assemblyId: assembly.id, reasoning, ...(expected?.length ? { expected } : {}), source, createdAt: now };
    } catch (err) {
      throw new Error(`Proposal ${n + 1}: ${errorText(err)}`);
    }
  });
  takeoffProposals.add(doc.store.fileHash, proposals);
  return {
    document: doc.name,
    proposed: proposals.length,
    ids: proposals.map((p) => p.id),
    note: 'Waiting for review: the person approves, edits or rejects each one in Tools › Takeoff › Review AI Proposals. Nothing is added to the takeoff until they approve it.',
  };
}

function listTakeoffProposals(docs: readonly AssistantDocument[], params: unknown) {
  const doc = findDocument(docs, record(params).document);
  const assemblies = doc.store.assemblies();
  return {
    document: doc.name,
    proposals: takeoffProposals.list(doc.store.fileHash).map((p) => {
      const m = doc.store.get(p.markupId);
      return {
        id: p.id,
        markup: m ? { id: m.id, ...(m.seq !== undefined ? { seq: m.seq } : {}), type: m.type, page: m.pageIndex + 1 } : null,
        assembly: assemblies.find((a) => a.id === p.assemblyId)?.name ?? null,
        state: proposalState(p, m),
        ...(m?.takeoffApproval?.proposalId === p.id ? { approvedBy: m.takeoffApproval.by, edited: !!m.takeoffApproval.edited } : {}),
        ...(p.rejected ? { rejectedBy: p.rejected.by } : {}),
      };
    }),
  };
}

// ---- Semantic compare: the person asks in the panel, the AI judges, the person reads --------------

/** Longest paragraph text one call returns. A paragraph past this is cut, with an ellipsis. */
const MAX_PARAGRAPH_CHARS = 4000;
const DEFAULT_PARAGRAPHS = 100;
const MAX_PARAGRAPHS = 400;

/**
 * Reads two open documents' text into paragraphs and starts a comparison, with the text comparison
 * ready at once. The panel's Compare button and `start_comparison` both come here.
 */
export async function beginComparison(docs: readonly AssistantDocument[], oldId: string, newId: string): Promise<void> {
  const oldDoc = docs.find((d) => d.id === oldId);
  const newDoc = docs.find((d) => d.id === newId);
  if (!oldDoc?.text || !newDoc?.text) throw new Error('Open both revisions in tabs first: their text is read from the open documents.');
  const [old, now] = await Promise.all([documentParagraphs('o', oldDoc.pageCount, oldDoc.text), documentParagraphs('n', newDoc.pageCount, newDoc.text)]);
  if (!old.length && !now.length) throw new Error('Neither document has text to compare. Scanned pages need OCR first (Document › OCR).');
  semanticCompare.begin({ oldId, newId, oldName: oldDoc.name, newName: newDoc.name, old, new: now });
}

/** The assistant starts a comparison: the person says Allow first, because both documents' text goes to the AI client. */
async function startComparison(docs: readonly AssistantDocument[], params: unknown, env: AssistantEnv) {
  const args = record(params);
  if (typeof args.oldDocument !== 'string' || typeof args.newDocument !== 'string') throw new Error('Say which two documents: oldDocument and newDocument, by name or id (list_open_documents).');
  const oldDoc = findDocument(docs, args.oldDocument);
  const newDoc = findDocument(docs, args.newDocument);
  if (oldDoc.id === newDoc.id) throw new Error('Choose two different documents.');
  await ask(env, `Allow the assistant to compare ${oldDoc.name} (old) with ${newDoc.name} (new)? The text of both documents goes to your AI client.`);
  await beginComparison(docs, oldDoc.id, newDoc.id);
  env.revealComparison?.();
  const state = semanticCompare.get()!;
  return { ok: true, requestId: state.request.id, old: oldDoc.name, new: newDoc.name, note: 'Started. The person sees a text comparison now. Call get_comparison, then report_document_changes with your verdicts.' };
}

function openComparison() {
  const state = semanticCompare.get();
  if (!state) throw new Error('No comparison is waiting. Start one with start_comparison, or ask the person to open Document › AI Compare Documents… and press Compare.');
  return state;
}

function leftOver(state: SemanticState) {
  return exactMatches(state.request.old, state.request.new);
}

function clipped(p: Paragraph) {
  return { id: p.id, page: p.page, text: p.text.length > MAX_PARAGRAPH_CHARS ? `${p.text.slice(0, MAX_PARAGRAPH_CHARS)}…` : p.text };
}

/**
 * What the person asked to have compared. Paragraphs with the same text were matched here already
 * (moved or not), so only the rest come back: what the AI is needed for, and a fraction of a long report.
 */
function getComparison(params: unknown) {
  const args = record(params);
  const state = openComparison();
  const { pairs, unmatchedOld, unmatchedNew } = leftOver(state);
  const reported = new Set(state.decisions.flatMap((d) => [d.old, d.new].filter((x): x is string => !!x)));
  const openOld = unmatchedOld.filter((p) => !reported.has(p.id));
  const openNew = unmatchedNew.filter((p) => !reported.has(p.id));
  const limit = args.limit === undefined ? DEFAULT_PARAGRAPHS : Math.min(wholeNumber(args.limit, 'Limit', 1), MAX_PARAGRAPHS);
  const offset = args.offset === undefined ? 0 : wholeNumber(args.offset, 'Offset', 0);
  const old = openOld.slice(offset, offset + limit);
  const now = openNew.slice(offset, offset + limit);
  const more = offset + limit < Math.max(openOld.length, openNew.length);
  return {
    requestId: state.request.id,
    old: { name: state.request.oldName, paragraphs: state.request.old.length },
    new: { name: state.request.newName, paragraphs: state.request.new.length },
    sameText: { matched: pairs.length },
    unmatched: { old: openOld.length, new: openNew.length },
    offset,
    ...(more ? { nextOffset: offset + limit } : {}),
    oldParagraphs: old.map(clipped),
    newParagraphs: now.map(clipped),
    instructions:
      'Match old and new paragraphs by meaning, not position. For each pair call report_document_changes with kind modified (reworded or edited), moved (the same content elsewhere, even if lightly reworded) or unchanged (same meaning, only reformatted). A new paragraph with no old counterpart is added; an old one with no new counterpart is removed. Paragraphs that have the same text were matched already and are not listed.',
  };
}

const DECISION_KINDS = new Set<ChangeKind>(['added', 'removed', 'modified', 'moved', 'unchanged']);

/** An AI client's verdicts on the paragraphs `get_comparison` listed. All are checked before any is kept. */
function reportDocumentChanges(params: unknown) {
  const args = record(params);
  const state = openComparison();
  if (typeof args.requestId === 'string' && args.requestId && args.requestId !== state.request.id) throw new Error('That comparison is no longer open: the person started another. Call get_comparison again.');
  const source = typeof args.source === 'string' && args.source.trim() ? args.source.trim().slice(0, 80) : 'AI assistant';
  if (typeof args.error === 'string' && args.error.trim()) {
    semanticCompare.fail(state.request.id, args.error.trim().slice(0, 500), source);
    return { ok: true, note: 'Noted. redcolumn shows its own text comparison and tells the person you could not judge it.' };
  }
  const { unmatchedOld, unmatchedNew } = leftOver(state);
  const openOld = new Set(unmatchedOld.map((p) => p.id));
  const openNew = new Set(unmatchedNew.map((p) => p.id));
  const items = batchItems(args.changes, 'Give changes: a list of { kind, old and/or new (paragraph ids from get_comparison), note }.');
  const seen = new Set<string>();
  const decisions: AiDecision[] = items.map((item, n) => {
    try {
      const kind = item.kind;
      if (typeof kind !== 'string' || !DECISION_KINDS.has(kind as ChangeKind)) throw new Error('kind has to be added, removed, modified, moved, or unchanged');
      const old = typeof item.old === 'string' && item.old ? item.old : undefined;
      const now = typeof item.new === 'string' && item.new ? item.new : undefined;
      if (kind === 'added' && (!now || old)) throw new Error('added takes the new paragraph only');
      if (kind === 'removed' && (!old || now)) throw new Error('removed takes the old paragraph only');
      if (kind !== 'added' && kind !== 'removed' && (!old || !now)) throw new Error(`${kind} takes both the old and the new paragraph`);
      for (const [id, open, side] of [[old, openOld, 'old'], [now, openNew, 'new']] as const) {
        if (!id) continue;
        if (!(id.startsWith(side === 'old' ? 'o' : 'n') && state.request[side].some((p) => p.id === id))) throw new Error(`${id} is not a paragraph of the ${side} revision`);
        if (!open.has(id)) throw new Error(`${id} has the same text in both revisions, so redcolumn matched it already`);
        if (seen.has(id)) throw new Error(`${id} is in the list twice`);
        seen.add(id);
      }
      const note = typeof item.note === 'string' && item.note.trim() ? item.note.trim().slice(0, 500) : undefined;
      return { kind: kind as ChangeKind, ...(old ? { old } : {}), ...(now ? { new: now } : {}), ...(note ? { note } : {}) };
    } catch (err) {
      throw new Error(`Change ${n + 1}: ${errorText(err)}`);
    }
  });
  const summary = typeof args.summary === 'string' && args.summary.trim() ? args.summary.trim().slice(0, 1000) : undefined;
  semanticCompare.decide(state.request.id, decisions, source, summary);
  const after = semanticCompare.get()!;
  const decided = new Set(after.decisions.flatMap((d) => [d.old, d.new].filter((x): x is string => !!x)));
  return {
    ok: true,
    accepted: decisions.length,
    stillUnreported: { old: unmatchedOld.filter((p) => !decided.has(p.id)).length, new: unmatchedNew.filter((p) => !decided.has(p.id)).length },
    note: 'Shown to the person in AI Compare Documents, marked as your judgement. Paragraphs you did not report are compared by text similarity and marked so.',
  };
}

export async function assistantCall(docs: readonly AssistantDocument[], method: string, params: unknown, env: AssistantEnv = {}): Promise<unknown> {
  if (method === 'list_open_documents') return { documents: assistantDocuments(docs).map(publicDocument) };
  if (method === 'list_markups') return listMarkups(docs, params);
  if (method === 'set_markup') return setMarkup(docs, params);
  if (method === 'set_markups') return setMarkups(docs, params);
  if (method === 'get_page') return getPage(docs, params);
  if (method === 'get_page_image') return getPageImage(docs, params);
  if (method === 'get_page_text') return getPageText(docs, params);
  if (method === 'place_markup') return placeMarkup(docs, params, env);
  if (method === 'place_markups') return placeMarkups(docs, params, env);
  if (method === 'update_markup') return updateMarkup(docs, params);
  if (method === 'update_markups') return updateMarkups(docs, params);
  if (method === 'delete_markups') return deleteMarkups(docs, params, env);
  if (method === 'select_markups') return selectMarkups(docs, params, env);
  if (method === 'list_commands') return listCommands(params, env);
  if (method === 'run_command') return runCommand(docs, params, env);
  if (method === 'list_assemblies') return listAssemblies(docs, params);
  if (method === 'propose_takeoff') return proposeTakeoff(docs, params);
  if (method === 'list_takeoff_proposals') return listTakeoffProposals(docs, params);
  if (method === 'start_comparison') return startComparison(docs, params, env);
  if (method === 'get_comparison') return getComparison(params);
  if (method === 'report_document_changes') return reportDocumentChanges(params);
  if ((STITCH_METHODS as readonly string[]).includes(method)) return stitchCall(docs, method, params, env);
  if ((VIEWPORT_METHODS as readonly string[]).includes(method)) return viewportCall(docs, method, params);
  if ((SCALE_METHODS as readonly string[]).includes(method)) return scaleCall(docs, method as ScaleMethod, params);
  if ((STANDARD_METHODS as readonly string[]).includes(method)) return standardCall(docs, method as StandardMethod, params, env.standards);
  if ((ASSEMBLY_PROPOSAL_METHODS as readonly string[]).includes(method)) return assemblyProposalCall(docs, method as AssemblyProposalMethod, params);
  throw new Error('Unknown tool.');
}
