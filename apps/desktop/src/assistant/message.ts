/** What a tool says when the desktop app is closed or the assistant is off. */
export const OFFLINE = 'Open redcolumn and turn on the local assistant.';

/** How long the page may take to answer. Long enough for someone to Allow or Deny. */
export const PAGE_WAIT_MS = 130_000;

/** The loopback client waits slightly longer than the page, so a slow answer is not "the app is closed". */
export const CLIENT_WAIT_MS = 140_000;

export const ASSISTANT_METHODS = [
  'list_open_documents',
  'list_markups',
  'set_markup',
  'set_markups',
  'list_commands',
  'run_command',
  'get_page',
  'get_page_image',
  'get_page_text',
  'select_markups',
  'place_markup',
  'place_markups',
  'update_markup',
  'update_markups',
  'delete_markups',
  'list_assemblies',
  'propose_takeoff',
  'list_takeoff_proposals',
  'propose_assembly',
  'list_assembly_proposals',
  'start_comparison',
  'get_comparison',
  'report_document_changes',
  'get_stitching',
  'set_stitch_region',
  'place_match_lines',
  'update_match_lines',
  'delete_match_lines',
  'get_stitch_image',
  'show_stitched_view',
  'stitch_to_page',
  'list_viewports',
  'place_viewport',
  'place_viewports',
  'set_viewport_scale',
  'get_sheet_info',
  'propose_page_scale',
  'get_drawing_standard',
  'list_sheet_links',
  'propose_sheet_links',
  'propose_standard_hints',
] as const;

export type AssistantMethod = (typeof ASSISTANT_METHODS)[number];

export function isAssistantMethod(method: string): method is AssistantMethod {
  return (ASSISTANT_METHODS as readonly string[]).includes(method);
}

/**
 * MCP tool hints. A client treats a tool that is not read-only as destructive unless it says
 * otherwise, so every tool states all of these. Reads can then run without a prompt, and the
 * tools that only add or change things (each one undo step in redcolumn) are not flagged.
 */
interface ToolAnnotations {
  readOnlyHint: boolean;
  destructiveHint: boolean;
  idempotentHint?: boolean;
  openWorldHint: false;
}

const READS: ToolAnnotations = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };
const ADDS: ToolAnnotations = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false };
const CHANGES: ToolAnnotations = { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false };
const REMOVES: ToolAnnotations = { readOnlyHint: false, destructiveHint: true, openWorldHint: false };

interface ToolDef {
  name: AssistantMethod;
  title: string;
  description: string;
  inputSchema: Record<string, unknown>;
  annotations: ToolAnnotations;
}

const documentProp = { type: 'string', description: 'Document name or id. The active document when omitted.' };

const pointsProp = {
  type: 'array',
  description: 'PDF points, origin at the top-left, y down.',
  items: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
};

const scaleProp = {
  type: 'string',
  description: "The scale as get_page writes it, such as 1/4\" = 1'-0\" or 1:100.",
};

const styleProp = {
  type: 'object',
  properties: {
    stroke: { type: 'string', description: 'CSS hex color, such as #ff0000.' },
    fill: { type: ['string', 'null'], description: 'CSS hex color, or null for none.' },
    width: { type: 'number', description: 'Line width in PDF points.' },
    opacity: { type: 'number', description: '0 to 1.' },
    fontSize: { type: 'number', description: 'Font size in points.' },
  },
};

const regionProp = {
  type: 'object',
  description: 'Part of the page, in PDF points from the top-left: x, y, w, h. The whole page when omitted.',
  properties: {
    x: { type: 'number' },
    y: { type: 'number' },
    w: { type: 'number' },
    h: { type: 'number' },
  },
  required: ['x', 'y', 'w', 'h'],
};

const setFields = {
  status: { type: 'string', description: 'Status name or id, such as Accepted or completed.' },
  comment: { type: 'string' },
  subject: { type: 'string' },
  field: { type: 'string', description: 'Custom column name or id.' },
  value: { type: 'string', description: 'Value for field.' },
};

const matchLineProp = {
  type: 'object',
  properties: {
    page: { type: 'integer', description: '1-based page.' },
    p0: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2, description: 'One end, in PDF points from the top-left, y down.' },
    p1: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2, description: 'The other end.' },
    keepSide: { type: 'string', enum: ['right', 'left'], description: 'The side of the line that stays, walking from p0 to p1.' },
    keepPoint: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2, description: 'A point inside the part of the sheet that stays.' },
  },
  required: ['page', 'p0', 'p1'],
};

const TOOLS: ToolDef[] = [
  {
    name: 'list_open_documents',
    title: 'List open drawings',
    description: "Lists drawings open in redcolumn: name, page count, the page on screen, and that page's sheet number.",
    inputSchema: { type: 'object', properties: {} },
    annotations: READS,
  },
  {
    name: 'list_markups',
    title: 'List markups',
    description:
      'Lists markups on an open drawing, with type, page, text, status, and measurement. Filter with page, type, status, or search, and page through with limit and offset (500 markups per call when limit is omitted; the reply has total, and nextOffset while more remain). detail "brief" leaves out points, style, and author. summary true returns only counts by type, status, and page. On a big drawing, start with summary or a filter.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number. Every page when omitted.' },
        type: { type: 'string', description: 'Only this markup type, such as rect, cloud, or length.' },
        status: { type: 'string', description: 'Only markups with this status name or id.' },
        search: { type: 'string', description: 'Only markups whose text, subject, or comment contains this, ignoring case.' },
        detail: { type: 'string', enum: ['full', 'brief'], description: 'brief leaves out points, style, and author. full when omitted.' },
        summary: { type: 'boolean', description: 'Counts by type, status, and page instead of the markups.' },
        limit: { type: 'integer', minimum: 1, maximum: 2000, description: 'Most markups to return. 500 when omitted.' },
        offset: { type: 'integer', minimum: 0, description: 'Markups to skip, from nextOffset of the last reply.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'set_markup',
    title: 'Set markup status or comment',
    description: "Updates one markup's status, comment, subject, or a custom column. The change is one undo step in redcolumn. For several markups use set_markups.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        id: { type: 'string', description: 'Markup id from list_markups.' },
        seq: { type: 'integer', description: 'Markup number shown in the list, when id is omitted.' },
        ...setFields,
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'set_markups',
    title: 'Set status or comment on many markups',
    description:
      'Gives many markups the same status, comment, subject, or custom column value, as one undo step in redcolumn. Name them with ids, seqs, or both (up to 500). Everything is checked first, so a bad status changes nothing. Markups that redcolumn then refuses to change are listed in failed.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        ids: { type: 'array', items: { type: 'string' }, description: 'Markup ids from list_markups.' },
        seqs: { type: 'array', items: { type: 'integer' }, description: 'Markup numbers shown in the list.' },
        ...setFields,
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'list_commands',
    title: 'List commands',
    description: 'Lists commands the menus and the command palette can run, with whether each is available right now. Narrow it with category, query, or available: the full list is long.',
    inputSchema: {
      type: 'object',
      properties: {
        category: { type: 'string', description: 'File, Edit, View, Document, Bulk, Tools, Window, or Help. Every category when omitted.' },
        query: { type: 'string', description: 'Only commands whose id or label contains this, ignoring case, such as zoom or rotate.' },
        available: { type: 'boolean', description: 'true leaves out commands that cannot run right now.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'run_command',
    title: 'Run a command',
    description: 'Runs one command from list_commands. Commands that open a dialog return immediately and wait for the person in redcolumn. Deletes, saves, and document changes wait for Allow in redcolumn before they run.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Command id from list_commands, such as edit.delete or view.zoomIn.' },
      },
    },
    annotations: REMOVES,
  },
  {
    name: 'get_page',
    title: 'Get page size and scale',
    description: 'The size of one page in PDF points, its scale when one is set, and how many viewports it has (list them with list_viewports). Point origin is the top-left of the page, y down.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number. The page on screen when omitted.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'get_page_image',
    title: 'See a page',
    description:
      'A picture of a page, or part of one, as the drawing shows it. The reply has the image and, as text, its scale in pixels per PDF point and the region it covers, so a pixel at (px, py) is the page point (region.x + px / scale, region.y + py / scale). The whole page is scaled to fit maxSize, so small linework and lettering blur: look at the whole page first to find what you need, then ask for a region a few hundred points across. Markups are drawn over the page unless markups is false.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number. The page on screen when omitted.' },
        region: regionProp,
        dpi: { type: 'number', minimum: 18, maximum: 600, description: 'Resolution in pixels per inch (72 = one pixel per point). The picture is still kept within maxSize, so a whole sheet may come back at less than this. 200 when omitted.' },
        maxSize: { type: 'integer', minimum: 256, maximum: 4096, description: 'Most pixels on the longer side. 1600 when omitted.' },
        markups: { type: 'boolean', description: 'false leaves the markups off, showing the drawing alone. true when omitted.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'get_page_text',
    title: 'Read page text',
    description:
      "The text on a page as the PDF holds it, word by word with each word's box in PDF points (x0, y0 top-left; x1, y1 bottom-right) and font size, in reading order. Narrow it with region or search. Scanned pages have no text until they are OCRed (document.ocr). 1000 words per call when limit is omitted; the reply has total, and nextOffset while more remain.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number. The page on screen when omitted.' },
        region: regionProp,
        search: { type: 'string', description: 'Only words containing this, ignoring case.' },
        limit: { type: 'integer', minimum: 1, maximum: 5000, description: 'Most words to return. 1000 when omitted.' },
        offset: { type: 'integer', minimum: 0, description: 'Words to skip, from nextOffset of the last reply.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'select_markups',
    title: 'Select markups',
    description: 'Selects markups on the drawing that is on screen, so a later run_command (align, group, copy) acts on them. An empty list clears the selection.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        ids: { type: 'array', items: { type: 'string' }, description: 'Markup ids from list_markups.' },
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'place_markup',
    title: 'Draw a markup',
    description:
      'Draws one markup on a page. One undo step. Image, stamp, signature, and file attachment need a file: run that tool with run_command instead. To draw several, use place_markups. Measurements of length, area and volume need a page (or viewport) with a scale: on a page without one, or marked NOT TO SCALE, they are refused (see get_sheet_info and propose_page_scale).',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number.' },
        type: {
          type: 'string',
          description:
            'Markup type, such as rect, cloud, length, or text. Also cloudPlus (a cloud with a callout; text goes on the callout) and dynamicFill (Smart Fill: one point inside a closed outline makes the fill type chosen in redcolumn, area by default). Both need the drawing on screen.',
        },
        points: {
          type: 'array',
          description: 'PDF points, origin at the top-left, y down. A line, rectangle, ellipse, cloud, cloudPlus, text box, or redaction is two opposite corners. A callout is the tip, the knee, and the box\'s two corners. A pen or highlight is the path. A polygon, area, or space is its vertices.',
          items: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
        },
        text: { type: 'string' },
        comment: { type: 'string' },
        subject: { type: 'string' },
      },
    },
    annotations: ADDS,
  },
  {
    name: 'place_markups',
    title: 'Draw many markups',
    description:
      'Draws many markups as one undo step in redcolumn (up to 500). Each item is what place_markup takes: type, points, and optionally page, text, comment, and subject; a page on the call is the default for items without one. Everything is checked first, so one bad item draws nothing. Not for cloudPlus or dynamicFill, which need place_markup. The reply lists each new markup\'s id, seq, type, and page.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number for items that do not give their own.' },
        markups: {
          type: 'array',
          minItems: 1,
          maxItems: 500,
          items: {
            type: 'object',
            properties: {
              page: { type: 'integer', description: '1-based page number.' },
              type: { type: 'string', description: 'Markup type, such as rect, cloud, length, or text.' },
              points: { ...pointsProp, description: 'PDF points, as for place_markup: two opposite corners for a line, rectangle, ellipse, cloud, text box, or redaction; the path for a pen or highlight; the vertices for a polygon, area, or space.' },
              text: { type: 'string' },
              comment: { type: 'string' },
              subject: { type: 'string' },
            },
            required: ['type', 'points'],
          },
        },
      },
      required: ['markups'],
    },
    annotations: ADDS,
  },
  {
    name: 'update_markup',
    title: 'Edit a markup',
    description: "Changes one markup's points, text, or style (stroke, fill, width, opacity, font size). Locked markups stay locked. One undo step. Status, comment, subject, and columns stay on set_markup. For several markups use update_markups.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        id: { type: 'string', description: 'Markup id from list_markups.' },
        seq: { type: 'integer', description: 'Markup number shown in the list, when id is omitted.' },
        points: {
          type: 'array',
          description: 'Replacement points, in the same order place_markup uses for this type.',
          items: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
        },
        text: { type: 'string', description: 'Replacement text. Empty clears it.' },
        style: styleProp,
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'update_markups',
    title: 'Edit many markups',
    description:
      'Changes the points, text, or style of many markups as one undo step in redcolumn (up to 500). Each item is what update_markup takes, with an id or seq. Everything is checked first, so one bad item changes nothing. Markups that redcolumn then refuses to change are listed in failed.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        markups: {
          type: 'array',
          minItems: 1,
          maxItems: 500,
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Markup id from list_markups.' },
              seq: { type: 'integer', description: 'Markup number shown in the list, when id is omitted.' },
              points: { ...pointsProp, description: 'Replacement points, in the same order place_markup uses for this type.' },
              text: { type: 'string', description: 'Replacement text. Empty clears it.' },
              style: styleProp,
            },
          },
        },
      },
      required: ['markups'],
    },
    annotations: CHANGES,
  },
  {
    name: 'delete_markups',
    title: 'Delete markups',
    description: 'Deletes markups by id. redcolumn asks for Allow first. One undo step after Allow.',
    annotations: REMOVES,
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        ids: { type: 'array', items: { type: 'string' }, description: 'Markup ids from list_markups.' },
      },
    },
  },
  {
    name: 'list_assemblies',
    title: 'List takeoff assemblies',
    description:
      "Lists the drawing's takeoff assemblies: each one's items with their units, formulas and unit costs, and its variables. formulaNames are the measured values a formula can use ([Length], [Area], [Count]...). Use this before propose_takeoff and propose_assembly.",
    inputSchema: { type: 'object', properties: { document: documentProp } },
    annotations: READS,
  },
  {
    name: 'propose_takeoff',
    title: 'Propose takeoff assemblies',
    description:
      'Proposes an assembly for each of some measurements (from list_markups), with your reasoning and, optionally, the quantities you expect. Proposals do not change the drawing: they wait in Takeoff Review, where the person approves, edits or rejects each one. Nothing is added to the takeoff until they approve it. Up to 500 per call; a new proposal for a measurement replaces its earlier pending one.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        source: { type: 'string', description: 'Your name, shown to the person as who proposed it.' },
        proposals: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Measurement markup id from list_markups.' },
              seq: { type: 'integer', description: 'Measurement number, when id is omitted.' },
              assembly: { type: 'string', description: 'Assembly name or id from list_assemblies.' },
              reasoning: { type: 'string', description: 'Why this assembly fits this measurement, for the person checking it.' },
              expected: {
                type: 'array',
                description: 'The quantities you expect, checked against what the assembly works out.',
                items: { type: 'object', properties: { item: { type: 'string' }, quantity: { type: 'number' }, unit: { type: 'string' } }, required: ['item', 'quantity'] },
              },
            },
            required: ['assembly', 'reasoning'],
          },
        },
      },
      required: ['proposals'],
    },
    annotations: ADDS,
  },
  {
    name: 'list_takeoff_proposals',
    title: 'List takeoff proposals',
    description: 'Lists takeoff proposals on a drawing and where each stands: pending, approved (and by whom, and whether they edited it) or rejected.',
    inputSchema: { type: 'object', properties: { document: documentProp } },
    annotations: READS,
  },
  {
    name: 'propose_assembly',
    title: 'Propose takeoff assemblies',
    description:
      'Proposes takeoff assemblies for the drawing: new ones, or edits of existing ones with replaces (their id or name). You can only propose: each assembly waits in Tools › Takeoff › Review AI Proposals, with your reasoning, until a person approves or rejects it. Approving writes the document; nothing changes until then. Each item is a quantity worked out by a formula over the measured values (formulaNames from list_assemblies) and the assembly\'s own variables. An edit keeps the assembly and its matching items\' identity, so measurements stay linked and keep any quantities set by hand. All assemblies are checked first; one bad item means none is kept.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        source: { type: 'string', description: 'Your name, shown to the person as who proposed it.' },
        assemblies: {
          type: 'array',
          minItems: 1,
          maxItems: 10,
          items: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'The assembly\'s name. It has to be free unless replaces names the assembly that has it.' },
              description: { type: 'string', description: 'What it measures and includes, shown with the assembly.' },
              variables: {
                type: 'array',
                maxItems: 20,
                description: 'Named numbers the formulas use, such as a stud spacing.',
                items: { type: 'object', properties: { name: { type: 'string' }, value: { type: 'number' } }, required: ['name', 'value'] },
              },
              items: {
                type: 'array',
                minItems: 1,
                maxItems: 40,
                description: 'The quantities the assembly produces for each measurement.',
                items: {
                  type: 'object',
                  properties: {
                    name: { type: 'string', description: 'The material or labour, e.g. Studs.' },
                    unit: { type: 'string', description: 'As shown: ea, sheets, lf, m².' },
                    formula: { type: 'string', description: 'e.g. [Length] / [Spacing] + 1 or [Wall Area] * [Layers] / 32.' },
                    waste: { type: 'number', description: 'Waste allowance in percent, 0 to 100, added before rounding.' },
                    roundUp: { type: 'boolean', description: 'Round the quantity up (sheets, boxes, lengths of pipe).' },
                    roundTo: { type: 'number', description: 'With roundUp, round up to a multiple of this instead of 1: 10 for pipe in 10 ft lengths.' },
                    unitCost: { type: 'number', description: 'Price per unit, for the cost. 0 or more.' },
                  },
                  required: ['name', 'unit', 'formula'],
                },
              },
              replaces: { type: 'string', description: 'The id or name of the existing assembly this proposes an edit of. Its items keep their identity where the item names match.' },
              reasoning: { type: 'string', description: 'Why these items and formulas, for the person deciding.' },
            },
            required: ['name', 'items', 'reasoning'],
          },
        },
      },
      required: ['assemblies'],
    },
    annotations: ADDS,
  },
  {
    name: 'list_assembly_proposals',
    title: 'List assembly proposals',
    description: 'Lists assembly proposals on a drawing and where each stands: pending (waiting for a person in Tools › Takeoff › Review AI Proposals) or rejected. Approved ones are written into the document, so list_assemblies has them.',
    inputSchema: { type: 'object', properties: { document: documentProp } },
    annotations: READS,
  },
  {
    name: 'start_comparison',
    title: 'Start a document comparison',
    description:
      'Starts an AI comparison of two open documents (old revision, new revision) in redcolumn. redcolumn asks the person for Allow first, because the text of both goes to you. The person then sees a text comparison at once in the AI Compare panel; you improve on it: call get_comparison to read the paragraphs that did not match word for word, then report_document_changes with your verdicts. Nothing in either document changes.',
    inputSchema: {
      type: 'object',
      properties: {
        oldDocument: { type: 'string', description: 'The old revision: name or id from list_open_documents.' },
        newDocument: { type: 'string', description: 'The new revision: name or id from list_open_documents.' },
      },
      required: ['oldDocument', 'newDocument'],
    },
    annotations: CHANGES,
  },
  {
    name: 'get_comparison',
    title: 'Read a document comparison',
    description:
      'The comparison in redcolumn that the person started (Document › AI Compare Documents…) or you started with start_comparison: two revisions of a document cut into paragraphs. Paragraphs with the same text in both revisions are matched already, moved or not, and only counted here. The rest come back with ids (o12 old, n7 new), pages and text: match them by meaning, not position, and answer with report_document_changes. 100 paragraphs per side per call when limit is omitted; the reply has nextOffset while more remain. Fails when no comparison has been started: use start_comparison.',
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: 'integer', minimum: 1, maximum: 400, description: 'Most paragraphs per side. 100 when omitted.' },
        offset: { type: 'integer', minimum: 0, description: 'Paragraphs to skip on each side, from nextOffset of the last reply.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'report_document_changes',
    title: 'Report document changes',
    description:
      'Gives your verdict on paragraphs from get_comparison. Each change is { kind, old, new, note }: modified (reworded or edited; old and new), moved (same content in another place, even if lightly reworded; old and new), unchanged (same meaning, only reformatted; old and new), added (new only) or removed (old only). Say why in note when it is not obvious. The person sees each one in redcolumn marked as your judgement, with the changed wording. Paragraphs you leave out are compared by text similarity and marked as that. Call it as often as you like, a page of paragraphs at a time; a later verdict on a paragraph replaces an earlier one. If you cannot judge the comparison, pass error with the reason instead of changes. Nothing in the document changes.',
    inputSchema: {
      type: 'object',
      properties: {
        requestId: { type: 'string', description: 'requestId from get_comparison, to be sure it is still the comparison the person has open.' },
        source: { type: 'string', description: 'Your name, shown to the person as who judged it.' },
        changes: {
          type: 'array',
          minItems: 1,
          maxItems: 500,
          items: {
            type: 'object',
            properties: {
              kind: { type: 'string', enum: ['added', 'removed', 'modified', 'moved', 'unchanged'] },
              old: { type: 'string', description: 'Paragraph id of the old revision, such as o12.' },
              new: { type: 'string', description: 'Paragraph id of the new revision, such as n7.' },
              note: { type: 'string', description: 'Why, in a sentence.' },
            },
            required: ['kind'],
          },
        },
        summary: { type: 'string', description: 'Two or three sentences on what changed overall, shown above the list.' },
        error: { type: 'string', description: 'Why you could not judge the comparison. Used instead of changes.' },
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'get_stitching',
    title: 'Read stitching',
    description:
      'How the sheets are stitched: the hand-placed match lines (id, page, p0, p1, keepSide, name, matchTo, problem), the regions that leave out title blocks, and the stitched groups (pages in order; each sheet\'s scale and rotation relative to the first; the joins, byHand or found from MATCH LINE labels). Narrow it to one page with page.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page. Every page when omitted.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'set_stitch_region',
    title: 'Set a stitch region',
    description:
      'The part of a sheet that holds the drawing, so the title block, legend, and notes are left out of the stitch: two opposite corners of a rectangle, or the corners of a polygon (an L-shaped drawing area). points null clears it, and the sheet falls back to its largest viewport, or the whole sheet. Match line ends have to lie inside the region. One undo step.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page.' },
        points: { ...pointsProp, type: ['array', 'null'] },
      },
      required: ['page'],
    },
    annotations: CHANGES,
  },
  {
    name: 'place_match_lines',
    title: 'Place match lines',
    description:
      'Joins sheets: each pair is the same match line on two pages, given by its two ends on each (a and b). The ends of the two lines meet, so a sheet at another scale is scaled to fit, and the order the ends are given in does not matter. keepSide says which side of each line stays ("right" or "left", walking from p0 to p1 as the page shows); keepPoint (any point inside the part that stays) does the same. Without either, the side with more of the sheet stays. Both lines take the pair\'s name (a station, such as STA 12+00; ML-n when omitted). Use get_page_image and get_page_text to find the match lines first. All pairs are one undo step, and one bad pair places nothing. The reply has the groups the sheets now form: check them with get_stitch_image.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        pairs: {
          type: 'array',
          minItems: 1,
          maxItems: 100,
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              a: matchLineProp,
              b: matchLineProp,
            },
            required: ['a', 'b'],
          },
        },
      },
      required: ['pairs'],
    },
    annotations: ADDS,
  },
  {
    name: 'update_match_lines',
    title: 'Adjust match lines',
    description:
      'Changes match lines by id: move their ends (p0, p1), change the side that stays (keepSide, keepPoint, or flip true), rename, pair with another line (matchTo an id, or null to unpair), or slide the sheet joined at the line by points (slide.along runs along the match line, slide.apart moves the sheets apart, negative closes the gap). resetSlide true puts the sheet back where the line ends place it. All changes are one undo step.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        lines: {
          type: 'array',
          minItems: 1,
          maxItems: 100,
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Match line id from get_stitching.' },
              p0: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
              p1: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
              keepSide: { type: 'string', enum: ['right', 'left'] },
              keepPoint: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
              flip: { type: 'boolean' },
              name: { type: 'string' },
              matchTo: { type: ['string', 'null'] },
              slide: { type: 'object', properties: { along: { type: 'number' }, apart: { type: 'number' } } },
              resetSlide: { type: 'boolean' },
            },
            required: ['id'],
          },
        },
      },
      required: ['lines'],
    },
    annotations: CHANGES,
  },
  {
    name: 'delete_match_lines',
    title: 'Remove match lines',
    description:
      'Removes match lines by id, clears the regions on regionPages, or with all true removes every match line and region. A line whose partner is removed waits for a new one. Pages made with stitch_to_page stay. redcolumn asks for Allow first; one undo step after Allow.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        ids: { type: 'array', items: { type: 'string' } },
        regionPages: { type: 'array', items: { type: 'integer' } },
        all: { type: 'boolean' },
      },
    },
    annotations: REMOVES,
  },
  {
    name: 'get_stitch_image',
    title: 'See stitched sheets',
    description:
      'A picture of the sheets stitched with page, as the stitched view joins them: each sheet cut at its region and match lines, outlined in blue and labelled with its page number. bounds is the whole group in its world space (the first sheet\'s page points); region (in the same space) picks part of it to look at a join closely. A pixel at (px, py) is the world point (region.x + px / scale, region.y + py / scale), and each sheet\'s toImage is the affine [a, b, c, d, e, f] from its page points to pixels (x\' = a·x + c·y + e, y\' = b·x + d·y + f).',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: 'Any 1-based page in the group. The page on screen when omitted.' },
        region: regionProp,
        dpi: { type: 'number', minimum: 18, maximum: 600, description: 'Pixels per inch of the first sheet. 100 when omitted.' },
        maxSize: { type: 'integer', minimum: 256, maximum: 4096, description: 'Most pixels on the longer side. 2000 when omitted.' },
        markups: { type: 'boolean', description: 'false leaves the markups off.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'show_stitched_view',
    title: 'Stitched view',
    description: 'Shows the sheets stitched with page as one drawing in redcolumn (View › Stitched View), or the normal view with on false. The drawing has to be on screen.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: 'Any 1-based page in the group. The page on screen when omitted.' },
        on: { type: 'boolean', description: 'false goes back to the normal view. true when omitted.' },
      },
    },
    annotations: CHANGES,
  },
  {
    name: 'stitch_to_page',
    title: 'Stitch to new page',
    description:
      'Stitch to New Page: the sheets stitched with page drawn onto one new page at the end of the document, kept vector, with the first sheet\'s scale so measurements work across the joins. Markups on the kept parts are copied across unless markups is false. The drawing has to be on screen. redcolumn asks for Allow first; one undo step. The reply has the new page number: look at it with get_page_image.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: 'Any 1-based page in the group. The page on screen when omitted.' },
        markups: { type: 'boolean', description: 'false leaves the markups off the new page. true when omitted.' },
      },
    },
    annotations: ADDS,
  },
  {
    name: 'list_viewports',
    title: 'List viewports',
    description: "A page's viewports: regions drawn at their own scale (a detail or section beside a plan). Each has an id, name, page, rect (x, y, w, h in PDF points, origin at the top-left, y down), and scale, written the way get_page writes it. Narrow it to one page with page.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page. Every page when omitted.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'place_viewport',
    title: 'Place a viewport',
    description:
      'Places a viewport on a page: a rectangle drawn at its own scale, so measurements inside it use that scale. points are two opposite corners in PDF points, origin at the top-left, y down; the part off the page is cut away, as drawing one on screen does. It starts at the page\'s scale unless scale is given. Set it later with set_viewport_scale. One undo step. The reply has the new viewport\'s id and rect. For several, use place_viewports.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number.' },
        points: { ...pointsProp, minItems: 2, maxItems: 2, description: 'Two opposite corners of the rectangle, in PDF points, origin at the top-left, y down: [[x, y], [x, y]].' },
        name: { type: 'string', description: 'Viewport 1, Viewport 2... in page order when omitted.' },
        scale: scaleProp,
      },
      required: ['page', 'points'],
    },
    annotations: ADDS,
  },
  {
    name: 'place_viewports',
    title: 'Place many viewports',
    description:
      'Places many viewports as one undo step (up to 100). Each item is what place_viewport takes: points (two opposite corners in PDF points, origin at the top-left, y down), and optionally page, name, and scale; a page on the call is the default for items without one. Everything is checked first, so one bad item places nothing. The reply lists each new viewport\'s id, name, page, rect, and scale.',
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number for items that do not give their own.' },
        viewports: {
          type: 'array',
          minItems: 1,
          maxItems: 100,
          items: {
            type: 'object',
            properties: {
              page: { type: 'integer', description: '1-based page number.' },
              points: { ...pointsProp, minItems: 2, maxItems: 2, description: 'Two opposite corners of the rectangle, in PDF points, origin at the top-left, y down: [[x, y], [x, y]].' },
              name: { type: 'string' },
              scale: scaleProp,
            },
            required: ['points'],
          },
        },
      },
      required: ['viewports'],
    },
    annotations: ADDS,
  },
  {
    name: 'set_viewport_scale',
    title: 'Set a viewport scale',
    description: "Sets the scale of one viewport, so measurements inside it use that scale and the page's is left alone. Text that is not a scale is refused and nothing changes. One undo step. The reply has the viewport as it now stands.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        id: { type: 'string', description: 'Viewport id from list_viewports or place_viewport.' },
        scale: scaleProp,
      },
      required: ['id', 'scale'],
    },
    annotations: CHANGES,
  },
  {
    name: 'get_sheet_info',
    title: 'Get sheet info and scale status',
    description:
      "What each sheet's title block says and where its scale stands: number, title, discipline, the scale text, the sheet size it was stated for and the size the page actually is, the scale that gives (corrected when an A1 drawing was saved at A3), the review status (ready, resized, check, matches, differs, views, nts, none), the page's scale and viewports, and whether it can be measured. For one page it also checks the scale against the page's scale bars and dimensions, and proposes viewports from its view titles. Without a page, every page is listed. Check this before measuring.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number. Every page when omitted (without the scale-bar check).' },
      },
    },
    annotations: READS,
  },
  {
    name: 'propose_page_scale',
    title: 'Propose a page scale',
    description:
      "Proposes the scale a page's title block states, when redcolumn did not read it (or read it wrong). It sets no scale: it waits in Document › Review Sheet Scales, marked as from AI with your reasoning, where a person checks it against the page and applies it. Sheet info from the drawing standard or typed by a person is never replaced. Give the scale as written (1:100, 1/4\" = 1'-0\", NTS, AS SHOWN) and the sheet size it is for when the title block says (A1).",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        page: { type: 'integer', description: '1-based page number.' },
        scale: { type: 'string', description: 'The scale as the title block writes it, such as 1:100 or 1/4" = 1\'-0".' },
        sheetSize: { type: 'string', description: 'The sheet size the scale is stated for (A1, ANSI D), when the title block says.' },
        reasoning: { type: 'string', description: 'Where on the sheet you read it, for the person checking it.' },
      },
      required: ['page', 'scale', 'reasoning'],
    },
    annotations: ADDS,
  },
  {
    name: 'get_drawing_standard',
    title: 'Get the drawing standard',
    description:
      "A drawing standard is how one client draws: where its title block keeps each field, the words its section, elevation and match-line titles use, how it writes chainages and names alignments, and the marks (callout bubbles) a person taught. This returns the standard, with the boxes the person drew in the setup as examples (each with the text read in it), and how each setup step did (band good, review or ai; references found and linked; links checked and wrong; marks found by shape but not read; chainages whose alignment is unclear). Steps in band ai or review are the ones to finish: read the sheets with get_page_text and get_page_image, compare with the examples, and propose links with propose_sheet_links and words with propose_standard_hints. Without standard, the one the document is read with.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        standard: { type: 'string', description: 'Standard name or id. The one the document is read with when omitted.' },
      },
    },
    annotations: READS,
  },
  {
    name: 'list_sheet_links',
    title: 'List sheet links',
    description:
      "The cross-references found in a set (callouts to details, sections and elevations, sheet references, match lines, chainages): where each is (page, rect), what it reads, the page it goes to, its confidence and why, its status (auto: kept; review: waiting for a person; accepted; rejected) and source (text, taught mark, pdf, ai). Filter by kind, status or page. Use it to see which references were missed or doubtful before proposing links.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        kind: { type: 'string', enum: ['detail', 'section', 'elevation', 'enlargedPlan', 'sheet', 'matchline', 'chainage'] },
        status: { type: 'string', enum: ['auto', 'review', 'accepted', 'rejected'] },
        page: { type: 'integer', description: '1-based page the links are on.' },
        offset: { type: 'integer', description: 'Skip this many (for the next part of a long list).' },
        limit: { type: 'integer', description: 'At most this many, 1 to 500 (100 when omitted).' },
      },
    },
    annotations: READS,
  },
  {
    name: 'propose_sheet_links',
    title: 'Propose sheet links',
    description:
      "Proposes cross-references the text and shape search missed or got wrong: a callout's place on its page and the page it goes to. Nothing is linked yet: each waits in the Links panel under Needs review, marked as from AI with your reasoning, until a person keeps it. Proposing the same place again updates it. Rects are in PDF points from the top-left of the page. All links are checked first; one bad item means none is added.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        links: {
          type: 'array',
          maxItems: 200,
          items: {
            type: 'object',
            properties: {
              page: { type: 'integer', description: '1-based page the callout is on.' },
              rect: { type: 'object', description: 'The callout, in PDF points: {x, y, w, h}.', properties: { x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' } }, required: ['x', 'y', 'w', 'h'] },
              label: { type: 'string', description: 'The callout as it reads: 3/C-501, A/000031, CH 100, MATCH LINE 000032.' },
              kind: { type: 'string', enum: ['detail', 'section', 'elevation', 'enlargedPlan', 'sheet', 'matchline', 'chainage'] },
              detail: { type: 'string', description: 'The detail, section or elevation identifier (3, A, E1), when the label has one.' },
              targetPage: { type: 'integer', description: '1-based page it goes to.' },
              targetRect: { type: 'object', description: 'Where on the target page it goes (the detail or section), in PDF points. The whole sheet when omitted.', properties: { x: { type: 'number' }, y: { type: 'number' }, w: { type: 'number' }, h: { type: 'number' } }, required: ['x', 'y', 'w', 'h'] },
              reasoning: { type: 'string', description: 'How you read it, for the person checking it.' },
            },
            required: ['page', 'rect', 'label', 'kind', 'targetPage', 'reasoning'],
          },
        },
      },
      required: ['links'],
    },
    annotations: ADDS,
  },
  {
    name: 'propose_standard_hints',
    title: 'Propose drawing standard words',
    description:
      "Proposes words the drawing standard should learn, so text search finds more on this set and the next: section, elevation and match-line title words, chainage prefixes, alignment names and their form (MC## with # a digit, @ a letter), and the title words of plan and cross-section sheets. Words are only ever added. Nothing changes yet: the proposal waits in Set Up Drawing Standard's Review step, with your reasoning, until a person accepts it.",
    inputSchema: {
      type: 'object',
      properties: {
        document: documentProp,
        standard: { type: 'string', description: 'Standard name or id. The one the document is read with when omitted.' },
        sectionWords: { type: 'array', items: { type: 'string' } },
        elevationWords: { type: 'array', items: { type: 'string' } },
        matchLineWords: { type: 'array', items: { type: 'string' } },
        chainagePrefixes: { type: 'array', items: { type: 'string' } },
        alignments: { type: 'array', items: { type: 'string' } },
        alignmentTemplate: { type: 'string' },
        planTitles: { type: 'array', items: { type: 'string' } },
        sectionTitles: { type: 'array', items: { type: 'string' } },
        reasoning: { type: 'string', description: 'What you saw on the sheets, for the person deciding.' },
      },
      required: ['reasoning'],
    },
    annotations: ADDS,
  },
];

export type ToolCall = (name: string, args: Record<string, unknown>) => Promise<unknown>;

/** Lines already split. A chunk that ends mid-line stays in `buf`. */
export function takeLines(state: { buf: string }, chunk: string, max = 8_000_000): string[] {
  state.buf += chunk;
  if (state.buf.length > max) {
    state.buf = '';
    throw new Error('Message too large');
  }
  const lines: string[] = [];
  let nl = state.buf.indexOf('\n');
  while (nl >= 0) {
    const line = state.buf.slice(0, nl).replace(/\r$/, '');
    state.buf = state.buf.slice(nl + 1);
    if (line) lines.push(line);
    nl = state.buf.indexOf('\n');
  }
  return lines;
}

function rpcResult(id: unknown, result: unknown): string {
  return JSON.stringify({ jsonrpc: '2.0', id, result });
}

function rpcError(id: unknown, code: number, message: string): string {
  return JSON.stringify({ jsonrpc: '2.0', id, error: { code, message } });
}

function toolText(text: string, isError: boolean) {
  return { content: [{ type: 'text', text }], isError };
}

/** A picture a tool returns beside its text: base64 bytes and their media type. */
interface ToolImage {
  mimeType: string;
  data: string;
}

function imageOf(value: unknown): ToolImage | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const image = (value as { image?: unknown }).image;
  if (!image || typeof image !== 'object' || Array.isArray(image)) return null;
  const { mimeType, data } = image as { mimeType?: unknown; data?: unknown };
  if (typeof mimeType !== 'string' || !mimeType.startsWith('image/') || typeof data !== 'string' || !data) return null;
  return { mimeType, data };
}

/**
 * What a tool answered, as MCP content. A reply with an `image` (get_page_image) becomes a text
 * block for the rest of it and an image block the client shows the model.
 */
export function toolContent(value: unknown) {
  if (typeof value === 'string') return toolText(value, false);
  const image = imageOf(value);
  if (!image) return toolText(JSON.stringify(value), false);
  const { image: _, ...rest } = value as { image: unknown } & Record<string, unknown>;
  return { content: [{ type: 'text', text: JSON.stringify(rest) }, { type: 'image', mimeType: image.mimeType, data: image.data }], isError: false };
}

function argsOf(params: unknown): Record<string, unknown> {
  if (!params || typeof params !== 'object' || Array.isArray(params)) return {};
  const record = params as { arguments?: unknown };
  const args = record.arguments;
  if (!args || typeof args !== 'object' || Array.isArray(args)) return {};
  return args as Record<string, unknown>;
}

/**
 * One MCP message, newline-delimited JSON, the framing Copilot, Claude, Codex, and Cursor use.
 * Notifications get no reply. A missing app becomes the tool's text, not a protocol failure.
 */
export async function handleMcpMessage(line: string, call: ToolCall, version = '0.1.0'): Promise<string | null> {
  let msg: unknown;
  try {
    msg = JSON.parse(line);
  } catch {
    return rpcError(null, -32700, 'Parse error');
  }
  if (!msg || typeof msg !== 'object' || Array.isArray(msg)) return rpcError(null, -32600, 'Invalid request');
  const rec = msg as { id?: unknown; method?: unknown; params?: unknown };
  const hasId = 'id' in rec && rec.id !== undefined && rec.id !== null;
  const id = hasId ? rec.id : null;
  if (typeof rec.method !== 'string') return hasId ? rpcError(id, -32600, 'Invalid request') : null;
  if (!hasId) return null;

  try {
    switch (rec.method) {
      case 'initialize': {
        const params = rec.params as { protocolVersion?: unknown } | undefined;
        const protocolVersion = typeof params?.protocolVersion === 'string' && params.protocolVersion ? params.protocolVersion : '2025-03-26';
        return rpcResult(id, {
          protocolVersion,
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: 'redcolumn', version },
        });
      }
      case 'ping':
        return rpcResult(id, {});
      case 'tools/list':
        return rpcResult(id, { tools: TOOLS });
      case 'tools/call': {
        const params = rec.params as { name?: unknown } | undefined;
        const name = params?.name;
        if (typeof name !== 'string' || !isAssistantMethod(name)) return rpcResult(id, toolText('Unknown tool.', true));
        try {
          const value = await call(name, argsOf(rec.params));
          return rpcResult(id, toolContent(value));
        } catch (err) {
          const message = err instanceof Error && err.message ? err.message : OFFLINE;
          return rpcResult(id, toolText(message, message !== OFFLINE));
        }
      }
      default:
        return rpcError(id, -32601, 'Method not found');
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Internal error';
    return rpcError(id, -32603, message);
  }
}
