# Connect an AI client

In the desktop app, an AI client on your computer (Copilot, Claude, Codex or Cursor) can look at your open drawings and help with them. It does this through redcolumn's local **MCP server**. The client gets a set of **tools**. This section explains each one: what it does, what it returns, where it stops, and which AI model suits it.

> **Note:** This is for the desktop app only. The web version has no MCP server.

## Turn it on

1. Choose {{File > Preferences…}} and open **MCP**.
2. Tick **Enable MCP server**. It takes effect straight away, not when you press **Save**.
3. Find your client in the list, click **Copy**, and add that text to the client's settings (some have an install button).
4. Leave redcolumn open while you use the client.

## What to know first

- **Everything stays on your computer.** The server listens only on this computer, and only a client started by this installation can talk to it. The drawings do not go to redcolumn. They go to the AI client you connected, as the client reads them.
- **The app must be open.** With redcolumn closed or the switch off, a tool answers "Open redcolumn and turn on the local assistant."
- **Allow prompts.** Anything that deletes, saves, or changes a document waits for you to click **Allow** in redcolumn. The client waits about two minutes for your answer, then gives up.
- **Undo works.** Every change the client makes is one step in {{Edit > Undo}}, however many markups it touched.
- **Some things cannot be changed.** Read-only drawings, locked markups, and other people's markups. The tool says so.
- **Suggestions are not applied.** Takeoff proposals, assembly proposals, page scales, sheet links and drawing-standard words wait for you to review. The AI can only propose; you approve or reject each one.
- **Pages are numbered from 1.** Positions are in PDF points (72 to the inch), measured from the **top-left** of the page, with y going **down**.
- Most tools take a **document** (name or id). Leave it out and the drawing on screen is used.
- **Mistakes are safe.** When a client sends many items at once, redcolumn checks all of them first. One bad item means nothing is changed.

## All the tools

| Tool | What it does | Page |
|---|---|---|
| `list_open_documents` | Lists the open drawings. | [Open drawings and pages](#mcp-documents) |
| `get_page` | A page's size and scale. | [Open drawings and pages](#mcp-documents) |
| `list_markups` | Lists markups, with filters and counts. | [List markups](#mcp-list-markups) |
| `get_page_image` | A picture of a page or region. | [See and read a page](#mcp-page-view) |
| `get_page_text` | The page's text, word by word. | [See and read a page](#mcp-page-view) |
| `place_markup`, `place_markups` | Draws markups. | [Draw markups](#mcp-draw) |
| `update_markup`, `update_markups` | Changes points, text and style. | [Edit and delete markups](#mcp-edit) |
| `set_markup`, `set_markups` | Sets status, comment, subject, columns. | [Edit and delete markups](#mcp-edit) |
| `select_markups` | Selects markups. | [Edit and delete markups](#mcp-edit) |
| `delete_markups` | Deletes markups (asks first). | [Edit and delete markups](#mcp-edit) |
| `list_commands`, `run_command` | Finds and runs menu commands. | [Run commands](#mcp-commands), [Every command](#mcp-command-list) |
| `list_assemblies`, `propose_takeoff`, `list_takeoff_proposals` | Suggests takeoff assemblies. | [Takeoff proposals](#mcp-takeoff) |
| `propose_assembly`, `list_assembly_proposals` | Proposes new or edited assemblies (you decide). | [Takeoff proposals](#mcp-takeoff) |
| `get_sheet_info`, `propose_page_scale` | Reads and suggests sheet scales. | [Sheet scales](#mcp-scale) |
| `get_drawing_standard`, `list_sheet_links`, `propose_sheet_links`, `propose_standard_hints` | Finishes a drawing standard's links. | [Drawing standards](#drawing-standards) |
| `start_comparison`, `get_comparison`, `report_document_changes` | AI comparison of two revisions. | [Compare documents](#mcp-compare) |
| `get_stitching`, `get_stitch_image`, `place_match_lines`, `update_match_lines`, `delete_match_lines`, `set_stitch_region`, `show_stitched_view`, `stitch_to_page` | Joins sheets. | [Stitch sheets](#mcp-stitch) |
| `list_viewports`, `place_viewport`, `place_viewports`, `set_viewport_scale` | Regions with their own scale. | [Viewports](#mcp-viewports) |

## What the assistant cannot do

- Open or close files, or save: it can ask with `run_command`, and you click **Allow**.
- Place images, stamps, signatures or attachments by itself: it needs a file you choose.
- Read a scanned page's text until the page is OCRed.
- Work when redcolumn is closed or **Enable MCP server** is off.
- Change a read-only drawing, a locked markup, or someone else's markup.
- Apply its own suggestions.

If a tool fails, the message says why (for example "No markup 12 is on Plan.pdf"). See [Something is not working](#troubleshooting). To choose an AI model, see [Which AI model for which job](#mcp-models).
