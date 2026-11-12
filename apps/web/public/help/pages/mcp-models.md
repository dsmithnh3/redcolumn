# Which AI model for which job

A client that can pick its model (or hand work to helpers) should match the model to the job. Reading drawings and working out coordinates is where cheap models go wrong. Filing and tidying is where they do fine. Each tool page ends with a **Model** line that says the same thing for that tool.

## Use a strong model

These need good vision, careful geometry, or judgement, and a wrong answer costs you a bad takeoff or a bad scale.

- **Reading a sheet from `get_page_image`**, such as finding match lines, working out what a symbol means, or locating a title block. Small linework is easy to misread.
- **Stitching**: `place_match_lines`, `update_match_lines`, `set_stitch_region`. It is geometry from a picture: the ends of each line must land in the right place on both sheets, and the side that stays must be right.
- **`propose_takeoff`.** It has to match measurements to the right assembly and give a reason you can check.
- **`propose_assembly`.** The formulas and waste factors it writes have to be right, and the reasoning has to convince you.
- **`propose_page_scale` and sheet scale reviews** with `get_sheet_info`. Scale is the number every measurement depends on.
- **Placing viewports** from view titles on a busy sheet, and **placing markups at exact positions** worked out from a picture.
- **Comparing documents** (`get_comparison`, `report_document_changes`). Matching paragraphs by meaning, not position, is subtle, and a long comparison needs a model that stays consistent across many pages of results.
- **Open-ended requests** that mix several tools, such as "check every sheet's scale and mark the ones that look wrong".

## A cheaper model is enough

These are lookups and bulk edits with no judgement about the picture.

- **Listing and counting**: `list_open_documents`, `list_markups` (especially `summary`), `list_viewports`, `list_assemblies`, `list_takeoff_proposals`, `list_assembly_proposals`, `get_page`, `get_stitching`.
- **Searching text**: `get_page_text` with `search`, when the page has real text.
- **Setting status, comment, subject or a column** on markups you can already name (`set_markups`): "mark everything in the list as Accepted".
- **Restyling or moving markups** with exact values you give it (`update_markups`), such as "make all clouds red".
- **Selecting markups** and **running a command** by id (`select_markups`, `run_command` after `list_commands`).
- **Deleting by id** after you have picked what to delete.
- **Drawing from numbers you supply**, such as "put a text box at these points".

## Tips for any model

- **Count before listing.** Ask for `list_markups` with `summary` first, then filter. A big drawing can hold thousands of markups.
- **Use the "many" tools.** One `set_markups` call is one undo step. Fifty `set_markup` calls are fifty.
- **Look wide, then close.** Ask for the whole page first, then a region a few hundred points across.
- **Check scale before measuring.** `get_sheet_info` says whether a page can be measured.
- **Check its work.** After `place_match_lines`, `get_stitch_image` shows whether the join is right. After placing markups, `get_page_image` shows them in place.

Back to [Connect an AI client](#mcp).
