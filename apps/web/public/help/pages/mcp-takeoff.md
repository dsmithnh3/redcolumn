# Takeoff proposals

The client can suggest which [assembly](#assemblies) fits each measurement, and propose new assemblies, or edits to the ones the drawing has. It can only propose: it never adds anything to the takeoff or the document's assemblies itself. You decide.

## list_assemblies

**Does:** lists the drawing's assemblies.

**Returns:** each assembly's items with their units, formulas and unit costs, and its variables. `formulaNames` are the measured values a formula can use, such as `[Length]`, `[Area]` and `[Count]`; `columnNames` are the [list columns](#list-columns) it can also use, the drawing's own and the built-in ones.

**Limits:** it changes nothing. The client should run it before `propose_takeoff` and `propose_assembly`.

**Model:** a cheaper model is fine.

## propose_takeoff

**Does:** proposes an assembly for each of some measurements, with its reasoning and, if it likes, the quantities it expects.

**Takes:** a list. Each item names a measurement (by id or number), an assembly (name or id), a reason, and optionally the expected quantities.

**Where you see it:** the proposals wait in {{Tools > Takeoff > Review AI Proposals…}}, a panel that opens by itself when they arrive. You **approve**, **edit** or **reject** each one. Nothing is added to the takeoff until you approve it.

**Limits:**

- 500 proposals a call.
- A new proposal for a measurement replaces its earlier pending one.
- Each item must say *why*, so you can check it.
- The markup must be a measurement.

**Model:** a strong model. It has to match the right assembly to each measurement, from the drawing and the assembly list, and the quantities it expects are a check on its own work.

## list_takeoff_proposals

**Does:** lists the proposals and where each stands: pending, approved (by whom, and whether they edited it) or rejected.

**Limits:** it changes nothing.

**Model:** a cheaper model is fine.

## propose_assembly

**Does:** proposes takeoff assemblies for the drawing: new ones, or edits of existing ones with `replaces` (their name or id). The client can only propose: each assembly waits in {{Tools > Takeoff > Review AI Proposals…}}, with its items, variables and the client's reasoning, until you **approve** or **reject** it. Nothing is written into the document until you approve. When you approve you can also tick **Also save to My Assemblies**, which keeps a copy in this browser for any document (your My Assemblies are never read or changed by the client itself).

**Takes:** a list of assemblies (up to 10 a call). Each has a `name`, its `items` (name, unit, formula, and optionally waste, round up, round to, unit cost), its `variables`, a `description` of what it measures, and a `reasoning` saying why. `replaces` names the assembly an edit belongs to; where item names match, the edit keeps those items' identity, so measurements keep any quantities you set by hand. An edit that sends no `description` keeps the assembly's notes; an empty one clears them.

**Where you see it:** the same Review AI Proposals panel as takeoff proposals. An edit shows what it changes: the name, the notes, the variables (added, removed, and old → new values) and the items (added, removed, changed).

**Limits:**

- Everything is checked first: one bad item means nothing is kept.
- A name the drawing already uses is refused unless `replaces` names the assembly that has it; a `replaces` that matches nothing is refused too, and two edits of the same assembly in one call are refused.
- Names are at most 200 characters, units 40, notes 2000, formulas 1000. Names, units and variable names cannot contain characters that do not display (control, bidi override or zero-width).
- Item names must be unique within the assembly; waste is 0 to 100 (the Assemblies window allows a little more, so an edit that copies a larger waste is refused); round to is 0.000001 to 1000000; unit cost is 0 to 1000000000.
- Formulas are tried once before anything is kept, with the assembly's variables at their values and every measured value and list column at 1: unknown names, broken syntax, a result that is text, and a division by zero in that trial run are refused. A formula that passes can still fail on real measurements — an `IF` that returns text for some lengths, or a division by a measured value that is 0 on some marks: those show as errors in the takeoff after you approve.
- 40 items and 20 variables at most per assembly, and 200 proposals kept for one drawing (the oldest rejected ones go first).

**Model:** a strong model. The formulas have to be right, and the reasoning has to convince you.

> **Note:** After you approve, {{Edit > Undo}} takes the document change back, but the proposal does not return to the list (unlike takeoff proposals): ask the client to propose it again. **Also save to My Assemblies** merges by name: an assembly already saved there with the same name is replaced, and the message says which.

## list_assembly_proposals

**Does:** lists assembly proposals and where each stands: pending (waiting for you) or rejected. Approved ones have left the list: they are the drawing's assemblies now, so `list_assemblies` has them.

**Limits:** it changes nothing.

**Model:** a cheaper model is fine.

Back to [Connect an AI client](#mcp).
