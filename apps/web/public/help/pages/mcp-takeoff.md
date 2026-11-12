# Takeoff proposals

The client can suggest which [assembly](#assemblies) fits each measurement, and propose new assemblies, or edits to the ones the drawing has. It can only propose: it never adds anything to the takeoff or the document's assemblies itself. You decide.

## list_assemblies

**Does:** lists the drawing's assemblies.

**Returns:** each assembly's items with their units, formulas and unit costs, and its variables. `formulaNames` are the measured values a formula can use, such as `[Length]`, `[Area]` and `[Count]`.

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

**Takes:** a list of assemblies (up to 10 a call). Each has a `name`, its `items` (name, unit, formula, and optionally waste, round up, round to, unit cost), its `variables`, and a `reasoning` saying why. `replaces` names the assembly an edit belongs to; where item names match, the edit keeps those items' identity, so measurements keep any quantities you set by hand.

**Where you see it:** the same Review AI Proposals panel as takeoff proposals. An edit shows what it adds, removes and changes.

**Limits:**

- Everything is checked first: one bad item means nothing is kept.
- A name the drawing already uses is refused unless `replaces` names the assembly that has it; a `replaces` that matches nothing is refused too.
- Item names must be unique within the assembly; waste is 0 to 100; unit cost is 0 or more.
- Formulas must work out to a number, using the measured values (`formulaNames` from `list_assemblies`) and the assembly's own variables. Unknown names, broken syntax and a division by zero in the trial run (every value 1) are refused. A formula that divides by a measured value which is 0 on some marks is not caught: it shows its error on those marks after you approve.
- 40 items and 20 variables at most per assembly.

**Model:** a strong model. The formulas have to be right, and the reasoning has to convince you.

> **Note:** After you approve, {{Edit > Undo}} takes the document change back, but the proposal does not return to the list (unlike takeoff proposals): ask the client to propose it again.

## list_assembly_proposals

**Does:** lists assembly proposals and where each stands: pending (waiting for you) or rejected. Approved ones have left the list: they are the drawing's assemblies now, so `list_assemblies` has them.

**Limits:** it changes nothing.

**Model:** a cheaper model is fine.

Back to [Connect an AI client](#mcp).
