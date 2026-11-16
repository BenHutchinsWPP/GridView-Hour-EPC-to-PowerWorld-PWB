# Idea: area scaling through CaseBridge, without PowerWorld in the middle

Status: idea only (2026-11-15). Nothing here is built. CaseBridge has not yet converted an EPC into another
program's case, so this is the step after that works.

## Today
Both cases are opened in PowerWorld through SimAuto: the GridView EPC to read its loads and gens, and the PWB to
scale. Then the PWB is moved to the targets by solving Newton-Raphson again and again in small steps, holding back
the gens/loads which make it diverge. That costs one to two hours per hour-EPC, and needs Windows, PowerWorld and a
SimAuto license. Some of what goes wrong is PowerWorld-in-the-middle trouble, not trouble with the targets.

## The idea
1. CaseBridge reads the GridView EPC directly (its EPC reader exists).
2. The seed PWB is saved once as an EPC (or RAW) from PowerWorld, and CaseBridge reads that as the second case.
3. The targets are computed the same way as `Scripts/area.py` does now: load per area, generation per area and
   GenType. CaseBridge already holds the unit and fuel code lists with each code's EPC and PowerWorld integer and a
   generic group (`core/schemas/codes/generator_codes.json`), so classification is by the EPC integers with no
   PowerWorld lookup.
4. The scaled case is written out by CaseBridge, and solved in a program which can solve it straight to the target,
   e.g. CalGrid with HELM, instead of being stepped there with Newton-Raphson.

## What it would remove
- The stepping, holding and retrying in `Scripts/scale.py`, if HELM reaches the solution (or proves there is none)
  in one go.
- The Windows / SimAuto dependence of the targets step.

## What has to exist first
- A CaseBridge EPC -> (CalGrid or other HELM-capable format) conversion that is checked against a reference.
- A decision on where the giant swing / losses go when no stepping is done.
- A way back to PWB, if PWB stays the deliverable.

The target maths in `Scripts/area.py` is plain pandas over load and gen tables, so it should move across with only
the field names changed.
