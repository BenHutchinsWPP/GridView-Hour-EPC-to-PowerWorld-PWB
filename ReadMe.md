# Overview

## Purpose
Given (Input):
- A solved seed-case PowerWorld Binary (PWB) file that was used to generate the GridView models, and
- EPC(s) which have been exported from GridView for one or more hour snapshots.

Produce (Output): a solved PWB for every provided GridView hour. Two modes, set by `mode` in `settings.toml`:
- **`area`** (default): the same load MW & Mvar in each area, and the same generation MW in each area for each generation type (Solar, Hydro, Wind, Thermal, Battery Storage, Unknown). Topology is not matched: nothing is created or deleted, and no branch or load is opened or closed to match GridView. (The last-ditch solve aids, STATCOMs and dropping collapsed sections, can still change the network to keep the case solving: `add_statcoms`, `drop_collapsed`.) Gens are brought online or taken offline only to suit their area & type target ([Unit Commitment](#unit-commitment)). Each gen and load whose key is in both cases starts from its own EPC value; whatever doesn't match key by key is made up by scaling its whole area (or area & type) group ([Key Basis](#key-basis)).
- **`topology`**: the same topology (buses, branches, loads, gens, and their statuses) and the same MW/Mvar on every load and gen (unless `balance_swing` moves them: see [Topology Mode](#topology-mode)). Missing elements are created in the PWB.

## Status
Area mode: first draft (2026-11-15), first written in `~/projects/areascale` (the work is on branch `area-mode` here; `areascale-draft` is the old draft). Neither mode's current code has been run against PowerWorld yet. The offline tests in `tests/test_lib.py` pass.

## Prerequisites
- PowerWorld v23+ with a SimAuto license
- Python 3.11+ (`settings.toml` is read with the built-in `tomllib`)
- Python libraries: `pip install -r requirements.txt`

## How to Use
- Install Visual Studio Code and Python, and open `Workspace.code-workspace`.
- Place GridView EPC Exports into `./HourEPCs/`, and remove any sample cases.
- Place your PWB seed-case into `./Seed/`, and remove any sample cases.
- Set `settings.toml` (see [Settings](#settings)).
- Run `01 Prepare Seed.py`. It uses the first `.pwb` in `./Seed/` and writes `./Prepared/Prepared.pwb`, with each intermediate step saved beside it. Topology mode also uses the EPCs: the first one (alphabetically) for the topology, all of them for distributed generation loads.
- Run `02 Scale Hours.py`. For each EPC in `./HourEPCs/` it writes, in `./Output/`:
  - `<hour>_Log.xlsx`: one workbook per hour, written once before scaling (the plan, with the `Final` columns blank, so a failed run still leaves it) and again at the end. If the second write fails (e.g. the file is open in Excel) the case is still saved. Sheets:
    - `areas` (area mode only): a row per area & group (each GenType, or `Load`): `Seed MW`, `Target MW`, `Unmet MW` (what the group's own units could not take; with `makeup`, the area's other GenTypes cover what they can, shown as their `Make-up MW`), `Not Reached MW` (held, excluded or diverged units), `Make-up MW` (what the make-up moved onto the group to cover a shortfall in it or, for gens, in another GenType of its area), `Balance MW` (what swing balancing moved onto the group), `Final MW` (read back from PowerWorld), `Error MW` (Final - Target), Mvar and online-unit counts, the area's losses on its `Load` row (`Seed Loss MW`, `Est. Loss MW` that the pre-balance assumed, `Final Loss MW` read back), and a `Note`. Per row, Final = Target - Unmet - Not Reached + Make-up + Balance; `Note` says `unexplained` when it isn't.
    - `exceptions` (both modes): a row per gen, load, STATCOM or opened branch which is not on plan, with a `Reason` naming the run (`main`, `midrun<N>`, `makeup`, `balance<N>`).
    - `runs` (both modes): a row per scaling run (`prebalance`, `main`, `midrun<N>`, `distgen`, `makeup`, `balance<N>`): swing before and after, MW moved, steps, diverged steps and a note.
    Not logged: the per-unit plan, shunt switching in the seed preparation.
  - `<hour>.pwb`. EPCs which already have an output `.pwb` are skipped; delete it to redo that hour. If one EPC fails, the error is printed and the rest continue.
  - Note: It may take ~1-2 hours per EPC, depending on your CPU's speed.
- Run `03 Merge Reports.py` to merge every hour's log into `03 Merge Reports.xlsx`.
- Code layout: `Scripts/pw.py` talks to PowerWorld. `Scripts/prep.py` holds the seed preparation both modes share, `Scripts/topo.py` the topology mode's matching and targets, and `Scripts/area.py` the area mode's targets and the swing balancing. `Scripts/scale.py` moves the case to the targets. `Scripts/settings.py` reads `settings.toml`. `tests/test_lib.py` runs without PowerWorld: `python tests/test_lib.py`.

## Settings
`settings.toml`, at the top of the repo, with a comment on each setting. `[run]` holds the switches the two step scripts read:
- `mode`: `"area"` or `"topology"`.
- `create_giant_swing`: add a giant swing (`ID`=`xx`), or use the case's own swing (the gens on its slack bus). `swing_bus`: `"slack"`, `"max_fault_duty"` or a bus number.
- `balance_swing`: move MW off the swing onto the whole case's generation (load stays on its targets), during scaling whenever the swing leaves its band (the giant swing: `[area] balance_trigger_mw` and/or its MW limits, `balance_at_swing_limits`; the case's own slack: `balance_trigger_mw` from where the plan expects it, or its MWMin to MWMax), and after scaling until it is at its goal (the giant swing 0 MW; the case's own slack its EPC MW within its limits, or anywhere within its limits if its key isn't in the EPC). The swing unit itself is never moved, in either mode.
- `adjust_seed_shunts`, `gen_terminal_voltage_control`, `compute_pvqv`: the seed preparation steps.
- `pretest_gens`, `use_pvqv`: extra scaling exclusions, off by default. Turn them on for a case which is hard to solve. `use_pvqv` needs step 01's `compute_pvqv` (on by default).
- `[scale] add_statcoms`, `drop_collapsed`: last-ditch solve aids, on by default (see [Iterate to Gen and Load Targets](#iterate-to-gen-and-load-targets)).

The other sections (`[area]`, `[scale]`, `[prep]`, `[topology]`, `[pw]`) set the tuning constants of the `Scripts/` file they name: e.g. `[scale] exclude_vmin` sets `EXCLUDE_VMIN` in `Scripts/scale.py`. The code holds the defaults, so a setting left out of the file keeps its default, and a misspelt one is an error. `[area.unit_gentype]` / `[area.fuel_gentype]` add to or override the GenType maps.

# Process Notes

## Area Mode

### Methodology Summary
- Prepare the seed (step 01): see [Prepare Seed](#prepare-seed). The topology is left as it is.
- Read the loads and gens of the GridView EPC and of the prepared PWB (both opened in PowerWorld).
- Compute targets (`Scripts/area.py`):
  - **Load, per `AreaNumber`.** Target = the EPC's total MW and Mvar of Closed loads (constant power + current + impedance; the EPC's distributed generation is ignored). In the PWB, on every Closed load: distributed generation goes Open & 0, constant current and impedance go to 0, and constant power MW / Mvar take the load's S + I + Z total, scaled. Positive and negative load are matched separately: the area's positive loads are scaled to the EPC's positive total for that area, and its negative loads (DC converters, generation modelled as load) to the EPC's negative total. That is four factors per area: MW +, MW −, Mvar +, Mvar −. EPC-only negative loads (key not in the PWB, e.g. GridView's X1/X2 distributed generation loads) have nothing in the PWB to be placed on, so they are netted into the area's positive target instead (MW and Mvar), not left to the swing. If no seed load in the area has Mvar of a sign the EPC needs, that Mvar is spread over the area's zero-Mvar loads in proportion to |MW|. Loads that aren't scaled (Open ones, and non-conforming ones unless `scale_nonconforming`) keep their MW, but their distributed generation is also set Open & 0: GridView ignores distributed generation, so no hour keeps any. After the main scaling, every load still carrying dist gen (never touched, held or excluded) has it cleared in one solve-checked edit, and the final balance (if `balance_swing`) takes up the MW. The `runs` sheet lists it as `distgen`, with the swing before and after, and names any load which kept its dist gen because the case diverged without it.
  - **Generation, per `AreaNumber` & GenType.** Each gen's GenType comes from its `UnitTypeCode`, or from its `FuelTypeCode` when the unit type is unstated (`UN`, `OT`, `FC` or blank); see [GenType](#gentype). Target = the EPC's total `MWSetPoint` of Closed gens in that bucket. Each unit's basis MW (its EPC MW if its key is in the EPC, else its seed MW) is scaled by one factor to the bucket's target, capped at `MWMax`, with offline units brought online if the online ones can't reach it ([Key Basis](#key-basis), [Unit Commitment](#unit-commitment)). A bucket with nothing to scale shares its target in proportion to `MWMax` (`|MWMin|` for a negative target). The giant swing and DC ties (unit type `DC`, set by hand) are never scaled, and are left out of both cases' totals. So are all gens on the PWB's slack bus, even when there is a giant swing: the solution sets their MW. If more than `max_unknown_share` (0.5) of the EPC's generation has no GenType, the run stops before scaling: see [GenType](#gentype).
  - **Non-conforming loads** (PowerWorld's load `AGC` field = NO: RAW's SCALE flag 0, and presumably an EPC's `nonc` flag 1) are never scaled, and are left out of both cases' area totals, so conforming load is matched against conforming load. This is decided by key: a load which is non-conforming in the PWB is left out of the EPC's totals too, whatever the EPC flag says.  `areas` notes the load's non-conforming MW. `scale_nonconforming = true` in `[area]` scales them like the rest. Topology mode is unaffected: it sets every load to its own EPC value.
  - **Key basis** (see [Key Basis](#key-basis)): within each group, a load or gen whose key is in both cases starts from its EPC value, the rest from their seed values, and that is what gets scaled.
  - What a bucket's own units can't take (no PWB units, or a target above the bucket's MWMax) is reported as `Unmet MW` in `areas`. With `[area] makeup` (on), the plan puts it onto the area's other GenTypes (online units with a positive target, in proportion to their planned MW, capped at MWMax; what is still left going up re-dispatches the area's Thermal and Unknown GenTypes which have offline units, with a bigger target, so units come online by the [Unit Commitment](#unit-commitment) rules; Wind, Solar, Hydro and Storage units are never brought online this way), so the area's total generation still meets its EPC total (a negative shortfall, e.g. storage charging more than the PWB units can take, brings the area's other gens down; Wind and Solar move like any other type); `areas` shows that as `Make-up MW` on the groups which took it. What the whole area can't take, and load that can't be placed (an area with no in-service PWB load of the sign needed), ends up on the swing.
- Scale the case to the targets with the adaptive stepping engine (see [Iterate to Gen and Load Targets](#iterate-to-gen-and-load-targets)). Gens/loads which diverge are held back and retried, so `Final MW` can differ from the target: check `Error MW` and the `exceptions` sheet. The `Final` columns are read back from PowerWorld after scaling.
- **Islands** (both modes). When the targets are built, the seed is solved and every gen/load on a bus whose `IslandNumber` isn't 1 (another island, or blank) is taken out of scaling and **turned off**: written Open at 0 MW once, before scaling (each write solve-checked; one which diverges stays on, and the console names it). That includes a second island's own slack gens and DC ties; never the giant swing. They take no share of the targets, so the main island's units carry each area's whole target (area mode), and they are left out of the `areas` sheet's Seed / Final MW like non-conforming loads (`Note`: `outside main island, turned off: <seed MW>`). The `exceptions` sheet lists each one Closed in the seed as `main: Island <N>, not the main island` (or `main: No island`), in the plan too. The pre-balance, the loss estimate and the balances leave them out. If the seed doesn't solve, or no swing bus is in island 1, islands aren't checked (printed).
- **Balance the swing** (`balance_swing` in `settings.toml`, both modes; `Scripts/area.py`). The swing unit itself is never moved: it solves to whatever it needs, taking up every difference including losses. What it carries beyond its goal is what gets moved onto the rest of the case's generation. The goal: the giant swing `xx` aims at 0 MW (it shouldn't exist in the final case). Without one, the case's own slack gen(s) aim at their EPC MW when their key is in the EPC (0 if Open there), else anywhere within their MWMin to MWMax. The pre-balance counts the slack's move from its seed MW to its EPC MW, so the other gens make it up; load is never moved, so each area's load stays on its EPC total. With the swing at +X MW, every Closed gen with MW > 0 (not on the slack bus) goes up by one factor (X in total), so every area & GenType group ends about the same share above (or below) its target. Only its movable units take a share (Closed, MW > 0, not held, excluded or outside the voltage band), so a group with many such units moves a little less. What a group can't take (its units at MWMax going up, or at 0 going down) goes onto the other GenTypes of its area first, in proportion to MW, then onto every gen in the case. What no gen can take stays on the swing. This happens:
  - **Before scaling** (`prebalance_targets()`): the targets themselves usually don't balance: the change in load − gen − dist gen + losses from the seed to the targets (e.g. 12 GW) would land on the swing a little at every step. Gens, loads and their dist gen count only while Closed (an Open unit injects nothing, whatever its MW). In area mode, each area's losses (PowerWorld's `Area` `LossMW` in the seed) are expected to go as (its target load / its seed load) ^ `loss_exponent` (2, in `[area]`; 0 keeps them at seed), counting positive load only and capping the ratio at 2. The total is folded into the gen targets first by the same rule (one factor on every included gen's target, capped at MWMax, spilling to the area then the case; load targets unchanged), so the swing stays put while scaling. The `runs` sheet lists it as `prebalance`, and `areas` counts it in `Balance MW`.
  - **During scaling** (`midrun_balancer()`): after each step, if the swing is outside its band. The case's own slack: the pre-balance plans it to move from where it starts to its goal as the scaling goes, so at each step it is expected at that point of the way; its band is that point ±`balance_trigger_mw`, within its MWMin to MWMax but at least half of `balance_trigger_mw` either side of where a balance aims (so a slack with MWMin = MWMax still has a band; with `balance_trigger_mw` 0 the band is its MWMin to MWMax, and a range under 100 MW is not balanced during scaling), and a balance brings it back to that point, kept at least half of `balance_trigger_mw` (or a quarter of its range) inside its limits so a slack planned at its Pmax still has room to drift (a slack far off its plan, ~2 GW on a large case, can diverge). After each step the balancer also measures how fast the swing drifts and caps the next step so that it can't carry the swing past its band before a balance sees it (both swings). The giant swing: the band is ±`balance_trigger_mw` (1000 MW; 0 turns it off), and/or, with `balance_at_swing_limits`, the swing unit's own MWMin to MWMax (the giant swing's are `[prep] swing_mw_min` / `swing_mw_max`, ±20000 MW). With both, the narrower one applies. The units moved get their start and target values scaled by the same factor, so the rest of the scaling keeps the balance instead of undoing it. If a balance can't bring the swing back inside (e.g. gens at MWMax), the band widens to where the swing got to plus a margin (`balance_trigger_mw`, or 10% of the band), so the next one waits until it has drifted that much further. Gens/loads the scaling holds back or has excluded (here and in the final balance), and those on buses outside `exclude_vmin`–`exclude_vmax` or outside the main island, are left out of a balance; the rest take their share, and what gens at MWMax can't take is spread over the others. A balance aims the swing at 0 MW kept within the middle half of the band: with a ±1000 MW band that is 0, with a swing unit's 0 to 826 MW it is about 207 MW (so the next small drop doesn't set off another balance). A balance which fails part-way keeps what it applied and the hour carries on; its row in `runs` gives the error in `Note`. The `runs` sheet lists each balance as `midrun<N>`. A gen's target is never raised above its `MWMax` by a balance, and a gen the scaling takes to 0 MW or below (decommitting) is left out of a balance, since the rest of the scaling would undo its share.
  - **Make-up** (area mode, `[area] makeup`, on by default; `makeup()`): after the main scaling, before the final balance. Per area & GenType (and per area's load), what the excluded or still-held units fell short of their targets goes onto the same group's other units: Closed gens with MW > 0, in proportion to MW, capped at MWMax (floor 0 going down); positive loads, one factor on MW and Mvar. What a GenType's own units can't take goes onto the area's gens of every other GenType, the same way, so the area's total generation still meets its target; what the whole area can't take is left to the swing (and so the final balance). Islanded units, non-conforming loads, swing gens and units outside the voltage band are left out. The `runs` row `makeup` notes the MW asked for and placed; `areas` shows it as `Make-up MW`. It keeps a group's shortfall in that group, or at least in its area, instead of leaving it to the swing (and so to the balance, spread over everyone).
  - **After scaling** (`balance_swing()`): round after round until the swing is within `balance_tol_mw` (1 MW) of its goal (0 for the giant swing; the slack's EPC MW, or its MW limits), at most `balance_max_rounds` (10) times, since losses move with each round. The `runs` sheet lists each round as `balance<N>`.

  Either way the area gen totals move off their targets by what was balanced (load totals don't), which `Balance MW` and `Error MW` show.
- Save the case and the log. Then the engineer handles the [Manual Work](#manual-work).


### Unit Commitment
`commit` in `[area]` of `settings.toml` (on). These dispatch rules apply to an area & GenType group with nothing to scale: a target of 0, or no online MW in the target's direction (otherwise see [Key Basis](#key-basis)). They also pick which offline units come online when a scaled group is capped at MWMax.
- Hydro (run of the river), Wind, Solar and Battery Storage (`all_on_types`): every unit online, all at the same % of MWMax (of `|MWMin|` when storage is charging). At a target of 0, Wind and Solar go offline (`off_at_zero_types`); Hydro and Storage keep their seed statuses, at 0 MW.
- Thermal (and Unknown): start from the units online in the seed. Units online in the EPC (same key) are preferred. Scaling up: offline units which are online in the EPC come online first, then the biggest, until the target is at most `commit_max_cf` (0.9) of the online MWMax. Scaling down: online units which are offline in the EPC go offline first, then the smallest, as long as the rest still cover the target within that 0.9. A target of 0 takes them all offline. The online units then run at the same % of MWMax.
- Units with no capacity in the target's direction (e.g. synchronous condensers, MWMax 0) keep their seed status.
- The scaling engine closes newly committed units in at 0 MW and ramps them up, and ramps decommitted units down to 0 MW before opening them. A unit which diverges while closing or opening is excluded and left where it is.
- `areas` shows `Online Seed` and `Online Final` per group (STATCOMs added during scaling are not counted), and notes `committed by capacity rules`.
- MWMin above 0 is not enforced: a committed unit can be planned below its MWMin.
- With `commit = false` every gen keeps its seed status, and only online units share the target.

### Key Basis
Loads and gens are matched key by key where they can be, and scaled where they can't:
- **Loads**, per area, positive and negative separately, MW and Mvar separately: each Closed PWB load starts from its EPC value if its key is in the EPC (0 if it is Open there), else from its seed value (S + I + Z). That basis is scaled by one factor so the group meets the EPC total. Where every key matches, the factor is 1 and every load gets its own EPC value.
- **Gens**, per area & GenType: each unit starts from its EPC status (with `commit`) and MW if its key is in the EPC, else from its seed status and MW. The online units' basis is scaled by one factor to the group target, capped at MWMax; what a capped unit can't take is spread over the others. If all are capped, the rest brings offline units online by the [Unit Commitment](#unit-commitment) rules. The whole group scales, so matched units move off their exact EPC values by that factor, and scaling down doesn't take units offline: a group scaled below 30% of its basis (`LOW_FACTOR_WARNING`) prints a warning, since its units stay online at low output. Groups with no matching key scale from their seed dispatch the same way (decided by Fable 2026-11-15: the seed is a solved operating point, scaling it stays close to that, and one rule for every group is predictable). Only a group with nothing to scale (a target of 0, or no online MW in the target's direction) is committed and dispatched by the Unit Commitment rules.

### GenType
`Scripts/area.py` maps the letter codes PowerWorld uses (`UnitTypeCode` e.g. `HY`, `FuelTypeCode` e.g. `NG`) to the generic types. The maps follow the groups in CaseBridge's `core/schemas/codes/generator_codes.json`, which also lists each code's EPC and PowerWorld integers. Where this project departs from those groups, the entry is marked `DECIDE` in the code:
- `PS` (pumped storage) is Hydro (CaseBridge: Storage).
- `CP` (concentrated solar) is Solar (CaseBridge: Storage).
- `SC`, `SV`, `MP` (condensers, SVCs, motor/pumps) are Unknown whatever their fuel.
- `DC` (DC ties) are not scaled at all: `manual_unit_types` in `settings.toml`.
- Battery Storage = `BA`, `FW`, `ES`, `CE` units, and the `MWH` fuel.
- To change a mapping, or give pumped storage its own group, use `[area.unit_gentype]` / `[area.fuel_gentype]` and `gen_types` in `settings.toml`. Both cases are always classified by the same maps.
- Codes in neither map are Unknown. The `areas` sheet shows the MW in each bucket, so an Unknown bucket with a lot of MW means a code is missing from the maps. If more than `max_unknown_share` (0.5, in `[area]`; 1.0 turns the check off) of the EPC's generation is Unknown, the run stops before scaling and names the MW.

### Open Questions (area mode)
- Do the GridView EPC exports keep the unit and fuel types, so that PowerWorld reads them into `UnitTypeCode` / `FuelTypeCode`? If they come back blank, every EPC gen is Unknown, and the run stops before scaling (`max_unknown_share`).
- The `DECIDE` entries above.
- GridView's X1/X2 distributed generation loads (in the EPC, not in the PWB) count in the area load target: positive ones add to the area's positive target, negative ones are netted into it. 
- Gen sharing by `MWMax` ignores `MWMin` above 0, and never exceeds `MWMax`. Should a target above the bucket's capacity push units past `MWMax` instead of going to the swing?
- `AreaNumber` is the gen's / load's own area field. Should it be the bus's area?

## Topology Mode

### Methodology Summary
- Setup GridView and PWB cases to have the same Dummy-Bus-Numbers for Multi-Section Lines. 
- Bring all missing elements from GridView into the PWB. 
  - Gets all X1/X2/X3 etc loads from all GridView EPCs into the prepared seed (GridView seems to creates loads to represent distributed generation on-the-fly, differently in each hour case)
- Fix bad transformer taps (If taps are > 15% from Nominal kV, set to 1.0 Taps). 
- Make PWB branch statuses match GridView. 
  - Ignore unsupported GridView branch types (`UNSUPPORTED_GV_BRANCH_TYPES` in `Scripts/topo.py`), such as Breakers/Disconnects. 
- Put all generators in terminal voltage-control mode, at current solved Vpu as the setpoint. 
- Place a giant slack on the case's existing swing bus (or another bus: `swing_bus`), to balance while scaling. Turn off Area AGC in the solution settings. 
- Calculate target Load/Gen MW, MVAR, & Status. 
- Optional (off by default): Do a PVQV sensitivity test. If dvdp and dvdq tests indicate a > 10% change in voltage, exclude those targets from scaling. 
- Optional (off by default): Test all gen targets 1 at a time. If they diverge the base-case, exclude them from scaling. 
- Close all loads & gens which will get scaled. 
- Scale all loads & gens to their targets in adaptive steps: start at 25% (a balance, a small move, starts at 100%), double after each step which solves at its first size (one which only solved once halved keeps its size), halve after each which diverges (down to 1%). A 1% step which still diverges is bisected, and the gens/loads which diverge are held back. After each step: 
  - If Bus Vpu > 8% high/low? Adjust shunt Capacitors/Reactors.
  - If Bus Vpu > 12% high/low? Hold loads/gens back from scaling until it recovers. 
  - Retry held loads/gens. Those which solve catch up with the rest.
  - If Bus Vpu < 85%? Create an infinite STATCOM on that bus. 
  - If Bus Vpu < 80%? Drop branches to isolate collapsed area. 
- Retry held loads/gens once more, bisecting. Exclude those which still diverge, where they got to. 
- Set final load & gen statuses 1 at a time. Roll back upon any divergent edits. 
- Balance the swing (`balance_swing`, on by default): this moves gens off their exact EPC values (one factor on all of them, up to MWMax, spread by area & GenType as in area mode; loads keep theirs) until the swing is at its goal, as in area mode (the giant swing 0 MW; the case's own slack its EPC MW, within its limits; the same `area.swing_epc()` rule as area mode, built from the main island's gens only, so another island's slack has no goal). Set `balance_swing = false` in `settings.toml` for the exact EPC values, and the swing then keeps the imbalance. The gens on the PWB's slack bus are left alone in this mode too.
- Save resulting model, and all logs. 
- Engineer to manually adjust, as desired:
  - Phase Shifters
  - DC Lines in `MTDCConverter` and `DCTransmissionLine` tables.
  - Generation/load until the `ID`=`xx` giant swing generator is close to 0MW, and can be deleted from the case, replacing it with another system swing.
  - Areas where STATCOMs have been added to support voltage. 
  - Areas where branches have been dropped to isolate collapsed system elements. 
  - Areas where Gens/Loads have been excluded from scaling. 

### Topological Seed
The first step is to produce a solved PWB which has the same topology as the GridView model EPCs. 

#### Dummy Buses
There are some key differences in methodology for how PowerWorld and GE PSLF handle Multi-Section (MS) Lines. 
- PowerWorld, Branch: Key = [From, To, Circuit]
- GE PLSF, Line: Key = [From, To, Circuit, Section]

A line in PSLF may be represented like:
- `From Bus` --- SE1 --- SE2 --- SE3 --- `To Bus`

In PSLF, the middle-buses do not get assigned numbers or names. However, when PowerWorld reads an EPC file, it automatically creates middle buses between each line section. 
- `From Bus` --- SE1 --- `Middle Bus` --- SE2 --- `Middle Bus` --- SE3 --- `To Bus`

The bus numbers get auto-assigned based on your settings. When performing topological comparisons between models, if those middle bus numbers do not align, it will look like there are new or deleted branches to be handled, since the keys [From, To, Circuit] do not match. To get two models to have the same middle-bus-numbers, one could perform such a task by hand as follows: 
- Open Base Case file in Powerworld
- Click Aggregations>Multi-section lines 
- Click Save Auxiliary File (all records and columns) or hit Ctrl + Alt + A
- Save as DummyBus.aux and choose Number as the identifier
- Open the aux file in a text-editor and replace `&` signs with a space and replace `<SUBDATA Bus>` with `<SUBDATA BusRenumber>`
- Save DummyBus.aux file
- Open Other Case
- Hit edit mode and load DummyBus.aux file 

`create_dummy_bus_aux()` creates the DummyBus.aux file following the above steps. 

#### Missing Elements
It's common for engineers to add generation, load, lines, transformers, etc into a GridView model, which may not have been present in the original PWB case which was the seed to the GridView model. Those additional elements must be brought into the PWB file prior to proceeding. 

`create_missing_elements()` was designed with this in mind. Given the full topological details of a left and right model from `get_case_data()`, this function will modify the currently open case (the "right" model) to have all elements from the "left" model. The routine adds those elements in an initially "Out of Service" state, so the model can be immediately solved. 

#### Transformer Taps
Since GridView does not take voltage into account, transformer taps from a GridView model may be very far from nominal voltages and therefore divergent when added into a powerflow model. `fix_transformer_taps()` was designed to resolve this issue. Any transformers with taps further than a threshold from nominal kV (`tap_threshold`, default 0.15) will be set to 1.0 taps on the from & to sides. 

#### Branch Statuses
Branch statuses are set to match those which come from GridView in an iterative fashion in `set_branch_statuses()`. If an element exists in the PWB model but not the GridView EPC, those must be turned off. Then for all matching branches, those statuses must be modified to match GridView. The routine will first attempt to modify all branch statuses at once; if that fails, it will split the changes in half and retry each half, recursively, down to individual branches. If any individual branch status operation fails (diverges), the system state is rolled back and that branch is listed in a list of failed status change attempts. 

Note that in GridView, some branch types may go missing from the exported EPCs. This may include:
- Breakers
- Disconnects
- Fuses
- Ground Disconnects
- Load Break Disconnects
- ZBR

As such, those branch types are excluded when setting branch statuses to match the model topology. The list is `unsupported_gv_branch_types` in `settings.toml`. 

## Prepare Seed
### Adjust Shunts
GridView does not take voltage into consideration. As such, shunt statuses may need to be modified to ensure good system voltage. `adjust_shunts()` will toggle shunt statuses Open/Closed based on the voltage on the bus. By default, vlow=0.92, vhigh=1.08. 

### Generation Terminal Voltage Control
It is common practice to set generation voltage controls to use a remote Potential Transformer (PT) such as one on a high-side bus, or measure the low-side (terminal) bus voltage and calculate the high-side bus voltage through an R & X compensation factor. 

However, from a Newton-Raphson solution stability standpoint, remote regulation settings in a PowerFlow model has several downsides:
- The solution may cause the terminal bus voltage to be very high/low. 
- The solution may cause several generators which are near eachother to have circulating VARs.
- If the generator is set to control a very high fault-duty bus, then any fractional changes to that bus voltage solution will cause the generator to throw from Max MVAR output to Min MVAR output iteratively between solves. 

All of the above reasons can cause significant instability in solving Newton-Raphson. Furthermore, generation voltage setpoints may get modified season to season as their voltage-support needs change operationally throughout the year. This means that for a targeted GridView hour, those voltage setpoints are quite possibly going to be different than the base-case PWB model. 

For all of the above reasons, prior to tuning the case to match a GridView hour, the `Scripts/GenTerminalVoltageControl.aux` script is run on the prepared seed case to set all generation to terminal voltage control mode. The voltage target is set equal to the current voltage solution in the base-case, so that there should be no change to system voltages pre/post change. However, if terminal voltages are >1.1 PU or <0.9 PU, then the setpoint is clamped to 1.1 PU or 0.9 PU respectively to ensure stable solutions. The original regulated bus and setpoint are stored in `CustomInteger:0` and `CustomFloat:0`. 

### Get Fault Duty & PVQV Sensitivities
Set by `swing_bus = "max_fault_duty"` (off by default, since it is slow) and `compute_pvqv` (on by default) in `settings.toml`. 
- Fault Duty: is computed on all buses, and saved in `Prepared/fault_duty.csv`. 
- PVQV: Bus voltage sensitivity to additional MW or MVAR (linearly) is computed and saved to `Prepared/pvqv.csv`. 

### Add a Giant Swing/Slack Generator
Since the GridView EPC exports are unsolved powerflow models, they will not have accurate accounting for system losses. In systems such as the WECC (with peak loads > 100 GW), even a few percent difference in losses could be several GW of power imbalance, which has to be placed somewhere in the model until the user can later distinguish where to balance the difference. As such, a giant swing/slack generator is created and placed on the case's existing swing bus (or the maximum fault-duty bus, or a given bus: `swing_bus`) to allow such balancing to occur. Area AGC is turned off either way. With `create_giant_swing = false`, the case's own swing does this instead; it must be able to take up the imbalance while the case is scaled.

## Load and Gen Scaling
Scaling starts from `Prepared/Prepared.pwb`. Both modes use the same engine, in `Scripts/scale.py`. It scales constant power, current and impedance load, and distributed generation. In area mode, statuses change only through gen commitment.

### Compute Targets
Topology mode: `topo.compute_pw_targets()` puts the PowerWorld case side by side with the GridView values for each gen and load. Area mode: `area.build_targets()`, see [Area Mode](#area-mode).

### Test Gen Targets
In GridView, a resource of virtually any size can be placed virtually anywhere, regardless of system impedances. Since GridView doesn't solve powerflows, there would be no issue if we placed a 500MW generator on a 5MVA transformer. However, this would not work when we move to solving a powerflow case. What makes this even more challenging, is when scaling generation up linearly, it can be hard to identify such circumstances, since the generator is likely holding the voltage constant. There could be no indication of a solution stability problem until the divergence occurs. To identify these situations ahead of time, we need to test each generation target ahead of time. 

`test_gen_targets_parallel()` opens several instance of PowerWorld, to split this task up. Each instance of PowerWorld opens the Prepared.pwb case, then begins testing each target. It applies the generation target, solves, then rolls it back, and keeps track of whether the solve was a "Success". Those generation targets which are not successful individually should be excluded from the Load & Gen scaling process. 

This is `pretest_gens = true` in `settings.toml`. It is off by default: instead, gens/loads which diverge are found while scaling, by bisecting the step (see below). That avoids the extra PowerWorld instances. 

### Iterate to Gen and Load Targets
`iterate_to_gen_load_targets()` contains the logic to linearly scale loads and gens to their final targets. 
- `compute_pvqv_exclusions()` (only with `use_pvqv = True`) excludes Loads & Gens from scaling, if the $\Delta P$ and $\Delta Q$ indicates more than a 10% $\Delta V$ from a linear PVQV analysis standpoint. 
  - Variables:
    - $\Delta P$ = $P_{final} - P_{initial}$
    - $\Delta Q$ = $Q_{final} - Q_{initial}$
    - PVQV Results provide linear estimates of $\frac{dV}{dP}$ and $\frac{dV}{dQ}$
  - Gen Check: 
    - $|\Delta P \times \frac{dV}{dP}| > 10\%$ 
  - Load Check:
    - $|\Delta P \times \frac{dV}{dP} + \Delta Q \times \frac{dV}{dQ}| > 10\%$ 
- `hold_voltage()` holds Loads & Gens back from scaling while their respective buses exceed +/- 12% of nominal voltage (rows already at their target are not held). 
- `adjust_shunts()` opens/closes shunt capacitors/reactors when voltages reach +/- 8% of nominal. 
- `close_all_related_gen_load()` closes gens/loads which are open now but closed in the target. In area mode, these are the gens committed online.
- `create_statcom_on_lowestv_bus()` will place a large STATCOM (`ID`=`SC`) on a bus which is beginning to see signs of collapse (<85% of Nominal Voltage). A bus which already has one (from any run, including a balance) is skipped. The `exceptions` sheet lists each STATCOM (Table `STATCOM`, the run in `Reason`) and whether it could be closed in, and likewise each dropped branch (Table `Branch`). 
- `drop_collapsed_sections()` is a last-ditch effort which drops branches on network sections which are showing voltage collapse (<80% Voltage). Sometimes, Newton-Raphson solves into a state where network sections have 0.2 PU voltages and it still shows as solved. This is intended to catch those situations to some degree, by dropping those problematic areas to be resolved later by engineering review. 

The routine moves the loads and generators to their final targets in adaptive steps, measured in 1/100ths (`RESOLUTION`) of the way from start to target, while using the above techniques to avoid case divergence along the way. Tuning is in `[scale]` of `settings.toml`. 
- The first step is `INITIAL_STEP` (25); a swing balance's own scaling starts at `MAX_STEP`, since it moves units by a few percent. Each step which solves at its first size doubles the next, up to `MAX_STEP` (100); one which only solved once halved keeps its size, since doubling straight back would diverge again. Each which diverges is rolled back and retried at half the size.
- A 1/100 step which still diverges is bisected (applied in halves, recursively) to find the gens/loads which cause divergence. Those are held back at their last good values, and the rest go ahead. The step then returns to the size which first diverged.
- Held gens/loads aren't excluded straight away, since a unit which diverges early may solve after other changes have taken effect. After each step, `catch_up()` tries to move them all up to one step closer to the rest (in one attempt). Those which catch up resume scaling.
- Once the rest reach their targets, `retry_held()` retries the held gens/loads, bisecting, starting with a jump to target and halving the jump when nothing moves. Those which still diverge are excluded, left partway between their start and target values.
- The `runs` sheet counts each run's steps and diverged steps.

At the end of the scaling process, it will set the final statuses on loads/gens. 

## Manual Work

### DC Converters
There are three tables with DC converters: 
- Load/gen: The load/gen tables contain DC converters which go to adjacent areas. Gens with unit type `DC` are not scaled. Converters in the load table are scaled with their area's positive or negative load, by their sign. 
- MTDCConverter: Multi-Terminal DC Line setpoints do not appear to get modified by GridView in the EPC exports. 
- DCTransmissionLine: Two-Terminal DC Lines may get the from/to buses swapped by GridView depending on the modeled flows. Since this looks like a topological change from an element-key standpoint, these were not touched by the routines. It is intended that the engineer manually edit the DC setpoints to their desired values to ensure correct DC converter modeling. 

### Phase Shifters
Phase shifter angles are not adjusted in the EPCs exported by GridView. 

### Remove Infinite Swing
Step 02 balances the swing to within 1 MW of 0 (see [Methodology Summary](#methodology-summary)). If it gets there, the giant swing (`ID`=`xx`) can be deleted and the case's own swing restored. If a round is cut short (gens at MWMax, gens/loads which diverge), the swing keeps the rest.

## Future: CaseBridge instead of PowerWorld in the middle
See [docs/casebridge-route.md](docs/casebridge-route.md).

---

Ben Hutchins (WPP) designed and scoped this application; the implementation was done with an AI-assist.
