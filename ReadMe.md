# Overview

## Purpose
Given (Input):
- A solved seed-case PowerWorld Binary (PWB) file that was used to generate the GridView models, and
- EPC(s) which have been exported from GridView for one or more hour snapshots.

Produce (Output): a solved PWB for every provided GridView hour. Two modes, set by `mode` in `settings.toml`:
- **`area`** (default): the same load MW & Mvar in each area, and the same generation MW in each area for each generation type (Solar, Hydro, Wind, Thermal, Battery Storage, Unknown). Topology is not matched: nothing is created or deleted, and no branch or load is opened or closed. Gens are brought online or taken offline only to suit their area & type target ([Unit Commitment](#unit-commitment)). Areas whose gens (or loads) have exactly the same keys in both cases are matched unit by unit.
- **`topology`**: the same topology (buses, branches, loads, gens, and their statuses) and the same MW/Mvar on every load and gen. Missing elements are created in the PWB.

## Status
Area mode: first draft (2026-11-15), first written in `~/projects/areascale` (branch `areascale-draft` here). Neither mode's current code has been run against PowerWorld yet. The offline tests in `tests/test_lib.py` pass.

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
  - `<hour>_01_Target.xlsx`: the per-element `load` / `gen` targets. Area mode adds `area_load` and `area_gen` (target vs seed vs planned per area / per area & type, with Unmet MW, and gens committed / decommitted).
  - `<hour>_03_ScaleLog.xlsx`: the scaling logs. Area mode adds the area sheets with what the case holds after scaling (`Final MW`, `Final Error MW`). `balance_df` lists the swing balancing rounds.
  - `<hour>.pwb`. EPCs which already have an output `.pwb` are skipped; delete it to redo that hour. If one EPC fails, the error is printed and the rest continue.
  - Note: It may take ~1-2 hours per EPC, depending on your CPU's speed.
- Run `03 Merge Reports.py` to merge every hour's scale log into `03 Merge Reports.xlsx`.
- Code layout: `Scripts/pw.py` talks to PowerWorld. `Scripts/prep.py` holds the seed preparation both modes share, `Scripts/topo.py` the topology mode's matching and targets, and `Scripts/area.py` the area mode's targets and the swing balancing. `Scripts/scale.py` moves the case to the targets. `Scripts/settings.py` reads `settings.toml`. `tests/test_lib.py` runs without PowerWorld: `python tests/test_lib.py`.

## Settings
`settings.toml`, at the top of the repo, with a comment on each setting. `[run]` holds the switches the two step scripts read:
- `mode`: `"area"` or `"topology"`.
- `create_giant_swing`: add a giant swing (`ID`=`xx`), or use the case's own swing (the gens on its slack bus). `swing_bus`: `"slack"`, `"max_fault_duty"` or a bus number.
- `balance_swing`: move MW off the swing onto the whole case, half gen and half load, during scaling whenever the swing leaves its band (`[area] balance_trigger_mw` and/or the swing unit's MW limits, `balance_at_swing_limits`), and after scaling until it is ~0 MW. The swing unit itself is never moved, in either mode.
- `adjust_seed_shunts`, `gen_terminal_voltage_control`, `compute_pvqv`: the seed preparation steps.
- `pretest_gens`, `use_pvqv`: optional scaling exclusions (both off).

The other sections (`[area]`, `[scale]`, `[prep]`, `[topology]`, `[pw]`) set the tuning constants of the `Scripts/` file they name: e.g. `[scale] exclude_vmin` sets `EXCLUDE_VMIN` in `Scripts/scale.py`. The code holds the defaults, so a setting left out of the file keeps its default, and a misspelt one is an error. `[area.unit_gentype]` / `[area.fuel_gentype]` add to or override the GenType maps.

# Process Notes

## Area Mode

### Methodology Summary
- Prepare the seed (step 01): see [Prepare Seed](#prepare-seed). The topology is left as it is.
- Read the loads and gens of the GridView EPC and of the prepared PWB (both opened in PowerWorld).
- Compute targets (`Scripts/area.py`):
  - **Load, per `AreaNumber`.** Target = the EPC's total MW and Mvar of Closed loads (constant power + current + impedance; the EPC's distributed generation is ignored). In the PWB, on every Closed load: distributed generation goes Open & 0, constant current and impedance go to 0, and constant power MW / Mvar take the load's S + I + Z total, scaled. Positive and negative load are matched separately: the area's positive loads are scaled to the EPC's positive total for that area, and its negative loads (DC converters, generation modelled as load) to the EPC's negative total. That is four factors per area: MW +, MW −, Mvar +, Mvar −. If no seed load in the area has Mvar of a sign the EPC needs, that Mvar is spread over the area's zero-Mvar loads in proportion to |MW|. `area_load` shows each part (`Target MW +`, `Unmet MW -`, ...). Open loads aren't scaled, but their distributed generation is also set Open & 0.
  - **Generation, per `AreaNumber` & GenType.** Each gen's GenType comes from its `UnitTypeCode`, or from its `FuelTypeCode` when the unit type is unstated (`UN`, `OT`, `FC` or blank); see [GenType](#gentype). Target = the EPC's total `MWSetPoint` of Closed gens in that bucket. The bucket's units are committed ([Unit Commitment](#unit-commitment)), then the online ones share the target in proportion to `MWMax` (the same capacity factor on each), capped at `MWMax`. A negative target (net storage charging) is shared in proportion to `|MWMin|`. The giant swing and DC ties (unit type `DC`, set by hand) are never scaled, and are left out of both cases' totals.
  - **Non-conforming loads** (PowerWorld's load `AGC` field = NO: RAW's SCALE flag 0, and presumably an EPC's `nonc` flag 1) are never scaled or balanced, and are left out of both cases' area totals, so conforming load is matched against conforming load. `area_load` shows each case's non-conforming MW. `scale_nonconforming = true` in `[area]` scales them like the rest. Topology mode is unaffected: it sets every load to its own EPC value.
  - **Matched keys.** If an area has exactly the same list of loads (by key) in both cases, each of its loads targets its own EPC value instead (0 where it is Open in the EPC). The same holds for gens, separately. Load statuses aren't changed; with `commit`, gens take their EPC status.
  - What can't be placed (a bucket with no PWB units, a target above the bucket's MWMax, an area with no in-service PWB load of the sign needed) is reported as `Unmet MW` / `Unmet Mvar`, and ends up on the giant swing.
- Scale the case to the targets with the adaptive stepping engine (see [Iterate to Gen and Load Targets](#iterate-to-gen-and-load-targets)). Gens/loads which diverge are held back and retried, so `Final MW` can differ from the target: check `Final Error MW`. The `Final` columns are read back from PowerWorld after scaling.
- **Balance the swing** (`balance_swing` in `settings.toml`, both modes; `Scripts/area.py`). The swing unit itself is never moved: it solves to whatever it needs, taking up every difference including losses. Its output (the giant swing `xx`, or if there is none, the gens on the slack bus) is what gets moved onto the rest of the case, half onto generation and half onto load. With the swing at +X MW, every Closed gen with MW > 0 (not on the slack bus) goes up by one factor (X/2 in total, capped at MWMax), and every Closed load with MW > 0 goes down by one factor (X/2 in total, Mvar with it). This happens:
  - **During scaling** (`midrun_balancer()`): after each step, if the swing is outside its band. The band is ±`balance_trigger_mw` (1000 MW; 0 turns it off), and/or, with `balance_at_swing_limits`, the swing unit's own MWMin to MWMax (the giant swing's are `[prep] swing_mw_min` / `swing_mw_max`, ±20000 MW). With both, the narrower one applies. The units moved get their start and target values scaled by the same factor, so the rest of the scaling keeps the balance instead of undoing it. If a balance can't bring the swing back inside (e.g. gens at MWMax), the band widens to where the swing got to plus a margin (`balance_trigger_mw`, or 10% of the band), so the next one waits until it has drifted that much further. The `midrun_balance_df` sheet lists each one.
  - **After scaling** (`balance_swing()`): round after round until the swing is within `balance_tol_mw` (1 MW) of 0, at most `balance_max_rounds` (10) times, since losses move with each round. The `balance_df` sheet lists each round.

  Either way the area totals move off their targets by what was balanced, which `Final Error MW` shows.
- Save the case and the logs. Then the engineer handles the [Manual Work](#manual-work).


### Unit Commitment
`commit` in `[area]` of `settings.toml` (on). Per area & GenType bucket, outside matched-key areas:
- Hydro (run of the river), Wind, Solar and Battery Storage (`all_on_types`): every unit online, all at the same % of MWMax (of `|MWMin|` when storage is charging). At a target of 0, Wind and Solar go offline (`off_at_zero_types`); Hydro and Storage keep their seed statuses, at 0 MW.
- Thermal (and Unknown): start from the units online in the seed. Units online in the EPC (same key) are preferred. Scaling up: offline units which are online in the EPC come online first, then the biggest, until the target is at most `commit_max_cf` (0.9) of the online MWMax. Scaling down: online units which are offline in the EPC go offline first, then the smallest, as long as the rest still cover the target within that 0.9. A target of 0 takes them all offline. The online units then run at the same % of MWMax.
- Units with no capacity in the target's direction (e.g. synchronous condensers, MWMax 0) keep their seed status.
- In a matched-key area, each gen takes its EPC status.
- The scaling engine closes newly committed units in at 0 MW and ramps them up, and ramps decommitted units down to 0 MW before opening them. A unit which diverges while closing or opening is excluded and left where it is.
- `area_gen` shows `Seed Online`, `Planned Online`, `Committed`, `Decommitted` and the online MWMax per bucket.
- MWMin above 0 is not enforced: a committed unit can be planned below its MWMin.
- With `commit = false` every gen keeps its seed status, and only online units share the target.

### GenType
`Scripts/area.py` maps the letter codes PowerWorld uses (`UnitTypeCode` e.g. `HY`, `FuelTypeCode` e.g. `NG`) to the generic types. The maps follow the groups in CaseBridge's `core/schemas/codes/generator_codes.json`, which also lists each code's EPC and PowerWorld integers. Where this project departs from those groups, the entry is marked `DECIDE` in the code:
- `PS` (pumped storage) is Hydro (CaseBridge: Storage).
- `CP` (concentrated solar) is Solar (CaseBridge: Storage).
- `SC`, `SV`, `MP` (condensers, SVCs, motor/pumps) are Unknown whatever their fuel.
- `DC` (DC ties) are not scaled at all: `manual_unit_types` in `settings.toml`.
- Battery Storage = `BA`, `FW`, `ES`, `CE` units, and the `MWH` fuel.
- To change a mapping, or give pumped storage its own group, use `[area.unit_gentype]` / `[area.fuel_gentype]` and `gen_types` in `settings.toml`. Both cases are always classified by the same maps.
- Codes in neither map are Unknown. The `area_gen` sheet shows the MW in each bucket, so an Unknown bucket with a lot of MW means a code is missing from the maps.

### Open Questions (area mode)
- Do the GridView EPC exports keep the unit and fuel types, so that PowerWorld reads them into `UnitTypeCode` / `FuelTypeCode`? If they come back blank, every EPC gen is Unknown.
- The `DECIDE` entries above.
- Should the GridView X1/X2 distributed generation loads (in the EPC, not in the PWB) count in the area load target? Today they do (net load is matched). Their MW shows as `EPC-only MW` in `area_load`.
- Gen sharing by `MWMax` ignores `MWMin` above 0, and never exceeds `MWMax`. Should a target above the bucket's capacity push units past `MWMax` instead of going to the swing?
- `AreaNumber` is the gen's / load's own area field. Should it be the bus's area?
- The scaling engine's solve aids do change the network: it switches shunts, may add STATCOMs (`ID`=`SC`) and may open branches into collapsed sections. These are not there to match GridView, only to keep the case solving. `add_statcoms` and `drop_collapsed` in `[scale]` of `settings.toml` turn the last two off. Should they be off by default here?

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
- Scale all loads & gens to their targets in adaptive steps: start at 25%, double after each step which solves, halve after each which diverges (down to 1%). A 1% step which still diverges is bisected, and the gens/loads which diverge are held back. After each step: 
  - If Bus Vpu > 8% high/low? Adjust shunt Capacitors/Reactors.
  - If Bus Vpu > 12% high/low? Hold loads/gens back from scaling until it recovers. 
  - Retry held loads/gens. Those which solve catch up with the rest.
  - If Bus Vpu < 85%? Create an infinite STATCOM on that bus. 
  - If Bus Vpu < 80%? Drop branches to isolate collapsed area. 
- Retry held loads/gens once more, bisecting. Exclude those which still diverge, where they got to. 
- Set final load & gen statuses 1 at a time. Roll back upon any divergent edits. 
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
- `hold_voltage()` holds Loads & Gens back from scaling while their respective buses exceed +/- 12% of nominal voltage. 
- `adjust_shunts()` opens/closes shunt capacitors/reactors when voltages reach +/- 8% of nominal. 
- `close_all_related_gen_load()` closes gens/loads which are open now but closed in the target. In area mode, these are the gens committed online.
- `create_statcom_on_lowestv_bus()` will place a large STATCOM (`ID`=`SC`) on a bus which is beginning to see signs of collapse (<85% of Nominal Voltage). The scale log lists each STATCOM bus, and whether the STATCOM could be closed in. 
- `drop_collapsed_sections()` is a last-ditch effort which drops branches on network sections which are showing voltage collapse (<80% Voltage). Sometimes, Newton-Raphson solves into a state where network sections have 0.2 PU voltages and it still shows as solved. This is intended to catch those situations to some degree, by dropping those problematic areas to be resolved later by engineering review. 

The routine moves the loads and generators to their final targets in adaptive steps, measured in 1/100ths (`RESOLUTION`) of the way from start to target, while using the above techniques to avoid case divergence along the way. Tuning is in `[scale]` of `settings.toml`. 
- The first step is `INITIAL_STEP` (25). Each step which solves doubles the next, up to `MAX_STEP` (100). Each which diverges is rolled back and retried at half the size.
- A 1/100 step which still diverges is bisected (applied in halves, recursively) to find the gens/loads which cause divergence. Those are held back at their last good values, and the rest go ahead. The step then returns to the size which first diverged.
- Held gens/loads aren't excluded straight away, since a unit which diverges early may solve after other changes have taken effect. After each step, `catch_up()` tries to move them all up to one step closer to the rest (in one attempt). Those which catch up resume scaling.
- Once the rest reach their targets, `retry_held()` retries the held gens/loads, bisecting, starting with a jump to target and halving the jump when nothing moves. Those which still diverge are excluded, left partway between their start and target values.
- The `step_df` sheet of the scale log lists each step, its result, how many held gens/loads caught up, and how many are still held.

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
