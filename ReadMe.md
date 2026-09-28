# Overview

## Purpose
Given (Input): 
- A solved seed-case PowerWorld Binary (PWB) file that was used to generate the GridView models, and
- EPC(s) which have been exported from GridView for one or more hour snapshots. 

Produce (Output):
- A solved PWB for every provided GridView hour, which contains the same:
  - Topology
    - Buses
    - Transformers, Lines, & Line Shunts
    - Loads & Generators
  - System State
    - Load MW, MVAR, Status
    - Generation MW, Status
    - Branch Statuses

## Status
The restructured scripts (`Scripts/pw.py`, `topo.py`, `scale.py`) have not yet been run against PowerWorld. The offline tests in `tests/test_lib.py` pass.

## Prerequisites
- PowerWorld v23+ with a SimAuto license
- Python 3.9+
- Python libraries: `pip install -r requirements.txt`

## How to Use
- Ensure PowerWorld Simulator is installed, with a license for SimAuto. 
- Install Visual Studio Code and Python. 
- Open `Workspace.code-workspace` 
- Place GridView EPC Exports into `./HourEPCs/`, and remove any sample cases. 
- Place your PWB seed-case into `./Seed/`, and remove any sample cases. 
- `01 Topological Seed.py` uses the first `.pwb` in `./Seed/` and the first `.epc` (alphabetically) in `./HourEPCs/` to create the topological seed-case. 
- Run `01 Topological Seed.py`
- Review the results in `./TopoSeed/` to ensure that `TopoSeed.PWB` is satisfactorily matching your GridView topology. 
- Run `02 Load and Gen Scaling.py`
  - `pretest_gens` and `use_pvqv` at the top are off by default. Gens/loads which diverge are found by bisecting the scaling step instead. See [Test Gen Targets](#test-gen-targets).
  - Note: It may take ~1-2 hours per EPC converted, depending on your CPU's speed. 
  - EPCs which already have an output `.pwb` in `./Output/` are skipped, so the script can be re-run after a failure. Delete an output `.pwb` to redo that hour.
  - If one EPC fails, the error is printed and the remaining EPCs continue.
- Run `03 Merge Reports.py`
- Review `03 Merge Reports.xlsx` to see what has been adjusted. 
- Tuning thresholds (voltage bands, iteration count, worker count, etc.) are at the top of `Scripts/topo.py` (step 01) and `Scripts/scale.py` (step 02). 
- Code layout: `Scripts/pw.py` talks to PowerWorld, `Scripts/topo.py` holds the step 01 operations, and `Scripts/scale.py` holds the step 02 operations. `tests/test_lib.py` runs without PowerWorld: `python tests/test_lib.py`. 

# Process Notes

## Methodology Summary
- Setup GridView and PWB cases to have the same Dummy-Bus-Numbers for Multi-Section Lines. 
- Bring all missing elements from GridView into the PWB. 
  - Gets all X1/X2/X3 etc loads from all GridView EPCs into the TopoSeed (GridView seems to creates loads to represent distributed generation on-the-fly, differently in each hour case)
- Fix bad transformer taps (If taps are > 15% from Nominal kV, set to 1.0 Taps). 
- Make PWB branch statuses match GridView. 
  - Ignore unsupported GridView branch types (`UNSUPPORTED_GV_BRANCH_TYPES` in `Scripts/topo.py`), such as Breakers/Disconnects. 
- Put all generators in terminal voltage-control mode, at current solved Vpu as the setpoint. 
- Place a giant slack on the case's existing swing bus (or the max fault-duty MVA bus, with `compute_fault_duty`), to balance while scaling. Turn off Area AGC in the solution settings. 
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

## Topological Seed
The first step is to produce a solved PWB which has the same topology as the GridView model EPCs. 

### Dummy Buses
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

### Missing Elements
It's common for engineers to add generation, load, lines, transformers, etc into a GridView model, which may not have been present in the original PWB case which was the seed to the GridView model. Those additional elements must be brought into the PWB file prior to proceeding. 

`create_missing_elements()` was designed with this in mind. Given the full topological details of a left and right model from `get_case_data()`, this function will modify the currently open case (the "right" model) to have all elements from the "left" model. The routine adds those elements in an initially "Out of Service" state, so the model can be immediately solved. 

### Transformer Taps
Since GridView does not take voltage into account, transformer taps from a GridView model may be very far from nominal voltages and therefore divergent when added into a powerflow model. `fix_transformer_taps()` was designed to resolve this issue. Any transformers with taps further than a threshold from nominal kV (default = 0.15) will be set to 1.0 taps on the from & to sides. 

### Branch Statuses
Branch statuses are set to match those which come from GridView in an iterative fashion in `set_branch_statuses()`. If an element exists in the PWB model but not the GridView EPC, those must be turned off. Then for all matching branches, those statuses must be modified to match GridView. The routine will first attempt to modify all branch statuses at once; if that fails, it will split the changes in half and retry each half, recursively, down to individual branches. If any individual branch status operation fails (diverges), the system state is rolled back and that branch is listed in a list of failed status change attempts. 

Note that in GridView, some branch types may go missing from the exported EPCs. This may include:
- Breakers
- Disconnects
- Fuses
- Ground Disconnects
- Load Break Disconnects
- ZBR

As such, those branch types are excluded when setting branch statuses to match the model topology. The list is `UNSUPPORTED_GV_BRANCH_TYPES` in `Scripts/topo.py`. 

### Adjust Shunts
GridView does not take voltage into consideration. As such, shunt statuses may need to be modified to ensure good system voltage. `adjust_shunts()` will toggle shunt statuses Open/Closed based on the voltage on the bus. By default, vlow=0.92, vhigh=1.08. 

### Generation Terminal Voltage Control
It is common practice to set generation voltage controls to use a remote Potential Transformer (PT) such as one on a high-side bus, or measure the low-side (terminal) bus voltage and calculate the high-side bus voltage through an R & X compensation factor. 

However, from a Newton-Raphson solution stability standpoint, remote regulation settings in a PowerFlow model has several downsides:
- The solution may cause the terminal bus voltage to be very high/low. 
- The solution may cause several generators which are near eachother to have circulating VARs.
- If the generator is set to control a very high fault-duty bus, then any fractional changes to that bus voltage solution will cause the generator to throw from Max MVAR output to Min MVAR output iteratively between solves. 

All of the above reasons can cause significant instability in solving Newton-Raphson. Furthermore, generation voltage setpoints may get modified season to season as their voltage-support needs change operationally throughout the year. This means that for a targeted GridView hour, those voltage setpoints are quite possibly going to be different than the base-case PWB model. 

For all of the above reasons, prior to tuning the case to match a GridView hour, the `Scripts/GenTerminalVoltageControl.aux` script is run on the topological seed case to set all generation to terminal voltage control mode. The voltage target is set equal to the current voltage solution in the base-case, so that there should be no change to system voltages pre/post change. However, if terminal voltages are >1.1 PU or <0.9 PU, then the setpoint is clamped to 1.1 PU or 0.9 PU respectively to ensure stable solutions. The original regulated bus and setpoint are stored in `CustomInteger:0` and `CustomFloat:0`. 

### Get Fault Duty & PVQV Sensitivities
Set by `compute_fault_duty` (off by default, since it is slow) and `compute_pvqv` (on by default) in `01 Topological Seed.py`. 
- Fault Duty: is computed on all buses, and saved in `TopoSeed/fault_duty.csv`. 
- PVQV: Bus voltage sensitivity to additional MW or MVAR (linearly) is computed and saved to `TopoSeed/pvqv.csv`. 

### Add a Giant Swing/Slack Generator
Since the GridView EPC exports are unsolved powerflow models, they will not have accurate accounting for system losses. In systems such as the WECC (with peak loads > 100 GW), even a few percent difference in losses could be several GW of power imbalance, which has to be placed somewhere in the model until the user can later distinguish where to balance the difference. As such, a giant swing/slack generator is created and placed on the case's existing swing bus (or the maximum fault-duty bus, with `compute_fault_duty`) to allow such balancing to occur. 

## Load and Gen Scaling
Load and generation scaling starts from the Topological seed. 

### Compute Targets
`compute_pw_targets()` will put the right (PowerWorld) case side by side with the left (Target / GridView) case values for Loads and Gens. 

### Test Gen Targets
In GridView, a resource of virtually any size can be placed virtually anywhere, regardless of system impedances. Since GridView doesn't solve powerflows, there would be no issue if we placed a 500MW generator on a 5MVA transformer. However, this would not work when we move to solving a powerflow case. What makes this even more challenging, is when scaling generation up linearly, it can be hard to identify such circumstances, since the generator is likely holding the voltage constant. There could be no indication of a solution stability problem until the divergence occurs. To identify these situations ahead of time, we need to test each generation target ahead of time. 

`test_gen_targets_parallel()` opens several instance of PowerWorld, to split this task up. Each instance of PowerWorld opens the TopoSeed.pwb case, then begins testing each target. It applies the generation target, solves, then rolls it back, and keeps track of whether the solve was a "Success". Those generation targets which are not successful individually should be excluded from the Load & Gen scaling process. 

This is `pretest_gens = True` in `02 Load and Gen Scaling.py`. It is off by default: instead, gens/loads which diverge are found while scaling, by bisecting the step (see below). That avoids the extra PowerWorld instances. 

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
- `close_all_related_gen_load()` closes all generation and load which is to be scaled up/down. If the element is in-service in the target (GridView EPC), or in-service in the source (TopoSeed.PWB), then it gets closed by this routine. Any loads/gens which were previously out of service are closed in at 0 MW / 0 MVAR to begin the scaling. 
- `create_statcom_on_lowestv_bus()` will place a large STATCOM (`ID`=`SC`) on a bus which is beginning to see signs of collapse (<85% of Nominal Voltage). The scale log lists each STATCOM bus, and whether the STATCOM could be closed in. 
- `drop_collapsed_sections()` is a last-ditch effort which drops branches on network sections which are showing voltage collapse (<80% Voltage). Sometimes, Newton-Raphson solves into a state where network sections have 0.2 PU voltages and it still shows as solved. This is intended to catch those situations to some degree, by dropping those problematic areas to be resolved later by engineering review. 

The routine moves the loads and generators to their final targets in adaptive steps, measured in 1/100ths (`RESOLUTION`) of the way from start to target, while using the above techniques to avoid case divergence along the way. Tuning is at the top of `Scripts/scale.py`. 
- The first step is `INITIAL_STEP` (25). Each step which solves doubles the next, up to `MAX_STEP` (100). Each which diverges is rolled back and retried at half the size.
- A 1/100 step which still diverges is bisected (applied in halves, recursively) to find the gens/loads which cause divergence. Those are held back at their last good values, and the rest go ahead. The step then returns to the size which first diverged.
- Held gens/loads aren't excluded straight away, since a unit which diverges early may solve after other changes have taken effect. After each step, `catch_up()` tries to move them all up to one step closer to the rest (in one attempt). Those which catch up resume scaling.
- Once the rest reach their targets, `retry_held()` retries the held gens/loads, bisecting, starting with a jump to target and halving the jump when nothing moves. Those which still diverge are excluded, left partway between their start and target values.
- The `step_df` sheet of the scale log lists each step, its result, how many held gens/loads caught up, and how many are still held.

At the end of the scaling process, it will set the final statuses on loads/gens. 

## Manual Work

### DC Converters
There are three tables with DC converters: 
- Load/gen: The load/gen tables contains DC converters which go to adjacent areas. These are captured in the load/gen scaling process. 
- MTDCConverter: Multi-Terminal DC Line setpoints do not appear to get modified by GridView in the EPC exports. 
- DCTransmissionLine: Two-Terminal DC Lines may get the from/to buses swapped by GridView depending on the modeled flows. Since this looks like a topological change from an element-key standpoint, these were not touched by the routines. It is intended that the engineer manually edit the DC setpoints to their desired values to ensure correct DC converter modeling. 

### Phase Shifters
Phase shifter angles are not adjusted in the EPCs exported by GridView. 

### Remove Infinite Swing
It is intended that the study engineer manually work to rebalance the case as needed to get the infinite swing machine to 0 MW, then remove it from the model. 

