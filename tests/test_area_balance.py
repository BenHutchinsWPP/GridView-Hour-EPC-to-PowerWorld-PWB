"""
Tests for area.py swing balances, midrun, make-up, final balance and the areas ledger. Needs no PowerWorld.
    python tests/test_area_balance.py      (all tests: python tests/run.py [name part])
"""
import contextlib
import io
from typing import Any
import numpy as np
import pandas as pd
from helpers import cell, FakeSimAuto, LOAD_P_ALL, SwingSimAuto, gens, loads, saved_settings, swing_case, targets_df
from Scripts import pw, prep, scale, area

def test_swing_mw():
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 2, '1', 'Closed', 'GT', 0.0, 100.0, 0.0)]).assign(MW=[120.0, 50.0], BusNum=[7, 8])
    assert area.swing_mw(gen, slack_buses=[8]) == (120.0, [7])                  # The giant swing wins.
    assert area.swing_mw(gen[gen['ID'] != 'xx'], slack_buses=[8]) == (50.0, [8])  # Else gens on slack buses.

def test_balance_targets_moves_gen_only():
    # Swing at +100 MW: gens up 100 MW by one factor (50 each); load untouched. g2 (area 2) caps at MWMax 120 and area 2
    # has no other gen, so its last 30 MW go onto g1.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 2, '2', 'Closed', 'HY', 100.0, 120.0, 0.0),
                ('neg', 2, '3', 'Closed', 'BA', -10.0, 50.0, -50.0), ('sw', 3, '1', 'Closed', 'GT', 500.0, 1e4, 0.0),
                ('dc', 3, '2', 'Closed', 'DC', 90.0, 500.0, -500.0)]).assign(BusNum=[1, 2, 2, 9, 3])
    load = loads([('l1', 1, 'Closed', 150.0, 0.0, 0.0, 30.0), ('l2', 2, 'Closed', 350.0, 0.0, 0.0, 70.0),
                  ('lneg', 2, 'Closed', -40.0, 0.0, 0.0, 0.0), ('lo', 2, 'Open', 99.0, 0.0, 0.0, 0.0)])
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[9])
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 180.0, 'g2': 120.0}
    assert t['Load'].empty and (t['Gen']['MWSetPoint_Target'] - t['Gen']['MWSetPoint_Seed']).sum() == 100.0
    area.rescale_targets({'Gen': targets_df(ObjectID=['g1'], MWSetPoint=[150.0], MWSetPoint_Target=[150.0]),
                          'Load': targets_df(ObjectID=['l1'], **{c: [1.0] for c in area.LOAD_P + area.LOAD_Q})}, t)   # Empty load table is fine.

def test_spread_by_group_fills_the_group_then_its_area_then_everyone():
    # Area 1: Thermal t1 (100, cap 110), Hydro h1 (100, cap 1000). Area 2: Thermal t2 (100, cap 1000).
    mw = pd.Series([100.0, 100.0, 100.0], index=['t1', 'h1', 't2'])
    cap = pd.Series([110.0, 1000.0, 1000.0], index=mw.index)
    area_, group = pd.Series([1, 1, 2], index=mw.index), pd.Series(['Thermal', 'Hydro', 'Thermal'], index=mw.index)
    ask = pd.Series({(1, 'Thermal'): 50.0})
    x, left = area.spread_by_group(ask, mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': 10.0, 'h1': 40.0, 't2': 0.0} and left == 0.0       # Group first, then its area.
    cap['h1'] = 120.0
    x, left = area.spread_by_group(ask, mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': 10.0, 'h1': 20.0, 't2': 0.0} and left == 20.0      # The area is full: stays short.
    x, left = area.spread_by_group(ask, mw, cap, area_, group)
    assert x.to_dict() == {'t1': 10.0, 'h1': 20.0, 't2': 20.0} and left == 0.0      # spill_all: everyone else.
    x, left = area.spread_by_group(pd.Series({(1, 'Thermal'): -150.0}), mw, cap, area_, group, spill_all=False)
    assert x.to_dict() == {'t1': -100.0, 'h1': -50.0, 't2': 0.0} and left == 0.0   # Down: floor 0, then the area.
    x, _ = area.spread_evenly(-30.0, mw, cap, area_, group)
    assert x.round(9).to_dict() == {'t1': -10.0, 'h1': -10.0, 't2': -10.0}        # One factor on every group.
    x, left = area.spread_evenly(50.0, mw.iloc[:0], cap.iloc[:0], area_.iloc[:0], group.iloc[:0])
    assert left == 50.0                                                             # Nothing to place it on.

def test_spread_by_group_goes_by_mw_and_keeps_opposite_moves_apart():
    # By MW, not per group: t (100) and h (300) share +40 as 10 / 30.
    mw = pd.Series([100.0, 300.0], index=['t', 'h'])
    x, _ = area.spread_evenly(40.0, mw, pd.Series(1e4, index=mw.index), pd.Series(1, index=mw.index), pd.Series(['Thermal', 'Hydro'], index=mw.index))
    assert x.round(9).to_dict() == {'t': 10.0, 'h': 30.0}
    # Thermal (at MWMax) asks +50, Hydro asks -30, Wind has room. Thermal's 50 go onto Wind only, never back onto Hydro.
    mw = pd.Series([100.0, 100.0, 100.0], index=['t1', 'h1', 'w1'])
    cap = pd.Series([100.0, 1000.0, 1000.0], index=mw.index)
    group = pd.Series(['Thermal', 'Hydro', 'Wind'], index=mw.index)
    x, left = area.spread_by_group(pd.Series({(1, 'Thermal'): 50.0, (1, 'Hydro'): -30.0}), mw, cap,
                                   pd.Series(1, index=mw.index), group, spill_all=False)
    assert x.round(9).to_dict() == {'t1': 0.0, 'h1': -30.0, 'w1': 50.0} and left == 0.0

def test_balance_skips_held_and_out_of_band_units():
    # FakeSimAuto has buses 1 and 2 at 1.0 PU; bus 3 has no voltage, so it counts as outside the band.
    gen = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 300.0, 0.0), ('g2', 1, '2', 'Closed', 'GT', 100.0, 300.0, 0.0),
                ('g3', 1, '3', 'Closed', 'GT', 100.0, 300.0, 0.0)]).assign(BusNum=[1, 2, 3])
    load = loads([('l1', 1, 'Closed', 100.0, 0.0, 0.0, 0.0), ('l2', 1, 'Closed', 100.0, 0.0, 0.0, 0.0)])
    main = {'Gen': gen.set_index('ObjectID').assign(ExclusionReason='', HeldReason=['', 'Voltage', '']),
            'Load': load.set_index('ObjectID').assign(ExclusionReason=['', 'PVQV'], HeldReason='')}
    skip = area.balance_skips(FakeSimAuto(), {'Gen': gen, 'Load': load}, main)
    assert skip == {'g2', 'g3', 'l2'}
    t = area.balance_targets({'Gen': gen, 'Load': load}, swing=100.0, swing_buses=[], skip=skip)
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'g1': 200.0}     # The whole swing lands on what is left.
    assert t['Load'].empty

def test_midrun_balancer_keeps_the_swing_near_zero_and_never_moves_it():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    # Scale load from 200 to 3200 MW with gens fixed: without balancing, the swing would end at +3000 MW.
    targets = load_ramp_targets()
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    # The scaling log of every balance is kept in its run.
    assert [r.name for r in log] == [f'midrun{n}' for n in range(1, len(log) + 1)] and all(len(r.log['step_df']) for r in log)
    swing = lambda: area.swing_mw(sa.table('Gen'), [])[0]
    assert len(log) >= 1 and all(abs(r.after) < abs(r.before) for r in log)
    assert abs(swing()) <= area.BALANCE_TRIGGER_MW                      # Balanced as it went.
    assert ('Gen', '99') not in sa.applied                              # The swing unit was never written.
    # Balancing moves gen only: the load reached its own target.
    assert targets['Load'].loc['l1', 'SMW_Target'] == 3200 and abs(cell(targets['Load'], 'l1', 'SMW') - 3200) < 1e-6
    area.balance_swing(sa)
    assert abs(swing()) < area.BALANCE_TOL_MW and ('Gen', '99') not in sa.applied

def load_ramp_targets():
    """Load 200 -> 3200 MW with gens fixed: without balancing, the swing would end at +3000 MW."""
    return {'Gen': targets_df(ObjectID=['g1'], BusNum=[1], ID='1', Status='Closed', Status_Target='Closed',
                              MWSetPoint=[100.0], MWSetPoint_Target=[100.0]),
            'Load': targets_df(ObjectID=['l1'], BusNum=[3], ID='1', Status='Closed', Status_Target='Closed',
                               DistStatus='Open', DistStatus_Target='Open', AreaNumber=1,
                               **{c: [v] for c, v in zip(scale.SCALED['Load'], [200.0, 20.0, 0, 0, 0, 0, 0, 0])},
                               **{f'{c}_Target': [v] for c, v in zip(scale.SCALED['Load'], [3200.0, 320.0, 0, 0, 0, 0, 0, 0])})}

def test_swing_band_and_limit_trigger():
    gen, load = swing_case()
    gen.loc[gen['ID'] == 'xx', ['MWMax', 'MWMin']] = [500.0, -300.0]
    saved = area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS
    try:
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = 1000.0, True
        assert area.swing_band(gen, []) == (-300.0, 500.0)                   # The narrower of the two.
        area.BALANCE_TRIGGER_MW = 0.0
        assert area.swing_band(gen, []) == (-300.0, 500.0)
        area.BALANCE_AT_SWING_LIMITS = False
        assert area.swing_band(gen, []) == (-np.inf, np.inf)                 # Off.

        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        targets = load_ramp_targets()
        scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
        assert log == [] and targets['Load'].loc['l1', 'SMW'] == 3200      # Never balanced.

        area.BALANCE_AT_SWING_LIMITS = True                                  # Balance at the swing's MWMax (500 MW).
        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
        assert len(log) >= 1 and all(r.before > 500 for r in log)
        assert area.swing_mw(sa.table('Gen'), [])[0] <= 500
    finally:
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = saved

def test_rescale_targets():
    targets = {'Gen': targets_df(ObjectID=['g1', 'g2'], MWSetPoint=[50.0, 10.0], MWSetPoint_Start=[0.0, 10.0], MWSetPoint_Target=[100.0, 10.0]),
               'Load': targets_df(ObjectID=['l1'], **{c: [10.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']},
                                  **{f'{c}_Start': [10.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']},
                                  **{f'{c}_Target': [20.0] for c in ['SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar']})}
    moved = {'Gen': targets_df(ObjectID=['g1', 'gx'], MWSetPoint=[60.0, 5.0], MWSetPoint_Seed=[50.0, 4.0], MWMax=[500.0, 500.0])}
    area.rescale_targets(targets, moved)
    assert cell(targets['Gen'], 'g1', ['MWSetPoint', 'MWSetPoint_Start', 'MWSetPoint_Target']).tolist() == [60.0, 0.0, 120.0]
    assert targets['Gen'].loc['g2', 'MWSetPoint_Target'] == 10.0                    # Not moved.
    assert cell(targets['Load'], 'l1', ['SMvar', 'SMvar_Start', 'SMvar_Target']).tolist() == [10.0, 10.0, 20.0]   # Balances move gens only.

def test_midrun_balance_keeps_gens_within_mwmax():
    # Solar s1 (MWMax 100) has 150 MW to place in the EPC: planned at MWMax, 50 MW unmet, which loads the swing.
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('s1', 1, '1', 'Closed', 'PV', 10.0, 100.0, 0.0),
                ('t1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)]).assign(BusNum=[99, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 110.0, 0.0, 0.0, 10.0)]).assign(BusNum=3, NomkV=230.0, BusName='b', DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)
    epc_gen = gens([('e1', 1, '1', 'Closed', 'PV', 150.0, 200.0, 0.0), ('e2', 1, '2', 'Closed', 'GT', 450.0, 1000.0, 0.0)]).assign(BusNum=[1, 2])
    epc_load = loads([('e1', 1, 'Closed', 600.0, 0.0, 0.0, 60.0)])
    sa = SwingSimAuto(gen, load)
    with saved_settings():
        area.BALANCE_TRIGGER_MW = 20.0
        area.MAKEUP = False             # Else the plan puts s1's 50 MW on t1 and the swing never moves.
        targets, _ = area.build_targets({'Gen': epc_gen, 'Load': epc_load}, area.read_case(sa), swing_buses=[99])
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
        area.balance_swing(sa)
    assert len(log) >= 1
    final = sa.table('Gen').set_index('ObjectID')
    assert final.loc['s1', 'MWSetPoint'] <= final.loc['s1', 'MWMax'] + 1e-9
    assert (targets['Gen']['MWSetPoint_Target'] <= targets['Gen']['MWMax'] + 1e-9).all()

def test_balance_df_counts_the_whole_load_change():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load.assign(IMW=100.0))      # 300 MW of load against 200 MW of gen: the swing is at +100.
    first = scale.runs_sheet(area.balance_swing(sa)).iloc[0]
    assert first['Run'] == 'balance1' and abs(first['Swing MW Before'] - 100) < 1e-9
    assert first['Load MW Moved'] == 0 and abs(first['Gen MW Moved'] - 100) < 1e-6    # All onto gen, I part included.

def test_slack_bus_gens_are_the_swing_when_there_is_no_giant_swing():
    # Bus 9 is the slack bus: its unit carries the imbalance, and neither balance writes to it.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    assert prep.slack_buses(sa) == [9]
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert len(log) >= 1 and ('Gen', '9') not in sa.applied
    area.balance_swing(sa)
    # s1's key isn't in the EPC: it only has to end within its limits (0 to 1000 MW), not at 0.
    assert -area.BALANCE_TOL_MW < area.swing_mw(sa.table('Gen'), [9])[0] < 1000 + area.BALANCE_TOL_MW
    assert ('Gen', '9') not in sa.applied

def test_own_slack_with_wide_limits_is_kept_near_its_plan_while_scaling():
    # s1 can run -10 to +10 GW, but strays of ~2 GW diverge: it is balanced back once it is balance_trigger_mw (1000)
    # from where the plan expects it (here its start, 0 MW: its key isn't in the EPC). Load +3000 MW, no pre-balance.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 0.0, 1e4, -1e4), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa)
    seen = []
    def watch(SimAuto, targets, progress):
        seen.append(area.swing_mw(sa.table('Gen'), [9])[0])         # Before any balance: what the step left.
        return after_step(SimAuto, targets, progress)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=watch)
    assert len(log) >= 1 and all(abs(r.after) < 1.0 for r in log)                      # Back to its plan, 0 MW.
    assert max(abs(x) for x in seen) < 1100 and ('Gen', '9') not in sa.applied           # The step is capped too.
    # With an EPC goal, the plan moves from start to goal as scaling goes: at 50 of 100, half way.
    swing = pd.DataFrame({'Seed MW': [0.0], 'EPC MW': [800.0]}, index=['s1'])
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa, swing)
    sa.base['Load']['SMW'] += 2000.0                                                    # Swing jumps to ~2000.
    after_step(sa, {t: df.assign(HeldReason='') for t, df in load_ramp_targets().items()}, 50)
    assert len(log) == 1 and abs(log[0].after - 400.0) < 1.0

def test_own_slack_planned_at_its_limit_does_not_balance_every_step():
    # s1 is at its MWMax (0 MW) and stays planned there; load +3000 MW, no pre-balance. Balancing aims 500 MW inside
    # (half of balance_trigger_mw), so the drift has room: a few balances, not one per 1% step.
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 0.0, 0.0, -1e4), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa)
    out = scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert 1 <= len(log) <= 10 and len(out['step_df']) <= 25, (len(log), len(out['step_df']))
    assert all(abs(r.after + 500.0) < 1.0 for r in log)                                # 500 MW inside its MWMax.

def test_own_slack_ends_at_its_epc_mw_when_its_key_is_in_the_epc():
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('s1', 1, '1', 'Closed', 'GT', 300.0, 1000.0, 0.0)]).assign(BusNum=9)
    swing = area.swing_epc(epc_gen, gen.assign(MW=[150.0, 0.0, 0.0]), [9])
    assert swing.to_dict('index') == {'s1': {'Seed MW': 150.0, 'EPC MW': 300.0}}
    assert area.swing_goal(gen, [9], swing) == (300.0, 300.0) and area.swing_goal(gen, [9]) == (0.0, 1000.0)
    giant = pd.concat([gen, gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4)]).assign(BusNum=9)])
    assert area.swing_epc(epc_gen, giant, [9]).empty and area.swing_goal(giant, [9], swing) == (0.0, 0.0)
    sa = SwingSimAuto(gen, load, slack=[9])
    after_step, log = area.midrun_balancer(sa, swing)
    scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    area.balance_swing(sa, epc=swing)
    assert abs(area.swing_mw(sa.table('Gen'), [9])[0] - 300.0) < area.BALANCE_TOL_MW and ('Gen', '9') not in sa.applied
    # The goal stays within the slack's limits, 0 (not a positive MWMin) to MWMax: an EPC MW past MWMax is capped, and a
    # slack Open in the EPC aims at 0 although its MWMin is 50.
    lim = gen.assign(MWMin=[50.0, 0.0, 0.0])
    assert area.swing_goal(lim, [9], pd.DataFrame({'EPC MW': [1500.0]}, index=['s1'])) == (1000.0, 1000.0)
    assert area.swing_epc(epc_gen.assign(Status='Open'), lim.assign(MW=150.0), [9]).loc['s1', 'EPC MW'] == 0.0
    assert area.swing_goal(lim.assign(MWMax=np.nan), [9]) == (0.0, np.inf)                # A blank limit: unlimited.
    assert area.swing_goal(lim.assign(MWMin=-200.0), [9]) == (-200.0, 1000.0)            # A negative MWMin: kept.
    assert area.swing_band(lim.assign(MWMin=2000.0), [9]) == (0.0, 2000.0)                # Swapped limits: in order.
    # The pre-balance makes up the slack's move too: seed 150 -> EPC 300 leaves 150 MW less for the other gens.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 200.0
    t['Gen']['MWSetPoint'], t['Gen']['MWSetPoint_Target'], t['Gen']['MWMax'] = 400.0, 400.0, 1000.0
    area.prebalance_targets(t, swing_change=-150.0)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 250.0 and t['Load'].loc['l1', 'SMW_Target'] == 200.0

def ledger_run(mutate=None, slack=(99,), makeup=False):
    """Midrun + final balances, a diverging gen (g3, bus 4) and the areas sheet of the result. slack: slack buses (99 has xx)."""
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'HY', 100.0, 5000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('g1', 1, '1', 'Closed', 'GT', 200.0, 1000.0, 0.0), ('g2', 1, '2', 'Closed', 'HY', 150.0, 1000.0, 0.0),
                    ('g3', 1, '3', 'Closed', 'GT', 200.0, 1000.0, 0.0)])
    epc_load = loads([('l1', 1, 'Closed', 3200.0, 0.0, 0.0, 320.0)])
    sa = SwingSimAuto(gen, load, slack=[b for b in slack if b != 99])
    sa.bad = {'4'}                                              # Any write to g3 diverges: it is held back.
    targets, reports = area.build_targets({'Load': epc_load, 'Gen': epc_gen}, {'Load': load, 'Gen': gen}, swing_buses=list(slack))
    after_step, midrun_log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    mk = area.makeup(sa, targets) if makeup else []
    assert mk or not makeup
    runs = [*midrun_log, *mk, *area.balance_swing(sa, targets)]
    assert midrun_log and any(r.moved is not None for r in runs if r.name.startswith('balance'))
    return area.areas(area.read_case(sa), reports, targets, runs, swing_buses=list(slack)), targets

def test_areas_ledger_holds_with_midrun_and_final_balances_and_a_diverging_gen():
    out, targets = ledger_run()
    assert out.columns.tolist() == area.AREAS_COLUMNS and out.groupby(['AreaNumber', 'Group']).size().max() == 1
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out
    row = lambda g: out[out['Group'] == g].iloc[0]
    assert row('Thermal')['Not Reached MW'] > 0 and row('Thermal')['Balance MW'] != 0     # g3 held at 100 of 200 MW; g1 balanced.
    assert row('Load')['Balance MW'] == 0 and pd.isna(row('Thermal')['Target Mvar']) and row('Load')['Target Mvar'] == 320
    assert (out['Note'] == '').all() and (out['Online Final'] > 0).all()

def test_areas_before_scaling_has_blank_final_columns():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)])
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 150.0, 1000.0, 0.0), ('s1', 1, '2', 'Closed', 'PV', 40.0, 100.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case})
    out = area.areas(None, reports)
    assert out[['Group', 'Target MW', 'Unmet MW']].values.tolist() == [['Solar', 40.0, 40.0], ['Thermal', 150.0, 0.0]]
    assert out['Final MW'].isna().all() and (out['Not Reached MW'] == 0).all() and out.loc[0, 'Note'] == 'no PWB units'

def test_midrun_balance_aims_inside_the_band():
    # Swing limits 0..800 MW: 0 is the band's edge, so a balance aims at 200 MW (a quarter of the band in), not 0.
    gen, load = swing_case()
    gen.loc[gen['ID'] == 'xx', ['MWMax', 'MWMin']] = [800.0, 0.0]
    with saved_settings():
        area.BALANCE_TRIGGER_MW, area.BALANCE_AT_SWING_LIMITS = 1000.0, True
        sa = SwingSimAuto(gen, load)
        after_step, log = area.midrun_balancer(sa)
        scale.iterate_to_gen_load_targets(sa, load_ramp_targets(), after_step=after_step)
    assert len(log) >= 1 and abs(log[0].after - 200.0) < 1.0

def test_final_balance_keeps_the_hour_when_a_round_fails():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 900.0                       # Swing at +700 MW.
    def fails(*args, **kwargs):
        raise pw.SimAutoError('Could not solve the original input case.')
    saved = scale.iterate_to_gen_load_targets
    scale.iterate_to_gen_load_targets = fails
    try:
        runs = area.balance_swing(sa)
    finally:
        scale.iterate_to_gen_load_targets = saved
    assert [r.name for r in runs] == ['balance1'] and runs[0].note.startswith('stopped: Could not solve') and runs[0].after == runs[0].after

def test_makeup_keeps_the_hour_and_its_note_when_its_scaling_fails():
    # ledger_run's case, without the midrun balancer, but the make-up's own scaling raises: it still reads After and writes the asked/placed note.
    gen = gens([('xx', 1, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'HY', 100.0, 5000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    epc_gen = gens([('g1', 1, '1', 'Closed', 'GT', 200.0, 1000.0, 0.0), ('g2', 1, '2', 'Closed', 'HY', 150.0, 1000.0, 0.0),
                    ('g3', 1, '3', 'Closed', 'GT', 200.0, 1000.0, 0.0)])
    epc_load = loads([('l1', 1, 'Closed', 3200.0, 0.0, 0.0, 320.0)])
    sa = SwingSimAuto(gen, load)
    sa.bad = {'4'}
    targets, reports = area.build_targets({'Load': epc_load, 'Gen': epc_gen}, {'Load': load, 'Gen': gen}, swing_buses=[99])
    scale.iterate_to_gen_load_targets(sa, targets)
    real, reads = scale.iterate_to_gen_load_targets, []
    def fails_for_makeup(*args, **kwargs):
        if kwargs.get('label') == 'Make-up':
            raise pw.SimAutoError('Could not solve the original input case.')
        return real(*args, **kwargs)
    real_read = area.read_case
    def counting_read(*args, **kwargs):
        reads.append(1)
        return real_read(*args, **kwargs)
    scale.iterate_to_gen_load_targets, area.read_case = fails_for_makeup, counting_read
    out = io.StringIO()
    try:
        with contextlib.redirect_stdout(out):
            runs = area.makeup(sa, targets)
    finally:
        scale.iterate_to_gen_load_targets, area.read_case = real, real_read
    assert [r.name for r in runs] == ['makeup'] and runs[0].note.startswith('gen asked') and 'stopped: Could not solve' in runs[0].note
    assert runs[0].after == runs[0].after and len(reads) == 2           # The Before read and the After read.
    assert 'WARNING: Make-up stopped: Could not solve' in out.getvalue()

def test_midrun_balance_failure_keeps_the_hour_and_its_network_log():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    run = scale.iterate_to_gen_load_targets
    def balance_fails(SimAuto, targets, **kwargs):
        e = pw.SimAutoError('Could not solve the original input case.')
        e.log = {'statcom_bus_df': pd.DataFrame({'BusNum': [7], 'Closed': [True]})}
        raise e
    scale.iterate_to_gen_load_targets = balance_fails     # Only the balances' nested runs look it up by name.
    try:
        targets = load_ramp_targets()
        run(sa, targets, after_step=after_step)
    finally:
        scale.iterate_to_gen_load_targets = run
    assert targets['Load'].loc['l1', 'Progress'] == scale.RESOLUTION                 # The main scaling finished.
    assert log and log[0].note.startswith('stopped: Could not solve')
    assert log[0].log['statcom_bus_df']['BusNum'].tolist() == [7]                   # The STATCOM it added is listed.
    ex = scale.exceptions(log, None, None)
    # Each failed balance listed it (steps are capped now, so the swing is caught again before it runs far).
    assert ex[['Table', 'ObjectID', 'BusNum']].values.tolist() == [['STATCOM', '7 SC', 7]]
    assert ex['Reason'].iloc[0].startswith('midrun1: STATCOM added')

def test_final_balance_leaves_units_the_main_scaling_excluded():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 600.0                                          # Swing at +400 MW.
    scaled = {'Gen': targets_df(ObjectID=['g1'], BusNum=[1]).assign(ExclusionReason='PVQV', HeldReason=''),
              'Load': targets_df(ObjectID=['l1'], BusNum=[3]).assign(HeldReason='')}
    area.balance_swing(sa, scaled)
    out = sa.table('Gen').set_index('ObjectID')
    assert out.loc['g1', 'MWSetPoint'] == 100.0 and out.loc['g2', 'MWSetPoint'] > 100.0

def test_midrun_balance_leaves_units_the_main_scaling_excluded():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    targets = load_ramp_targets()
    targets['Gen']['ExclusionReason'] = 'PVQV'                                     # g1: excluded by the main scaling.
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    out = sa.table('Gen').set_index('ObjectID')
    assert log and out.loc['g1', 'MWSetPoint'] == 100.0 and out.loc['g2', 'MWSetPoint'] > 100.0

def test_areas_ledger_holds_with_a_second_slack_bus():
    # Bus 2 (hydro g2) is a slack bus beside the giant swing: left out of the areas, its balance moves must be too.
    out, _ = ledger_run(slack=(99, 2))
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out

def test_prebalance_targets_moves_gen_only_and_leaves_what_gens_cannot_take_on_the_swing():
    # Load 200 -> 600 MW, gens fixed: +400 MW would land on the swing. g1 can take +50 (MWMax 150); load keeps its target,
    # and the other 350 MW stay on the swing.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 600.0
    t['Gen']['MWMax'] = 150.0
    moved = area.prebalance_targets(t)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 150.0
    assert t['Load'].loc['l1', 'SMW_Target'] == 600.0 and t['Load'].loc['l1', 'SMvar_Target'] == 320.0
    assert moved is not None
    assert area.mw_changes(moved) == {'Gen MW Change': 50.0, 'Load MW Change': 0.0}
    t['Gen']['MWMax'] = 1000.0
    area.prebalance_targets(t)
    assert t['Gen'].loc['g1', 'MWSetPoint_Target'] == 500.0
    assert area.prebalance_targets(t) is None                                     # Balanced now: nothing to move.

def test_prebalanced_targets_need_no_midrun_balance():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    targets = load_ramp_targets()                                                # Without it: +3000 MW onto the swing.
    targets['Gen']['MWMax'] = 1e4                                                # g1 100 -> 3100, l1 stays 3200.
    area.prebalance_targets(targets)
    assert targets['Gen'].loc['g1', 'MWSetPoint_Target'] == 3100.0
    after_step, runs = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, targets, after_step=after_step)
    assert runs == [] and abs(area.swing_mw(sa.table('Gen'), [])[0]) < 1e-6

def test_prebalance_targets_counts_only_what_will_move():
    # l1's dist gen (40 MW, Closed) goes to 0: +40 MW onto the swing. l2's dist gen is Open: its 30 MW never injected.
    # l3 is Open in both: its 100 MW never counts. g_open goes Closed -> Open (MW unchanged): its 50 MW leave: +50.
    # gx is excluded: not counted, not moved. gcap (target 300) is already above its MWMax (200), so it isn't raised.
    # +90 in all: g1 takes it all; loads keep their targets.
    gen = targets_df(ObjectID=['g1', 'gx', 'g_open', 'gcap'], BusNum=[1, 2, 3, 4], ID='1', Status='Closed',
                     Status_Target=['Closed', 'Closed', 'Open', 'Closed'], MWSetPoint=[100.0, 100.0, 50.0, 300.0],
                     MWSetPoint_Target=[100.0, 150.0, 50.0, 300.0], MWMax=[1000.0, 1000.0, 1000.0, 200.0])
    gen.loc['gx', 'ExclusionReason'] = 'PVQV'
    load = targets_df(ObjectID=['l1', 'l2', 'l3'], BusNum=[5, 6, 7], ID='1', Status=['Closed', 'Closed', 'Open'],
                      Status_Target=['Closed', 'Closed', 'Open'], DistStatus=['Closed', 'Open', 'Closed'], DistStatus_Target='Open',
                      **{c: [0.0] * 3 for c in scale.SCALED['Load']}, **{f'{c}_Target': [0.0] * 3 for c in scale.SCALED['Load']})
    load['SMW'], load['SMW_Target'] = 100.0, 100.0
    load['DistMWInput'] = [40.0, 30.0, 20.0]
    area.prebalance_targets({'Gen': gen, 'Load': load})
    assert gen['MWSetPoint_Target'].round(9).to_dict() == {'g1': 190.0, 'gx': 150.0, 'g_open': 50.0, 'gcap': 300.0}
    assert load['SMW_Target'].round(9).tolist() == [100.0, 100.0, 100.0]

def test_prebalance_targets_counts_losses():
    # Targets balanced by themselves; losses fall 40 MW: gens -40, load unchanged.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 1100.0
    t['Gen']['MWSetPoint_Target'], t['Gen']['MWMax'] = 1000.0, 2000.0
    assert area.prebalance_targets(t, 0.0) is None
    area.prebalance_targets(t, loss_change=-40.0)
    assert abs(cell(t['Gen'], 'g1', 'MWSetPoint_Target') - 960.0) < 1e-9
    assert t['Load'].loc['l1', 'SMW_Target'] == 1100.0

def test_loss_estimate_scales_each_area_with_its_load():
    # Area 1: 200 -> 100 MW of scaled load plus 100 MW non-conforming: 300 -> 200, losses 9 * (2/3)^2 = 4.
    # Area 2: no seed load: keeps its losses. Area 3: no loads in the targets: ratio 1.
    load = targets_df(ObjectID=['l1', 'l2', 'lo'], BusNum=[1, 2, 3], ID='1', AreaNumber=[1, 2, 1], Status=['Closed', 'Closed', 'Open'],
                      Status_Target=['Closed', 'Closed', 'Open'], **{c: [0.0] * 3 for c in LOAD_P_ALL},
                      **{f'{c}_Target': [0.0] * 3 for c in LOAD_P_ALL})
    load['SMW'], load['SMW_Target'] = [200.0, 0.0, 500.0], [100.0, 50.0, 0.0]
    area_load = pd.DataFrame({'AreaNumber': [1, 2], 'Left As Is MW': [100.0, 0.0]})
    est = area.loss_estimate({'Load': load}, area_load, pd.Series({1: 9.0, 2: 3.0, 3: 5.0}))
    assert est['Est. Loss MW'].round(9).to_dict() == {1: 4.0, 2: 3.0, 3: 5.0} and est['Seed Loss MW'].to_dict() == {1: 9.0, 2: 3.0, 3: 5.0}
    with saved_settings():
        area.LOSS_EXPONENT = 0.0
        assert (area.loss_estimate({'Load': load}, area_load, pd.Series({1: 9.0}))['Est. Loss MW'] == 9.0).all()

def test_loss_estimate_uses_positive_load_and_caps_the_ratio():
    # Area 1: 100 MW load beside a -95 MW DC load; the load goes to 150. Net 5 -> 55 would give losses x121.
    # Positive load: 100 -> 150, losses 10 * 1.5^2 = 22.5. Area 2: 10 -> 100 MW: ratio capped at 2, losses x4.
    load = targets_df(ObjectID=['l1', 'dc', 'l2'], BusNum=[1, 2, 3], ID='1', AreaNumber=[1, 1, 2], Status='Closed', Status_Target='Closed',
                      **{c: [0.0] * 3 for c in LOAD_P_ALL}, **{f'{c}_Target': [0.0] * 3 for c in LOAD_P_ALL})
    load['SMW'], load['SMW_Target'] = [100.0, -95.0, 10.0], [150.0, -95.0, 100.0]
    est = area.loss_estimate({'Load': load}, pd.DataFrame({'AreaNumber': [1, 2], 'Left As Is MW': 0.0}), pd.Series({1: 10.0, 2: 1.0}))
    assert est['Est. Loss MW'].round(9).to_dict() == {1: 22.5, 2: 4.0}

def test_areas_shows_final_losses_without_a_plan_estimate():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 1000.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': case}, {'Load': loads([]), 'Gen': case})
    out = area.areas(None, reports, final_loss=pd.Series({1: 7.0}))
    row = out[out['Group'] == 'Load'].iloc[0]
    assert len(out) == 2 and row['Final Loss MW'] == 7.0 and pd.isna(row['Seed Loss MW'])

def test_balance_skips_units_outside_the_main_island():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.islands = {2: '2', 3: ''}                                                 # g2 (bus 2) and l1 (bus 3) are islanded.
    final = area.read_case(sa)
    assert area.balance_skips(sa, final) == {'g2', 'l1'}

def test_balance_skips_no_island_when_no_bus_is_in_island_1():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.islands = {n: '' for n in (1, 2, 3, 99)}
    assert area.balance_skips(sa, area.read_case(sa)) == set()

def test_areas_ledger_holds_when_an_island_unit_stays_on():
    # h2 (island 2) couldn't be turned off: out of Not Reached and Final, and its Note says so.
    case_gen = gens([('h1', 1, '1', 'Closed', 'HY', 100.0, 200.0, 0.0), ('h2', 1, '2', 'Closed', 'HY', 100.0, 200.0, 0.0)]).assign(BusNum=[1, 2])
    epc_gen = gens([('e1', 1, '1', 'Closed', 'HY', 150.0, 200.0, 0.0)])
    targets, reports = area.build_targets({'Gen': epc_gen, 'Load': loads([])}, {'Gen': case_gen, 'Load': loads([])},
                                          island_by_bus=pd.Series([1.0, 2.0], index=[1, 2]))
    g = targets['Gen']
    g.loc['h1', 'MWSetPoint'] = 150.0                                            # h1 reached; h2 still Closed at 100.
    final = {'Gen': case_gen.assign(MWSetPoint=[150.0, 100.0]), 'Load': loads([])}
    row = area.areas(final, reports, targets).set_index('Group').loc['Hydro']
    assert (row['Not Reached MW'], row['Final MW'], row['Error MW']) == (0.0, 150.0, 0.0)
    assert 'unexplained' not in row['Note'] and 'of which 100.0 MW stayed on' in row['Note']

def test_areas_ledger_holds_with_a_makeup_and_the_held_gens_shortfall_lands_on_its_group():
    out, targets = ledger_run(makeup=True)
    ledger = out['Target MW'] - out['Unmet MW'] - out['Not Reached MW'] + out['Make-up MW'] + out['Balance MW']
    assert (abs(out['Final MW'] - ledger) < 1e-6).all(), out
    assert not out['Note'].str.contains('unexplained').any(), out['Note']
    thermal = out[out['Group'] == 'Thermal'].iloc[0]
    assert abs(thermal['Make-up MW'] - thermal['Not Reached MW']) < 1e-6 and thermal['Make-up MW'] > 0   # g1 took g3's.
    assert (out.loc[out['Group'] != 'Thermal', 'Make-up MW'] == 0).all()
    # The shortfall no longer lands on the swing, so balancing moves 100 MW less over the case.
    before, _ = ledger_run()
    assert abs(before['Balance MW'].abs().sum() - out['Balance MW'].abs().sum() - 100.0) < 1e-6

def test_makeup_targets_stay_in_the_group_and_respect_caps_islands_and_signs():
    # Area 1 Thermal: gx excluded 60 MW short; g1 (MWMax 120) and g2 share it by MW but g1 caps at 120: g1 120, g2 140.
    # gi is islanded (target 0, still at 50): not counted. Area 2 Thermal (g9) untouched. Hydro h1 isn't Thermal.
    # Load area 1: lx excluded 30 MW short; l1 (100) takes it, the negative dc load and l9 (area 2) don't.
    tg = targets_df(ObjectID=['g1', 'g2', 'gx', 'gi', 'g9', 'h1'], AreaNumber=[1, 1, 1, 1, 2, 1], BusNum=[1, 2, 3, 4, 5, 6],
                    ID='1', UnitTypeCode=['GT'] * 5 + ['HY'], FuelTypeCode='', Status='Closed',
                    MWSetPoint=[100.0, 100.0, 40.0, 50.0, 100.0, 100.0], MWSetPoint_Target=[100.0, 100.0, 100.0, 0.0, 100.0, 100.0])
    tg.loc['gx', 'ExclusionReason'], tg.loc['gi', 'ExclusionReason'] = 'Diverged during scaling', 'Island 2, not the main island'
    tl = targets_df(ObjectID=['l1', 'lx', 'dc', 'l9'], AreaNumber=[1, 1, 1, 2], BusNum=[1, 2, 3, 5], ID='1', Status='Closed',
                    **{c: 0.0 for c in LOAD_P_ALL}, **{f'{c}_Target': 0.0 for c in LOAD_P_ALL})
    tl['SMW'], tl['SMW_Target'] = [100.0, 20.0, -50.0, 100.0], [100.0, 50.0, -50.0, 100.0]
    tl.loc['lx', 'ExclusionReason'] = 'Diverged during scaling'
    final = {'Gen': tg.reset_index().assign(MWMax=[120.0, 1000.0, 1000.0, 1000.0, 1000.0, 1000.0]),
             'Load': tl.reset_index().assign(AGC='YES', DistStatus='Open')}
    made = area.makeup_targets(final, {'Gen': tg, 'Load': tl}, swing_buses=[], skip={'gx', 'gi', 'lx'})
    assert made is not None
    moved, asked = made
    assert asked == {'Gen': 60.0, 'Load': 30.0}
    assert moved['Gen']['MWSetPoint_Target'].round(9).to_dict() == {'g1': 120.0, 'g2': 140.0}
    assert moved['Load'].index.tolist() == ['l1'] and moved['Load'].loc['l1', 'SMW_Target'] == 130.0
    assert area.makeup_targets(final, {'Gen': tg.assign(ExclusionReason=''), 'Load': tl.assign(ExclusionReason='')}, []) is None
    # g1 and g2 capped at MWMax: Thermal's 60 MW go onto the area's Hydro h1, not area 2's Thermal g9.
    final['Gen']['MWMax'] = [100.0, 100.0, 1000.0, 1000.0, 1000.0, 1000.0]
    made = area.makeup_targets(final, {'Gen': tg, 'Load': tl}, swing_buses=[], skip={'gx', 'gi', 'lx'})
    assert made is not None
    moved, _ = made
    assert moved['Gen']['MWSetPoint_Target'].round(9).to_dict() == {'h1': 160.0}

def test_a_gen_above_mwmax_is_not_raised_by_any_placement():
    # In each placement a unit sits above its MWMax (600 vs 500) and the move is up: it stays at 600 and the other takes it all.
    gen = gens([('hi', 1, '1', 'Closed', 'GT', 600.0, 500.0, 0.0), ('lo', 1, '2', 'Closed', 'GT', 100.0, 1000.0, 0.0)]).assign(BusNum=[1, 2])
    # balance_targets: swing +100 MW, all of it onto lo.
    t = area.balance_targets({'Gen': gen, 'Load': loads([])}, swing=100.0, swing_buses=[])
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'hi': 600.0, 'lo': 200.0}
    # prebalance_targets: load 200 -> 400 MW puts +200 on the swing; hi's target is above its MWMax, so lo takes all 200.
    t = load_ramp_targets()
    t['Load']['SMW_Target'] = 400.0
    t['Gen'] = targets_df(ObjectID=['hi', 'lo'], BusNum=[1, 2], ID='1', Status='Closed', Status_Target='Closed',
                          MWSetPoint=[600.0, 100.0], MWSetPoint_Target=[600.0, 100.0], MWMax=[500.0, 1000.0])
    with contextlib.redirect_stdout(io.StringIO()):
        area.prebalance_targets(t)
    assert t['Gen']['MWSetPoint_Target'].to_dict() == {'hi': 600.0, 'lo': 300.0}
    # makeup_targets: gx (same group) is 60 MW short; hi can't take any, lo takes all 60.
    tg = targets_df(ObjectID=['hi', 'lo', 'gx'], AreaNumber=1, BusNum=[1, 2, 3], ID='1', UnitTypeCode='GT', FuelTypeCode='',
                    Status='Closed', MWSetPoint=[600.0, 100.0, 40.0], MWSetPoint_Target=[600.0, 100.0, 100.0])
    tg.loc['gx', 'ExclusionReason'] = 'Diverged during scaling'
    tl = targets_df(ObjectID=['l1'], AreaNumber=1, BusNum=1, ID='1', Status='Closed',
                    **{c: 0.0 for c in LOAD_P_ALL}, **{f'{c}_Target': 0.0 for c in LOAD_P_ALL})
    final = {'Gen': tg.reset_index().assign(MWMax=[500.0, 1000.0, 1000.0]), 'Load': tl.reset_index().assign(AGC='YES', DistStatus='Open')}
    made = area.makeup_targets(final, {'Gen': tg, 'Load': tl}, swing_buses=[], skip={'gx'})
    assert made is not None
    moved, asked = made
    assert asked['Gen'] == 60.0 and moved['Gen']['MWSetPoint_Target'].round(9).to_dict() == {'lo': 160.0}
    # gen_targets plan make-up: Hydro (EPC 200 MW) has one 100 MW unit, so 100 MW go onto the area's Thermal, whose unit
    # th sits at 600 above its MWMax 500 (the EPC's 700 Thermal MW leave it there): tl takes the 100 (MWMax 400, at 100).
    case = gens([('h1', 1, '1', 'Closed', 'HY', 0.0, 100.0, 0.0), ('th', 1, '2', 'Closed', 'GT', 600.0, 500.0, 0.0),
                 ('tl', 1, '3', 'Closed', 'GT', 100.0, 400.0, 0.0)])
    epc = gens([('e1', 1, '1', 'Closed', 'HY', 200.0, 500.0, 0.0), ('e2', 1, '2', 'Closed', 'GT', 700.0, 1000.0, 0.0)])
    with saved_settings():
        area.MAKEUP = True
        df, _ = area.gen_targets(epc, case)
    assert df['MWSetPoint_Target'].round(9).to_dict() == {'h1': 100.0, 'th': 600.0, 'tl': 200.0}
    assert df['Planned MW'].round(9).to_dict() == {'h1': 100.0, 'th': 600.0, 'tl': 200.0}

def test_areas_online_final_leaves_out_statcoms():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0), ('sc', 1, scale.STATCOM_ID, 'Closed', '', 0.0, 0.0, 0.0)])
    epc = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0)])
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case.iloc[:1]})
    out = area.areas({'Gen': case, 'Load': loads([]).assign(NomkV=1.0)}, reports, targets)
    assert out[['Group', 'Online Final']].values.tolist() == [['Thermal', 1.0]]

def test_midrun_balance_skips_gens_the_scaling_takes_to_zero():
    # g1 is decommitted (Open, 100 -> 0 MW), g3 stays Closed at a 0 MW target; g2 stays at 100. The balances go onto g2 only,
    # so rescale_targets() has nothing to undo and the end targets keep every MW a balance moved.
    gen = gens([('xx', 9, 'xx', 'Closed', '', 0.0, 2e4, -2e4), ('g1', 1, '1', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0), ('g3', 1, '3', 'Closed', 'GT', 100.0, 5000.0, 0.0)]
               ).assign(BusNum=[99, 1, 2, 4], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 300.0, 0.0, 0.0, 30.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    t = load_ramp_targets()
    t['Gen'] = targets_df(ObjectID=['g1', 'g2', 'g3'], BusNum=[1, 2, 4], ID='1', Status='Closed', Status_Target=['Open', 'Closed', 'Closed'],
                          MWSetPoint=[100.0] * 3, MWSetPoint_Target=[0.0, 100.0, 0.0])
    sa = SwingSimAuto(gen, load)
    after_step, log = area.midrun_balancer(sa)
    scale.iterate_to_gen_load_targets(sa, t, after_step=after_step)
    moved: list[Any] = [r.moved for r in log]      # Optional in Run; None would fail the lookups below, as before.
    assert log and all(m['Gen'].index.tolist() == ['g2'] for m in moved)
    g = t['Gen']
    assert g.loc['g1', 'MWSetPoint_Target'] == 0 and g.loc['g3', 'MWSetPoint_Target'] == 0
    kept = g['MWSetPoint_Target'].sum() - 100.0
    assert kept > 100.0 and abs(kept - sum(m['Gen']['MWSetPoint'].sum() - m['Gen']['MWSetPoint_Seed'].sum() for m in moved)) < 1e-6

def own_slack_ramp(lo, hi, trigger, down=False):
    """A 3 GW load ramp with the case's own slack (bus 9, MWMin lo, MWMax hi) and no giant swing: (midrun runs, main steps, end swing, goal)."""
    gen = gens([('s1', 1, '1', 'Closed', 'GT', 100.0, hi, lo), ('g1', 1, '2', 'Closed', 'GT', 100.0, 5000.0, 0.0),
                ('g2', 1, '3', 'Closed', 'HY', 100.0, 5000.0, 0.0)]).assign(BusNum=[9, 1, 2], MW=0.0, NomkV=230.0, BusName='b')
    load = loads([('l1', 1, 'Closed', 200.0, 0.0, 0.0, 20.0)]).assign(BusNum=3, NomkV=230.0, BusName='b')
    sa = SwingSimAuto(gen, load, slack=[9])
    if down:        # Load 3200 -> 200 instead: the slack drifts down.
        sa.base['Load'].loc[3, 'SMW'] = 3200.0
    with saved_settings(), contextlib.redirect_stdout(io.StringIO()):
        area.BALANCE_TRIGGER_MW = trigger
        after_step, log = area.midrun_balancer(sa)
        t = load_ramp_targets()
        if down:
            for c in scale.SCALED['Load']:
                t['Load'][c], t['Load'][f'{c}_Target'] = t['Load'][f'{c}_Target'].copy(), t['Load'][c].copy()
        out = scale.iterate_to_gen_load_targets(sa, t, after_step=after_step)
        area.balance_swing(sa)
    return log, out['step_df'], area.swing_mw(sa.table('Gen'), [9])[0], area.swing_goal(sa.table('Gen'), [9])

def test_midrun_balance_of_an_own_slack_with_no_range_stays_cheap():
    log, steps, swing, (lo, hi) = own_slack_ramp(100.0, 100.0, 1000.0)       # MWMin == MWMax: 0..100 MW now (no width before 0 was the floor).
    assert len(log) <= 10 and len(steps) <= 25 and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW, (len(log), len(steps), swing)
    log, steps, swing, (lo, hi) = own_slack_ramp(100.0, 100.0, 1000.0, down=True)    # The lower half of the band, on a falling load.
    assert len(log) <= 10 and len(steps) <= 25 and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW, (len(log), len(steps), swing)

def test_midrun_balance_of_an_own_slack_is_off_with_no_trigger_and_a_narrow_range():
    # Trigger 0: the band is the slack's 0 (not its MWMin 50) to MWMax; a range under 100 MW is not balanced midrun (the
    # final balance still runs). 0..90 MW here.
    log, steps, swing, (lo, hi) = own_slack_ramp(50.0, 90.0, 0.0)
    assert lo == 0.0
    assert log == [] and lo - area.BALANCE_TOL_MW <= swing <= hi + area.BALANCE_TOL_MW
    log, steps, swing, _ = own_slack_ramp(0.0, 1000.0, 0.0)                  # A wide range: balanced at its limits, not every step.
    assert 1 <= len(log) <= 5, len(log)
    log, steps, swing, _ = own_slack_ramp(50.0, 100.0, 0.0)                  # 0..100, exactly MIN_OWN_RANGE_MW wide: still balanced (strict <).
    assert len(log) >= 1

def test_areas_keeps_a_seed_gen_whose_id_is_the_statcom_id():
    case = gens([('g1', 1, '1', 'Closed', 'GT', 100.0, 200.0, 0.0), ('g2', 1, scale.STATCOM_ID, 'Closed', 'GT', 50.0, 100.0, 0.0),
                 ('new', 1, scale.STATCOM_ID, 'Closed', '', 0.0, 0.0, 0.0)])
    epc = case.iloc[:2]
    targets, reports = area.build_targets({'Load': loads([]), 'Gen': epc}, {'Load': loads([]), 'Gen': case.iloc[:2]})
    out = area.areas({'Gen': case, 'Load': loads([]).assign(NomkV=1.0)}, reports, targets).set_index('Group')
    assert out.loc['Thermal', 'Final MW'] == 150.0 and out.loc['Thermal', 'Online Final'] == 2.0   # 'new' is a STATCOM added later.
    assert not out['Note'].str.contains('unexplained').any(), out['Note']

def test_final_balance_failing_inside_its_scaling_keeps_one_row():
    gen, load = swing_case()
    sa = SwingSimAuto(gen, load)
    sa.base['Load'].loc[3, 'SMW'] = 900.0                       # Swing at +700 MW.
    real = scale.iterate_to_gen_load_targets
    def boom(*a, **k):
        raise RuntimeError('inside')
    scale.iterate_to_gen_load_targets = boom
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            runs = area.balance_swing(sa)
    finally:
        scale.iterate_to_gen_load_targets = real
    assert [r.name for r in runs] == ['balance1'] and "failed: RuntimeError('inside')" in runs[0].note
    assert runs[0].before is not None and pd.notna(runs[0].after)

if __name__ == '__main__':
    tests = [f for name, f in dict(globals()).items() if name.startswith('test_')]
    for test in tests:
        test()
    print(f'{len(tests)} tests passed.')
