"""
Step 01, both modes: seed preparation which doesn't change topology. Switches shunts for voltage, fault duty and
PVQV sensitivities, the giant swing, and turning off area AGC. (Terminal voltage control is
Scripts/GenTerminalVoltageControl.aux, loaded by step 01.) Topology matching is in topo.py.
Also turns the seed's loads into constant power, so a scaled hour moves one number per load.
"""
import pandas as pd
from . import pw

# ------------------ Tuning thresholds ------------------
SHUNT_VLOW = 0.92               # Switch shunts when bus Vpu drops below this...
SHUNT_VHIGH = 1.08              # ...or rises above this.
SHUNT_MAX_ITERATIONS = 10       # Max one-at-a-time shunt switching passes.

SWING_ID = 'xx'     # ID of the giant swing generator.
SWING_MW_MAX = 20000.0      # Giant swing MW limits. With [area] balance_at_swing_limits, scaling rebalances
SWING_MW_MIN = -20000.0     # gen & load whenever the swing goes past them.

CONVERT_TOL_MW = 1.0        # Constant-power conversion: warn when an area's (or the swing's) net MW or Mvar moves more than this.
CONVERT_MAX_FAILED = 50     # Stop bisecting after this many loads fail to convert; the rest keep their current/impedance/dist gen.

# ------------------ Shunts ------------------
def suggested_shunt_status(df: pd.DataFrame, vlow: float, vhigh: float) -> pd.Series:
    """Switch reactors out / capacitors in on low voltage, and the reverse on high voltage."""
    reactor, capacitor = df['MvarNom'] < 0, df['MvarNom'] > 0
    closed, opened = df['Status'] == 'Closed', df['Status'] == 'Open'
    low, high = df['Vpu'] < vlow, df['Vpu'] > vhigh

    new_status = df['Status'].copy()
    new_status[(reactor & closed & low) | (capacitor & closed & high)] = 'Open'
    new_status[(reactor & opened & high) | (capacitor & opened & low)] = 'Closed'
    return new_status

def adjust_shunts(SimAuto):
    """Switches shunts to attempt to get buses back within a voltage band."""
    def shunt_changes():
        df = pw.get_df(SimAuto, 'Shunt', {'ObjectID': str, 'MvarNom': float, 'Status': str, 'Vpu': float})
        new_status = suggested_shunt_status(df, SHUNT_VLOW, SHUNT_VHIGH)
        changed = df['Status'] != new_status
        # Filter both sides: assigning a full Series to an empty frame makes one blank row per shunt.
        return df[changed].assign(Status=new_status[changed])

    # Adjust all shunts in one go, to get the bulk of the work done.
    changes = shunt_changes()
    if len(changes):
        with pw.trial(SimAuto):
            pw.set_df(SimAuto, 'Shunt', changes, ['ObjectID', 'Status'])

    # Adjust 1 shunt at a time, solve, see if any more need adjustment.
    # This helps fix situations where a bus has several shunts, and only some need to be online.
    failed: list[str] = []
    for _ in range(SHUNT_MAX_ITERATIONS):
        df = shunt_changes()
        df = df[~df['ObjectID'].isin(failed)]
        if len(df) == 0:
            return # Nothing left to do! All shunt voltages within bounds.
        # Adjust the smallest shunts first
        row_df = df.sort_values(by='MvarNom', key=abs).iloc[[0]]
        with pw.trial(SimAuto) as t:
            pw.set_df(SimAuto, 'Shunt', row_df, ['ObjectID', 'Status'])
        if not t.ok:
            failed.append(row_df['ObjectID'].iloc[0])

# ------------------ Fault duty, PVQV, swing ------------------
def get_fault_duty(SimAuto) -> pd.DataFrame:
    """
    Returns a dataframe of fault duty at every bus.
    Note: deletes any existing Fault objects in the case.
    """
    bus_df = pw.get_df(SimAuto, 'Bus', {'ObjectID': str, 'Number': int, 'Name': str, 'NomkV': float})

    fault_df = pd.DataFrame()
    fault_df['FaultName'] = bus_df['Number'].astype(str) + ' : ' + bus_df['Name'].astype(str) + ' : ' + bus_df['NomkV'].astype(str)
    fault_df['WhoAmI'] = bus_df['ObjectID']
    fault_df['CustomInteger:0'] = bus_df['Number']
    fault_df['FaultType'] = '3PB'
    fault_df['FaultImpedance'] = 0
    fault_df['FaultImpedance:1'] = 0

    with pw.creating(SimAuto):
        pw.run(SimAuto, 'Delete(Fault);')
        pw.set_df(SimAuto, 'Fault', fault_df, list(fault_df.columns))

    fault_fields = {
        'FaultName': str
        ,'WhoAmI': str
        ,'CustomInteger:0': int
        ,'FaultType': str
        ,'FaultImpedance': float
        ,'FaultImpedance:1': float
        ,'BusNomVolt': float
        ,'ABCPhaseI': float
    }

    pw.run(SimAuto, 'EnterMode(RUN);')
    pw.run(SimAuto, 'FaultMultiple(NO);')

    result_df = pw.get_df(SimAuto, 'Fault', fault_fields)
    result_df.rename(columns={"CustomInteger:0": "BusNumber"}, inplace=True)
    # MVA = sqrt(3) * kV * kA
    result_df['MVA'] = (3 ** 0.5) * result_df['BusNomVolt'] * result_df['ABCPhaseI']

    # Don't leave thousands of study faults in the saved case.
    pw.run(SimAuto, 'EnterMode(EDIT);')
    pw.run(SimAuto, 'Delete(Fault);')

    return result_df

def get_pvqv(SimAuto) -> pd.DataFrame:
    """
    Calculates bus voltage MW & MVAR self-sensitivity.
    Tools -> Sensitivities -> Flow and Voltage Sensitivities -> Self Sensitivity (Tab)
    """
    bus_fields = {
        'Number': int
        ,'Name': str
        ,'NomkV': float
        ,'AreaNumber': int
        ,'AreaName': str
        ,'SensdVdPself': float
        ,'SensdVdQself': float
        ,'Vpu': float
    }

    pw.run(SimAuto, 'EnterMode(RUN);')
    pw.run(SimAuto, 'CalculateVoltSelfSense();')
    return pw.get_df(SimAuto, 'Bus', bus_fields)

def max_fault_duty_bus(fault_df) -> int:
    return int(fault_df.loc[fault_df['MVA'].idxmax(), 'BusNumber'])

def slack_buses(SimAuto) -> list[int]:
    """The case's swing buses."""
    return pw.get_df(SimAuto, 'Bus', {'Number': int, 'Slack': str}).query("Slack == 'YES'")['Number'].tolist()

def create_giant_swing(SimAuto, bus_num: int) -> pd.DataFrame:
    """
    Creates a giant swing/slack unit on the given bus, e.g. max_fault_duty_bus() or slack_buses()[0].
    """
    # Get current bus information at the swing bus.
    bus_df = pw.get_df(SimAuto, 'Bus', {'Number': int, 'Name': str, 'NomkV': float, 'Vpu': float})
    at_bus = bus_df.loc[bus_df['Number'] == bus_num]
    if at_bus.empty:
        raise ValueError(f'The giant swing bus {bus_num} is not in the case. Set swing_bus to a bus number in it, "slack" or "max_fault_duty".')
    bus_row = at_bus.iloc[0]

    swing_dict = {
        'BusNum': bus_num
        ,'BusName': bus_row['Name']
        ,'NomkV': bus_row['NomkV']
        ,'ID': SWING_ID
        ,'VoltSet': bus_row['Vpu']
        ,'MW': 0.00
        ,'Status': "Closed"
        ,'AVR': "YES"
        ,'AGC': "YES"
        ,'MvarMax': 6573.70
        ,'MvarMin': -6573.70
        ,'Mvar': 0.00
        ,'MWMax': SWING_MW_MAX
        ,'MWMin': SWING_MW_MIN
        ,'FuelTypeCode': "MWH"
        ,'UnitTypeEPC': "42"
        ,'Memo': "Giant Swing for Initial Balance"
    }
    swing_df = pd.DataFrame([swing_dict])

    with pw.creating(SimAuto):
        pw.set_df(SimAuto, 'Gen', swing_df, list(swing_dict))

    # Disable any other swings, and make this bus the system swing.
    pw.run(SimAuto, 'SetData(Bus,[Slack],[NO],ALL);')
    pw.run(SimAuto, f'SetData(Bus,[Number,Slack],[{int(bus_num)}, YES]);')


    return swing_df

def disable_area_agc(SimAuto):
    """Turns off MW AGC in the solution options, so the swing takes up every imbalance while scaling."""
    pw.set_df(SimAuto, 'Sim_Solution_Options_Value', pd.DataFrame({'Option':['ChkMWAGC'], 'Value':['NO']}), ['Option', 'Value'])

# ------------------ Constant-power loads ------------------
LOAD_CONVERT_FIELDS = ['BusNum', 'ID', 'SMW', 'SMvar', 'IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput', 'DistStatus']

def constant_power_rows(load: pd.DataFrame) -> pd.DataFrame:
    """
    Closed loads as constant power: S = what the load draws now (NetMW/NetMvar, after voltage and dist gen), no I, Z or dist gen.
    Input columns: BusNum, ID, ObjectID, AreaNumber, Status, NetMW, NetMvar. The result is indexed by ObjectID.
    A load on a dead bus has NetMW 0 and so becomes 0 MW. Step 02 turns islanded loads Open anyway (when islands are checked).
    """
    out = load[load['Status'] == 'Closed'].set_index('ObjectID')[['BusNum', 'ID', 'AreaNumber', 'NetMW', 'NetMvar']]
    out = out.assign(SMW=out['NetMW'], SMvar=out['NetMvar'])
    out[['IMW', 'IMvar', 'ZMW', 'ZMvar', 'DistMWInput', 'DistMvarInput']] = 0.0
    out['DistStatus'] = 'Open'
    return out.drop(columns=['NetMW', 'NetMvar'])

def _swing_mw(SimAuto) -> float:
    """MW of the Closed gens on the swing buses. (Not area.swing_mw: area imports this module.)"""
    gen = pw.get_df(SimAuto, 'Gen', {'BusNum': int, 'Status': str, 'MW': float})
    return float(gen.loc[(gen['Status'] == 'Closed') & gen['BusNum'].isin(slack_buses(SimAuto)), 'MW'].sum())

def _read_loads(SimAuto) -> pd.DataFrame:
    return pw.get_df(SimAuto, 'Load', {
        'ObjectID': str, 'BusNum': int, 'ID': str, 'AreaNumber': int, 'Status': str, 'MW': float, 'SMW': float, 'SMvar': float,
        'IMW': float, 'IMvar': float, 'ZMW': float, 'ZMvar': float, 'NetMW': float, 'NetMvar': float, 'DistMW': float,
        'DistStatus': str, 'DistMWInput': float, 'DistMvarInput': float})

def _area_net(load: pd.DataFrame) -> pd.DataFrame:
    closed = load[load['Status'] == 'Closed']
    return pd.DataFrame({'Loads': closed.groupby('AreaNumber').size(),
                         'Nominal MW': closed.groupby('AreaNumber')[['SMW', 'IMW', 'ZMW']].sum().sum(axis=1),
                         'Dist MW': closed.groupby('AreaNumber')['DistMW'].sum(),
                         'Net MW': closed.groupby('AreaNumber')['NetMW'].sum(),
                         'Net Mvar': closed.groupby('AreaNumber')['NetMvar'].sum()})

def conversion_warnings(sheet: pd.DataFrame, done: pd.DataFrame) -> list[str]:
    """
    The checks on a conversion: an area's net MW/Mvar (sheet's area rows) or the swing (its 'Swing' row) moved more than
    CONVERT_TOL_MW, or converted loads (`done`, read back after the solve) draw MW other than their SMW or carry dist gen.
    """
    out = []
    areas, swing = sheet[sheet['AreaNumber'] != 'Swing'], sheet[sheet['AreaNumber'] == 'Swing']
    moved = areas[(areas['Diff MW'].abs() > CONVERT_TOL_MW) | (areas['Diff Mvar'].abs() > CONVERT_TOL_MW)]
    if len(moved):
        out.append(f'WARNING: Net load changed by more than {CONVERT_TOL_MW} MW/Mvar in area(s) {", ".join(map(str, moved["AreaNumber"]))}. See load_conversion.')
    d = float(swing['Diff MW'].sum())
    if abs(d) > CONVERT_TOL_MW:
        out.append(f'WARNING: The swing moved {d:.1f} MW in the conversion (more than {CONVERT_TOL_MW} MW). See load_conversion.')
    off = ((done['MW'] - done['SMW']).abs() > CONVERT_TOL_MW) | (done['DistMW'].abs() > 1e-6)
    if off.any():
        out.append(f'WARNING: {int(off.sum())} converted loads still draw MW different from their SMW or carry dist gen '
                   '(PowerWorld may reduce load at low voltage). Check them in the case.')
    return out

def convert_loads_to_constant_power(SimAuto) -> dict[str, pd.DataFrame]:
    """
    Turns every Closed load into constant power at its solved NetMW/NetMvar, in one bisected edit (loads which make the
    case diverge keep their current/impedance/dist gen; step 02 ramps those to 0). Open loads only lose their dist gen.
    Run on the open, solved case with AGC off. Warns, never stops, except when NetMW/NetMvar come back blank.
    Returns the sheets for the log: load_conversion (net MW/Mvar per area and the swing, before and after), and
    load_conversion_failed when there are any.
    """
    if not pw.solve(SimAuto):
        print('WARNING: The case does not solve, so loads were NOT converted to constant power. Step 02 ramps their current, impedance and dist gen to 0.')
        return {}
    before = _read_loads(SimAuto)      # NetMW is read before any write.
    swing_before = _swing_mw(SimAuto)
    closed = before['Status'] == 'Closed'
    blank = before.loc[closed, ['NetMW', 'NetMvar']].isna().any(axis=1).sum()
    if blank:
        # A wrong field name might come back as blanks rather than an error; writing them would zero every load.
        raise SystemExit(f'{blank} Closed loads have a blank NetMW or NetMvar. Check those Load field names in PowerWorld.')

    # Open loads inject nothing: clear their dist gen so it can't come back on, and leave their S, I and Z alone.
    dist_on = (before['Status'] == 'Open') & ((before['DistStatus'] != 'Open') | (before['DistMWInput'] != 0) | (before['DistMvarInput'] != 0))
    open_rows = before[dist_on].set_index('ObjectID')[['BusNum', 'ID']].assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)
    if len(open_rows):
        with pw.trial(SimAuto) as t:
            pw.set_df(SimAuto, 'Load', open_rows, ['BusNum', 'ID', 'DistStatus', 'DistMWInput', 'DistMvarInput'])
        if not t.ok:
            print(f'WARNING: Did not solve after clearing dist gen on {len(open_rows)} Open loads. Rolled back.')

    rows = constant_power_rows(before)
    failed = pw.set_df_bisect(SimAuto, 'Load', rows, LOAD_CONVERT_FIELDS, limit=CONVERT_MAX_FAILED)
    if not pw.solve(SimAuto):       # The last bisect trial may have ended on a rolled-back state.
        print('WARNING: The case does not solve after the conversion: the load_conversion figures are unsolved values.')
    after = _read_loads(SimAuto)
    swing_after = _swing_mw(SimAuto)

    # Sheet: one row per area, then the swing.
    b, a = _area_net(before), _area_net(after)
    sheet = pd.DataFrame({'Loads': b['Loads'], 'Nominal MW': b['Nominal MW'], 'Dist MW': b['Dist MW'],
                          'Net MW Before': b['Net MW'], 'Net MW After': a['Net MW'].reindex(b.index).fillna(0.0),
                          'Net Mvar Before': b['Net Mvar'], 'Net Mvar After': a['Net Mvar'].reindex(b.index).fillna(0.0)})
    sheet.insert(5, 'Diff MW', sheet['Net MW After'] - sheet['Net MW Before'])
    sheet['Diff Mvar'] = sheet['Net Mvar After'] - sheet['Net Mvar Before']
    sheet = sheet.rename_axis('AreaNumber').reset_index().astype({'AreaNumber': object})
    swing = pd.DataFrame([{'AreaNumber': 'Swing', 'Net MW Before': swing_before, 'Net MW After': swing_after, 'Diff MW': swing_after - swing_before}])
    sheet = pd.concat([sheet, swing], ignore_index=True)

    n_closed = int(closed.sum())
    print(f'Loads to constant power: {n_closed} loads, net {before.loc[closed, "NetMW"].sum():.1f} MW '
          f'(dist gen {before.loc[closed, "DistMW"].sum():.1f} MW folded in)')
    done = after[after['ObjectID'].isin(rows.index.difference(failed)) & (after['Status'] == 'Closed')]
    for w in conversion_warnings(sheet, done):
        print(w)
    if failed:
        capped = f' (stopped bisecting after {CONVERT_MAX_FAILED} failures; the rest were not tried)' if len(failed) >= CONVERT_MAX_FAILED else ''
        print(f'WARNING: {len(failed)} loads could not be converted without diverging{capped}. They keep current/impedance/dist gen; step 02 ramps those to 0.')

    out = {'load_conversion': sheet}
    if failed:
        out['load_conversion_failed'] = before[before['ObjectID'].isin(failed)].set_index('ObjectID')[
            ['BusNum', 'ID', 'AreaNumber', 'SMW', 'IMW', 'ZMW', 'DistMW', 'NetMW', 'NetMvar']]
    return out
