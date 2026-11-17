"""
Step 01, both modes: seed preparation which doesn't change topology. Switches shunts for voltage, fault duty and
PVQV sensitivities, the giant swing, and turning off area AGC. (Terminal voltage control is
Scripts/GenTerminalVoltageControl.aux, loaded by step 01.) Topology matching is in topo.py.
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
