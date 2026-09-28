"""
Step 01: make the PWB seed case match the GridView topology, and prepare it for scaling.
"""
from pathlib import Path
import pandas as pd
from . import pw

# ------------------ Tuning thresholds ------------------
TAP_THRESHOLD = 0.15            # Transformer taps / nominal kV further than this from 1.0 PU are reset to 1.0.
SHUNT_VLOW = 0.92               # Switch shunts when bus Vpu drops below this...
SHUNT_VHIGH = 1.08              # ...or rises above this.
SHUNT_MAX_ITERATIONS = 10       # Max one-at-a-time shunt switching passes.

# Branch types which GridView may drop from EPC exports. Excluded from branch status matching.
# Check these match BranchDeviceType values in your PowerWorld version.
UNSUPPORTED_GV_BRANCH_TYPES = ['Breaker', 'Disconnect', 'Fuse', 'Ground Disconnect', 'Load Break Disconnect', 'ZBR']

SWING_ID = 'xx'     # ID of the giant swing generator.

# ------------------ Case data ------------------
# All the required network-parameters to fully define each object, for copying missing elements.
def _owners(n: int) -> dict[str,type]:
    return {k: t for i in range(1, n + 1) for k, t in ((f'OwnerNum{i}', int), (f'OwnerPerc{i}', float))}

_LIMITS_MVA = {f'LimitMVA{c}': float for c in 'ABCDEFGHIJKLMNO'}

_EMS_BRANCH = {
    'EMSType': str
    ,'EMSID': str
    ,'EMSLineID': str
    ,'EMSCBTyp': str
    ,'EMSID2From': str
    ,'EMSID2To': str
    ,'DataMaintainerAssign': str
    ,'DataMaintainerInherit': str
    ,'AllLabels': str
}

BUS_PARAMS: dict[str,type] = {
    'ObjectID': str
    ,'Number': int
    ,'Name': str
    ,'NomkV': float
    ,'Slack': str
    ,'NomB': float
    ,'NomG': float
    ,'Vpu': float
    ,'Vangle': float
    ,'DCLossMultiplier': float
    ,'AreaNumber': int
    ,'ZoneNumber': int
    ,'BANumber': int
    ,'OwnerNumber': int
    ,'SubNumber': int
    ,'Monitor': str
    ,'LimitSet': str
    ,'UseSpecificLimits': str
    ,**{f'LimitLow{c}': float for c in 'ABCD'}
    ,**{f'LimitHigh{c}': float for c in 'ABCD'}
    ,'Latitude': float
    ,'Longitude': float
    ,'TopologyBusType': str
    ,'Priority': float
    ,'EMSType': str
    ,'EMSID': str
    ,'DataMaintainerAssign': str
    ,'DataMaintainerInherit': str
    ,'DataMaintainerInheritBlock': str
    ,'AllLabels': str
}

LOAD_PARAMS: dict[str,type] = {
    'ObjectID': str
    ,'BusNum': int
    ,'ID': str
    ,'Status': str
    ,'AGC': str
    ,'SMW': float
    ,'SMvar': float
    ,'IMW': float
    ,'IMvar': float
    ,'ZMW': float
    ,'ZMvar': float
    ,'DistStatus': str
    ,'DistMWInput': float
    ,'DistMvarInput': float
    ,'Interruptible': str
    ,'MWMax': float
    ,'MWMin': float
    ,'DistMWMax': float
    ,'DistMWMin': float
    ,'DistUnitTypeCode': str
    ,'LoadModelGroup': str
    ,'AreaNumber': int
    ,'ZoneNumber': int
    ,'BANumber': int
    ,'OwnerNumber': int
    ,'EMSType': str
    ,'EMSID': str
    ,'DataMaintainerAssign': str
    ,'DataMaintainerInherit': str
    ,'AllLabels': str
}

GEN_PARAMS: dict[str,type] = {
    # Note: ReactiveCapability curve data is not necessary in the comparisons.
    # The Anchor Data Set may have additional generators, but none of them use ReactiveCapability curves.
    'ObjectID': str
    ,'BusNum': int
    ,'ID': str
    ,'Status': str
    ,'VoltSet': float
    ,'VoltSetTol': float
    ,'RegBusNum': int
    ,'RegFactor': float
    ,'AGC': str
    ,'PartFact': float
    ,'MWSetPoint': float
    ,'MWMax': float
    ,'MWMin': float
    ,'EnforceMWLimit': str
    ,'AVR': str
    ,'MvarSetPoint': float
    ,'MvarMax': float
    ,'MvarMin': float
    ,'UseCapCurve': str
    ,'WindContMode': str
    ,'WindContModePF': float
    ,'UseLineDrop': str
    ,'Rcomp': float
    ,'Xcomp': float
    ,'VoltageDroopControl': str
    ,'MVABase': float
    ,'GenR': float
    ,'GenX': float
    ,'StepR': float
    ,'StepX': float
    ,'StepTap': float
    ,'GovRespLimit': str
    ,'UnitTypeCode': str
    ,'FuelTypeCode': str
    ,'AreaNumber': int
    ,'ZoneNumber': int
    ,'BANumber': int
    ,**_owners(8)
    ,'EMSType': str
    ,'EMSID': str
    ,'DataMaintainerAssign': str
    ,'DataMaintainerInherit': str
    ,'AllLabels': str
}

BRANCH_PARAMS: dict[str,type] = {
    # Non-transformer branches.
    'ObjectID': str
    ,'BusNumFrom': int
    ,'BusNumTo': int
    ,'Circuit': str
    ,'BranchDeviceType': str
    ,'ConsolidateAllow': str
    ,'OpenOrCloseBreakersAllow': str
    ,'Status': str
    ,'StatusNormal': str
    ,'ByPass': str
    ,'MeteredBus': str
    ,'R': float
    ,'X': float
    ,'B': float
    ,'G': float
    ,'LineLength': float
    ,'Monitor': str
    ,'LimitSet': str
    ,**_LIMITS_MVA
    ,**_owners(8)
    ,**_EMS_BRANCH
}

TRANSFORMER_PARAMS: dict[str,type] = {
    # Transformers in the branch table.
    'ObjectID': str
    ,'BusNumFrom': int
    ,'BusNumTo': int
    ,'Circuit': str
    ,'BranchDeviceType': str
    ,'Status': str
    ,'StatusNormal': str
    ,'ByPass': str
    ,'MeteredBus': str
    ,'ControlType': str
    ,'AutoControl': str
    ,'RegBusNum': int
    ,'UseLineDrop': str
    ,'Rcomp': float
    ,'Xcomp': float
    ,'RegMax': float
    ,'RegMin': float
    ,'RegTargetType': str
    ,'XFMVABase': float
    ,'XFNomkVbaseFrom': float
    ,'XFNomkVbaseTo': float
    ,'Rxfbase': float
    ,'Xxfbase': float
    ,'Gxfbase': float
    ,'Bxfbase': float
    ,'Gmagxfbase': float
    ,'Bmagxfbase': float
    ,'TapFixedFrom': float
    ,'TapFixedTo': float
    ,'TapMaxxfbase': float
    ,'TapMinxfbase': float
    ,'TapStepSizexfbase': float
    ,'Tapxfbase': float
    ,'Phase': float
    ,'ImpCorrTable': float
    ,'LineLength': float
    ,'Monitor': str
    ,'LimitSet': str
    ,**_LIMITS_MVA
    ,**_owners(8)
    ,**_EMS_BRANCH
}

LINESHUNT_PARAMS: dict[str,type] = {
    'ObjectID': str
    ,'BusNumFrom': int
    ,'BusNumTo': int
    ,'Circuit': str
    ,'ID': str
    ,'BusNumLoc': int
    ,'Status': str
    ,'MWNom': float
    ,'MvarNom': float
    ,**_owners(4)
    ,'DataMaintainerAssign': str
    ,'DataMaintainerInherit': str
}


def get_case_data(SimAuto) -> dict[str,dict[str,object]]:
    print('get_case_data()')
    return {
        'Bus':{
            'table_name': 'Bus'
            ,'df': pw.get_df(SimAuto, 'Bus', BUS_PARAMS)
        }
        ,'Load':{
            'table_name': 'Load'
            ,'df': pw.get_df(SimAuto, 'Load', LOAD_PARAMS)
        }
        ,'Gen':{
            'table_name': 'Gen'
            ,'df': pw.get_df(SimAuto, 'Gen', GEN_PARAMS)
        }
        ,'Branch':{
            'table_name': 'Branch'
            ,'df': pw.get_df(SimAuto, 'Branch', BRANCH_PARAMS, "BranchDeviceType notcontains 'Transformer'")
        }
        ,'Transformer':{
            'table_name': 'Branch'
            ,'df': pw.get_df(SimAuto, 'Branch', TRANSFORMER_PARAMS, "BranchDeviceType = 'Transformer'")
        }
        ,'LineShunt':{
            'table_name': 'LineShunt'
            ,'df': pw.get_df(SimAuto, 'LineShunt', LINESHUNT_PARAMS)
        }
    }

# ------------------ Dummy buses ------------------
def create_dummy_bus_aux(SimAuto, dummy_bus_fp: Path):
    """Saves the open case's multi-section line dummy bus numbers, as an aux file which renumbers another case to match."""
    pw.run(SimAuto, 'SaveData("'+str(dummy_bus_fp)+'", AUX, MultiSectionLine, [BusNumFrom,BusNameFrom,BusNumTo,BusNameTo,Circuit], [Bus], , [], NO, NO);')

    content = dummy_bus_fp.read_text(encoding="utf-8")
    content = content.replace("&", " ")
    content = content.replace("<SUBDATA Bus>", "<SUBDATA BusRenumber>")
    dummy_bus_fp.write_text(content, encoding="utf-8")

def open_gv_case(SimAuto, fp, dummy_bus_fp):
    """Opens a GridView EPC, and renumbers its multi-section line dummy buses to match the PWB seed."""
    pw.open_case(SimAuto, fp)
    pw.run(SimAuto, 'EnterMode(EDIT);')
    result = SimAuto.RunScriptCommand(f'LoadAux("{dummy_bus_fp}",YES);')
    # Printed rather than raised, as before: lines which aren't in this EPC may report errors.
    if result[0] != '':
        print(f'LoadAux DummyBus messages: {result[0]}')

# ------------------ Missing elements ------------------
def missing_rows(left_df: pd.DataFrame, right_df: pd.DataFrame) -> pd.DataFrame:
    """Rows of left_df whose ObjectID is not in right_df."""
    return left_df[~left_df['ObjectID'].isin(right_df['ObjectID'])].copy()

def create_missing_elements(SimAuto, left_case_dict, right_case_dict) -> dict[str,pd.DataFrame]:
    # Any elements on the left, which aren't on the right, will get created.
    print('create_missing_elements()')

    missing_dict: dict[str,pd.DataFrame] = {}

    with pw.creating(SimAuto):
        for element_type, left in left_case_dict.items():
            print(f'Creating missing {element_type} elements')
            missing_df = missing_rows(left['df'], right_case_dict[element_type]['df'])
            missing_dict[element_type] = missing_df

            # Insert the elements into the case with an initially open status.
            create_df = missing_df.copy()
            if 'Status' in create_df.columns:
                create_df['Status'] = 'Open'

            # Several GridView transformers are in automatic control with crazy tap settings. Disable automatic controls.
            if element_type == 'Transformer':
                create_df['AutoControl'] = 'NO'

            # To create objects, you cannot have 'ObjectID' in the list of fields.
            fields = [c for c in create_df.columns if c != 'ObjectID']
            pw.set_df(SimAuto, left['table_name'], create_df, fields)

    return missing_dict

def get_gv_loads(SimAuto, gv_fps: list[Path], dummy_bus_fp: Path) -> pd.DataFrame:
    """
    Gathers load data from all GridView EPCs, since X1/X2/X3 etc distributed generation loads
    are dynamically generated by GridView for each hour.
    """
    gv_load_df_list = []
    for gv_fp in gv_fps:
        open_gv_case(SimAuto, gv_fp, dummy_bus_fp)
        gv_load_df_list.append(pw.get_df(SimAuto, 'Load', LOAD_PARAMS))
        SimAuto.CloseCase()
    return pd.concat(gv_load_df_list, ignore_index=True).drop_duplicates(subset='ObjectID', keep='first')

def create_distgen_XN_loads(SimAuto, gv_load_df: pd.DataFrame) -> pd.DataFrame:
    """
    With the TopoSeed case open, creates all loads from gv_load_df which don't exist already.
    New loads will be in a normal-open status, with MW=0 MVAR=0 for all related values.
    Returns a dataframe of all distributed generation loads which were created.
    """
    pw_load_df = pw.get_df(SimAuto, 'Load', {'ObjectID': str})
    missing_df = missing_rows(gv_load_df, pw_load_df)
    missing_df[['Status', 'DistStatus']] = 'Open'
    missing_df[['SMW', 'SMvar', 'DistMWInput', 'DistMvarInput', 'IMW', 'IMvar', 'ZMW', 'ZMvar']] = 0

    # To create objects, you cannot have 'ObjectID' in the list of fields.
    with pw.creating(SimAuto):
        pw.set_df(SimAuto, 'Load', missing_df, [c for c in missing_df.columns if c != 'ObjectID'])

    return missing_df

# ------------------ Transformer taps ------------------
def transformer_is_okay(df: pd.DataFrame, threshold: float = TAP_THRESHOLD) -> pd.Series:
    def near_one(s):
        return (s - 1.0).abs() < threshold
    return (near_one(df['Tap'])
            & near_one(df['NomkVFrom'] / df['XFNomkVbaseFrom']) & near_one(df['TapFixedFrom'])
            & near_one(df['NomkVTo'] / df['XFNomkVbaseTo']) & near_one(df['TapFixedTo']))

def fix_transformer_taps(SimAuto, threshold = TAP_THRESHOLD):
    # In GridView, some transformer models may have transformer nominal kV and tap positions which are aggregious errors.
    # GridView doesn't have issues with this (DC Loadflow), but this certainly will cause problems in PowerWorld.
    # If Nominal kV and Tap positions are far out of range, fix them to 1.0 PU based on the bus nominal kV.
    fields: dict[str,type] = {
        'ObjectID': str
        ,'ControlType': str
        ,'NomkVFrom': float
        ,'XFNomkVbaseFrom': float
        ,'NomkVTo': float
        ,'XFNomkVbaseTo': float
        ,'TapFixedFrom': float
        ,'TapFixedTo': float
        ,'Tap': float
    }
    df = pw.get_df(SimAuto, 'Branch', fields, "BranchDeviceType = 'Transformer'")
    bad_df = df[~transformer_is_okay(df, threshold)]

    fix_df = bad_df.assign(
        XFNomkVbaseFrom=bad_df['NomkVFrom']
        ,XFNomkVbaseTo=bad_df['NomkVTo']
        ,TapFixedFrom=1.0
        ,TapFixedTo=1.0
        ,Tap=1.0
    )
    pw.set_df(SimAuto, 'Branch', fix_df, ['ObjectID', 'XFNomkVbaseFrom', 'XFNomkVbaseTo', 'TapFixedFrom', 'TapFixedTo', 'Tap'])

    return bad_df

# ------------------ Branch statuses ------------------
def branch_status_changes(left_case_dict, right_case_dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns:
        status_targets_df: branches whose status differs between the models (for reporting).
        changes_df: ['ObjectID', 'Status', 'StatusLeft'] for right-model branches which need a status change.
    """
    def branches(case_dict):
        # Get Transformer and Non-Transformer branch statuses by ObjectID.
        df = pd.concat([case_dict[k]['df'][['ObjectID','Status','BranchDeviceType']] for k in ('Branch', 'Transformer')], ignore_index=True)
        # GridView may drop some branch types from its EPC exports. Exclude them from the status changes.
        return df[~df['BranchDeviceType'].isin(UNSUPPORTED_GV_BRANCH_TYPES)].drop(columns=['BranchDeviceType'])

    # right df, with ['ObjectID', 'Status', 'StatusLeft']
    merged_df = branches(right_case_dict).merge(branches(left_case_dict), on='ObjectID', how='left', suffixes=('', 'Left'))

    # Save a copy of the status targets for reporting purposes.
    status_targets_df = merged_df[merged_df['Status'] != merged_df['StatusLeft']].sort_values(by='StatusLeft')

    # Any branches that don't exist on the left will need to be opened.
    merged_df['StatusLeft'] = merged_df['StatusLeft'].fillna('Open')
    changes_df = merged_df[merged_df['Status'] != merged_df['StatusLeft']].copy()

    return status_targets_df, changes_df

def set_branch_statuses(SimAuto, left_case_dict, right_case_dict):
    """
    With the "Right" case open, sets the "Right" model branch statuses to match those from the "Left" model.
    Returns [status_targets_df, fail_df].
    """
    status_targets_df, changes_df = branch_status_changes(left_case_dict, right_case_dict)
    failed = pw.set_df_bisect(SimAuto, 'Branch', changes_df.assign(Status=changes_df['StatusLeft']), ['ObjectID', 'Status'])
    return [status_targets_df, changes_df.loc[failed]]

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

def adjust_shunts(SimAuto, vlow: float = SHUNT_VLOW, vhigh: float = SHUNT_VHIGH, max_iterations: int = SHUNT_MAX_ITERATIONS):
    """Switches shunts to attempt to get buses back within a voltage band."""
    def shunt_changes():
        df = pw.get_df(SimAuto, 'Shunt', {'ObjectID': str, 'MvarNom': float, 'Status': str, 'Vpu': float})
        new_status = suggested_shunt_status(df, vlow, vhigh)
        return df[df['Status'] != new_status].assign(Status=new_status)

    # Adjust all shunts in one go, to get the bulk of the work done.
    with pw.trial(SimAuto):
        pw.set_df(SimAuto, 'Shunt', shunt_changes(), ['ObjectID', 'Status'])

    # Adjust 1 shunt at a time, solve, see if any more need adjustment.
    # This helps fix situations where a bus has several shunts, and only some need to be online.
    failed: list[str] = []
    for _ in range(max_iterations):
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

def slack_bus(SimAuto) -> int:
    """The case's existing swing bus (the first, if there are several)."""
    bus_df = pw.get_df(SimAuto, 'Bus', {'Number': int, 'Slack': str})
    return int(bus_df.loc[bus_df['Slack'] == 'YES', 'Number'].iloc[0])

def create_giant_swing(SimAuto, bus_num: int) -> pd.DataFrame:
    """
    Creates a giant swing/slack unit on the given bus, e.g. max_fault_duty_bus() or slack_bus().
    """
    # Get current bus information at the swing bus.
    bus_df = pw.get_df(SimAuto, 'Bus', {'Number': int, 'Name': str, 'NomkV': float, 'Vpu': float})
    bus_row = bus_df.loc[bus_df['Number'] == bus_num].iloc[0]

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
        ,'MWMax': 20000.00
        ,'MWMin': -20000.00
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

    # Turn off MW AGC
    pw.set_df(SimAuto, 'Sim_Solution_Options_Value', pd.DataFrame({'Option':['ChkMWAGC'], 'Value':['NO']}), ['Option', 'Value'])

    return swing_df
