"""
Topology mode: step 01 makes the PWB seed case match the GridView topology (missing elements, transformer taps,
branch statuses, distributed generation loads). Step 02 targets each gen & load to its own EPC value (build_targets()).
The preparation both modes share is in prep.py.
"""
from pathlib import Path
import pandas as pd
from . import area, prep, pw, scale

# ------------------ Tuning thresholds ------------------
TAP_THRESHOLD = 0.15            # Transformer taps / nominal kV further than this from 1.0 PU are reset to 1.0.

# Branch types which GridView may drop from EPC exports. Excluded from branch status matching.
# Check these match BranchDeviceType values in your PowerWorld version.
UNSUPPORTED_GV_BRANCH_TYPES = ['Breaker', 'Disconnect', 'Fuse', 'Ground Disconnect', 'Load Break Disconnect', 'ZBR']

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
def transformer_is_okay(df: pd.DataFrame) -> pd.Series:
    def near_one(s):
        return (s - 1.0).abs() < TAP_THRESHOLD
    return (near_one(df['Tap'])
            & near_one(df['NomkVFrom'] / df['XFNomkVbaseFrom']) & near_one(df['TapFixedFrom'])
            & near_one(df['NomkVTo'] / df['XFNomkVbaseTo']) & near_one(df['TapFixedTo']))

def fix_transformer_taps(SimAuto):
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
    bad_df = df[~transformer_is_okay(df)]

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

# ------------------ Targets (step 02) ------------------
TARGET_FIELDS = {
    table: {'ObjectID': str, 'BusNum': int, 'BusName': str, 'NomkV': float, 'ID': str,
            **{s: str for s in scale.STATUS[table]}, **{c: float for c in scale.SCALED[table]},
            # Gens: the swing balances raise them up to MWMax, spread by area & GenType like area mode.
            **({'MW': float, 'MWMin': float, 'MWMax': float, 'AreaNumber': int, 'UnitTypeCode': str, 'FuelTypeCode': str}
               if table == 'Gen' else {})}
    for table in scale.STATUS
}

def compute_pw_targets(SimAuto, gv_fp: Path, pw_fp: Path, dummy_bus_fp: Path) -> tuple[dict[str, pd.DataFrame], dict]:
    """
    Returns (targets {'Gen': df, 'Load': df}, reports {'swing': area.swing_epc() of the main island's gens}), taking the
    GridView EPC's gens & loads into the PWB case. See build_targets().
    """
    open_gv_case(SimAuto, gv_fp, dummy_bus_fp)
    gv = {table: pw.get_df(SimAuto, table, fields) for table, fields in TARGET_FIELDS.items()}

    pw.open_case(SimAuto, pw_fp)
    case = {table: pw.get_df(SimAuto, table, fields) for table, fields in TARGET_FIELDS.items()}
    slack = prep.slack_buses(SimAuto)
    islands = scale.seed_islands(SimAuto, slack)
    SimAuto.CloseCase()

    targets = build_targets(gv, case, slack)
    scale.exclude_islands(targets, islands)     # Gens/loads outside the main island are turned off before scaling.
    main_gen = case['Gen'][scale.island_reason(case['Gen'], islands) == '']    # A second island's slack has no goal.
    return targets, {'swing': area.swing_epc(gv['Gen'], main_gen, slack)}

def build_targets(gv: dict[str,pd.DataFrame], case: dict[str,pd.DataFrame], swing_buses=()) -> dict[str,pd.DataFrame]:
    """
    Puts the case's gens & loads side by side with their GridView targets (each element to its own EPC value).
    The swing is never scaled: the giant swing, and any gen on swing_buses (the slack bus). The solution sets their MW.
    """
    targets = {}
    for table, statuses in scale.STATUS.items():
        gv_df = gv[table]
        if table == 'Load':
            # GridView leaves distributed generation in the load table at seed-case values, unused. Target Open & 0 MW.
            gv_df = gv_df.assign(DistStatus='Open', DistMWInput=0.0, DistMvarInput=0.0)

        fields = [*statuses, *scale.SCALED[table]]
        df = case[table].merge(gv_df[['ObjectID', *fields]], on='ObjectID', how='left', suffixes=('', '_Target'))
        for f in fields:
            df[f'{f}_Seed'] = df[f]

        # Elements which aren't in the GridView hour are opened.
        df = df.fillna({**{f'{s}_Target': 'Open' for s in statuses}, **{f'{c}_Target': 0.0 for c in scale.SCALED[table]}})

        # Anything open, now or in the target, is at 0 MW / 0 Mvar.
        for suffix in ('', '_Target'):
            for status, zeroed in statuses.items():
                df.loc[df[status + suffix] == 'Open', [c + suffix for c in zeroed]] = 0.0

        df['ExclusionReason'] = ''
        targets[table] = df.set_index('ObjectID')

    gen_df = targets['Gen']
    swing = (gen_df['ID'] == prep.SWING_ID) | gen_df['BusNum'].isin(swing_buses)
    gen_df.loc[swing, 'Status_Target'] = 'Closed'
    scale.exclude(gen_df, swing, 'Swing Unit')
    return targets
