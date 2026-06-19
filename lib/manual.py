"""MANUAL"""

import pandas as pd
import numpy as np

def find_duplicated_master_device(master_device_type, df):
    """Returns a list of full location movex id containing the locations that has more than one master device, especified in 'master_device_type' variable"""
    df = df.sort_values('Movex Account Number')
    full_location_movex_id_with_two_master_devices = df.groupby('Full Location Movex Id').filter(
        lambda x: ((x['Sibex Name'].str.upper() == master_device_type) & (x['Status'] == 'Installed')).sum() == 2
    )
    return full_location_movex_id_with_two_master_devices['Full Location Movex Id'].unique().tolist()

def find_missing_master_device(master_device_type, df):
    """Returns a list of full location movex id containing the locations that has not a master device"""
    df = df.sort_values('Movex Account Number')
    full_location_movex_id_missing_md = df.groupby('Full Location Movex Id').filter(
        lambda x: (x['Sibex Name'].str.upper() == master_device_type).sum() == 0
    )
    return full_location_movex_id_missing_md['Full Location Movex Id'].unique().tolist()

def generate_dosisoft_comparison_report(all_data, sales_report, hierarchy_json_path, output_filename='DOSIsoft_Comparison_Report.xlsx'):
    """
    Processes CLM and Sales Report data for DOSIsoft, resolves mismatched column names,
    and generates a consolidated variance report saved as an Excel file.
    """
    import lib.mass_update as mu  # Imported here in case it's in a separate module
    
    # -------------------------------------------------------------------------
    # 1. CLM DATA PROCESSING (all_data)
    # -------------------------------------------------------------------------
    dosisoft_clm = mu.filter_df_by_hierarchy_json(all_data, hierarchy_json_path)
    dosisoft_clm_filtered = dosisoft_clm[dosisoft_clm['Sibex Name'] != 'DOSISOFT']
    
    group_cols = ['Movex Account Number', 'Account Name', 'Account Country', 'Account Region', 'Sibex Name']
    dosisoft_clm_grouped = dosisoft_clm_filtered.groupby(group_cols)['Quantity'].sum().reset_index()
    
    df_pivot = dosisoft_clm_grouped.pivot_table(
        index=['Movex Account Number', 'Account Name', 'Account Country', 'Account Region'], 
        columns='Sibex Name', 
        values='Quantity',
        aggfunc='sum',
        fill_value=0
    ).astype(int).reset_index()
    
    df_pivot.columns.name = None
    
    cols_thinkqa = df_pivot.filter(like='THINKQA').columns
    df_pivot['ThinkQA'] = df_pivot[cols_thinkqa].sum(axis=1)
    
    cols_hw = df_pivot.filter(regex='(?i)DOSISOFT HW').columns
    df_pivot['DOSISOFT HW'] = df_pivot[cols_hw].sum(axis=1)
    
    cols_to_drop = [c for c in list(set(cols_thinkqa.tolist() + cols_hw.tolist())) 
                    if c not in ['ThinkQA', 'DOSISOFT HW']]
    df1 = df_pivot.drop(columns=cols_to_drop)
    
    # -------------------------------------------------------------------------
    # 2. SALES REPORT PROCESSING & DYNAMIC RENAMING
    # -------------------------------------------------------------------------
    df2 = sales_report.copy()
    
    # Clean up column headers by converting to clean strings
    df2.columns = [str(c).replace('\n', ' ').strip() for c in df2.columns]
    
    # SMART SEARCH FOR ACCOUNT ID COLUMN
    id_col = 'Movex Account Number'
    matched_id_cols = []
    
    # Attempt 1: Look for 'movex account number'
    matched_id_cols = [c for c in df2.columns if 'movex account number' in c.lower()]
    
    # Attempt 2: If not found, look for any column that has BOTH 'movex' and 'account'
    if not matched_id_cols:
        matched_id_cols = [c for c in df2.columns if 'movex' in c.lower() and 'account' in c.lower()]
        
    # Attempt 3: If still not found, look for just 'movex'
    if not matched_id_cols:
        matched_id_cols = [c for c in df2.columns if 'movex' in c.lower()]

    # If a column is found, rename it. If not, raise an informative error with available columns.
    if matched_id_cols:
        df2 = df2.rename(columns={matched_id_cols[0]: id_col})
    else:
        raise ValueError(
            f"\n[ERROR] Could not automatically find the Account Number column in Sales Report.\n"
            f"Available columns in your file are:\n{list(df2.columns)}"
        )
    
    rename_mapping = {
        'Hardware': 'DOSISOFT HW',
        'MU2net': 'DOSISOFT SW MU2NET',
        'ThinkQA2 SDC': 'ThinkQA',
        'EPIbeam': 'DOSISOFT SW EPIB',
        'EPIgray': 'DOSISOFT SW EPIG'
    }
    df2 = df2.rename(columns=rename_mapping)
    
    info_cols = ['Account Name', 'Account Country', 'Account Region']
    products = ['DOSISOFT HW', 'DOSISOFT SW MU2NET', 'ThinkQA', 'DOSISOFT SW EPIB', 'DOSISOFT SW EPIG']
    
    for prod in products:
        if prod not in df2.columns:
            df2[prod] = 0
            
    df2_grouped = df2.groupby(id_col)[products].sum()
    df2_grouped.index = df2_grouped.index.fillna(0).astype(int)
    df2_grouped = df2_grouped.fillna(0).astype(int)
    
    df2_final = df2_grouped.reset_index()
    df2_final.columns.name = None
    
    # -------------------------------------------------------------------------
    # 3. COMPARISON AND DATA MERGING (OUTER JOIN)
    # -------------------------------------------------------------------------
    ids_sf = df1[id_col].unique()
    ids_sr = df2_final[id_col].unique()
    
    comparison = pd.merge(df1, df2_final, on=id_col, how='outer', suffixes=('_sf', '_sr')).fillna(0)
    
    final_df = pd.DataFrame()
    final_df[id_col] = comparison[id_col].astype(int)
    
    for col in info_cols:
        final_df[col] = comparison[col]
        
    for prod in products:
        sf = comparison[f'{prod}_sf'].astype(int)
        sr = comparison[f'{prod}_sr'].astype(int)
        
        final_df[prod] = np.where(
            sf == sr, 
            'Match', 
            'Has ' + sf.astype(str) + ' but should have ' + sr.astype(str)
        )
        
    def get_overall_status(row):
        account = row[id_col]
        row_values = row[products].values
        
        if account in ids_sf and account not in ids_sr:
            return 'Only_in_sf'
        if account in ids_sr and account not in ids_sf:
            return 'only_sr'
        if all(v == 'Match' for v in row_values):
            return 'Match'
        return 'Divergent_Products'
        
    final_df['Overall Status'] = final_df.apply(get_overall_status, axis=1)
    final_df.columns.name = None
    
    final_df.to_excel(output_filename, index=False)
    print(f"Successfully generated: {output_filename}")
    
    return final_df
