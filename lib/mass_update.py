"""MASS UPDATE"""

import pandas as pd
import json

# Restricts external access to only expose functions starting with 'clean'
__all__ = [
    'clean_quantity',
    'clean_product_hierarchy',
    'clean_status'
]

# --- PUBLIC FUNCTIONS ---

def clean_quantity(df, result_name='final_quantity_result.xlsx'):
    """
    Filters rows with negative or missing quantities, standardizes the values,
    exports the result directly to an Excel spreadsheet, and returns the formatted DataFrame.
    """
    # Isolate rows where Quantity is negative or missing
    result = df[(df['Quantity'] < 0) | (df['Quantity'].isna())].copy()
    
    # Standardize missing quantities to 1
    result['Quantity'] = result['Quantity'].fillna(1)
    
    # Convert negative quantities to positive
    result.loc[result['Quantity'] < 0, 'Quantity'] *= -1
    
    # Define exact columns to keep for the final output format
    columns_to_keep = ['Installed Product ID', 'Quantity', 'Sibex Name', 'Full Location Movex Id']
    result = result[columns_to_keep].copy()
    
    # Export the final spreadsheet directly to Excel
    result.to_excel(result_name, index=False)
    
    return result

def clean_product_hierarchy(df, product_hierarchy, result_name, ib_tracked_df=None):
    """
    Cleans the installed product hierarchy columns based on a filtered product hierarchy
    and exports only the finalized correction spreadsheet directly to Excel.
    """
    # Filter the hierarchy dictionary using valid codes from ib_tracked_df if provided
    if ib_tracked_df is not None:
        allowed_codes = set(ib_tracked_df['Product Code'].astype(str).str.strip())
        product_hierarchy = filter_hierarchy_by_allowed_codes(product_hierarchy, allowed_codes)
        if not product_hierarchy:
            product_hierarchy = {}
        
    # Find IDs with duplicated fathers and remove them from processing
    duplicated_father_ids = find_locations_with_duplicated_fathers(df, product_hierarchy)
    IPs_with_no_duplicated_father_filtered = filter_df(df, duplicated_father_ids)
    
    # Run the cleanup and population chain based on the filtered JSON structure
    IPs_Included_in_setup_cleaned = clear_parent_master_columns(IPs_with_no_duplicated_father_filtered)
    IPs_Included_in_setup_cleaned = populate_master_device_by_hierarchy(IPs_Included_in_setup_cleaned, product_hierarchy)
    IPs_Included_in_setup_cleaned = populate_parent_by_hierarchy_dict(IPs_Included_in_setup_cleaned, product_hierarchy)

    columns_to_keep = ['Installed Product ID', 'Parent: Installed Product', 'Master Device: Installed Product', 
                       'Sibex Name', 'Product: Product Code', 'Full Location Movex Id']
    IPs_Included_in_setup_cleaned = IPs_Included_in_setup_cleaned[columns_to_keep].copy()

    # Map current base, replacing old values with the correct IDs
    mapping_dict = dict(zip(
        IPs_with_no_duplicated_father_filtered['Installed Product'].astype(str), 
        IPs_with_no_duplicated_father_filtered['Installed Product ID']
    ))

    columns_to_export = [
        'Installed Product ID', 
        'Parent: Installed Product', 
        'Master Device: Installed Product', 
        'Sibex Name', 
        'Product: Product Code', 
        'Full Location Movex Id'
    ]

    parent_master_equals_installed_product_id = IPs_with_no_duplicated_father_filtered[columns_to_export].copy()

    for col in ['Parent: Installed Product', 'Master Device: Installed Product']:
        parent_master_equals_installed_product_id[col] = (
            parent_master_equals_installed_product_id[col]
            .astype(str)
            .map(mapping_dict)
            .fillna(parent_master_equals_installed_product_id[col])
        )

    # Compare original mapped base vs current correction applied
    correction_sub = IPs_Included_in_setup_cleaned[['Installed Product ID', 'Parent: Installed Product', 'Master Device: Installed Product']].copy()
    correction_sub.columns = ['Installed Product ID', 'Parent_Correct', 'Master_Correct']

    df_comparison = parent_master_equals_installed_product_id.merge(correction_sub, on='Installed Product ID', how='left')

    def standardize(value):
        val_str = str(value).strip().lower()
        if val_str in ['nan', 'none', '<na>', '', 'nat', 'null']:
            return 'null_value'
        return val_str

    p_old = df_comparison['Parent: Installed Product'].apply(standardize)
    p_new = df_comparison['Parent_Correct'].apply(standardize)
    m_old = df_comparison['Master Device: Installed Product'].apply(standardize)
    m_new = df_comparison['Master_Correct'].apply(standardize)

    mask_different = (p_old != p_new) | (m_old != m_new)
    is_alphanumeric = df_comparison['Parent: Installed Product'].astype(str).str.match(r'^[a-zA-Z0-9]+$', na=False)

    df_comparison['Line affected?'] = 'No'
    df_comparison.loc[mask_different & is_alphanumeric, 'Line affected?'] = 'Yes'

    # Filter to isolate only the affected rows
    affected_ids = df_comparison.loc[df_comparison['Line affected?'] == 'Yes', 'Installed Product ID']
    correction_filtered = IPs_Included_in_setup_cleaned[IPs_Included_in_setup_cleaned['Installed Product ID'].isin(affected_ids)].copy()
    
    # Filter out records where the newly generated Master Device is empty or null
    correction_filtered = correction_filtered[
        correction_filtered['Master Device: Installed Product'].notna() & 
        (correction_filtered['Master Device: Installed Product'].astype(str).str.strip() != '')
    ].copy()

    # Export only the final corrected spreadsheet to Excel
    correction_filtered.to_excel(result_name, index=False)
    
    return correction_filtered

def clean_status(df, product_hierarchy, cutoff_date_str, result_name):
    """
    Cleans structural status alignment anomalies based on chronological hierarchy metrics,
    drops locations with duplicated fathers, and writes the output to an Excel spreadsheet.
    """
    # Parse parameter date structures
    cutoff_date = pd.to_datetime(cutoff_date_str)
    
    df_working = df.copy()
    df_working['Created Date'] = pd.to_datetime(df_working['Created Date'], errors='coerce')
    df_working['Date Shipped'] = pd.to_datetime(df_working['Date Shipped'], errors='coerce')
    df_working['Customer/Device Acceptance Date'] = pd.to_datetime(df_working['Customer/Device Acceptance Date'], errors='coerce')
    
    # CRITICAL: Exclude locations containing duplicated parent structures
    duplicated_father_ids = find_locations_with_duplicated_fathers(df_working, product_hierarchy)
    df_filtered = filter_df(df_working, duplicated_father_ids)
    
    if df_filtered.empty:
        empty_df = pd.DataFrame()
        empty_df.to_excel(result_name, index=False)
        return empty_df
        
    # Evaluate chronological logic cutoff metrics per unique movex structural node group
    group_dates = df_filtered.groupby('Full Location Movex Id').agg({'Created Date': 'max', 'Date Shipped': 'max'})
    ids_to_process = group_dates[(group_dates['Created Date'] < cutoff_date) & (group_dates['Date Shipped'] < cutoff_date)].index
    
    df_to_process = df_filtered[df_filtered['Full Location Movex Id'].isin(ids_to_process)].copy()
    
    # Remove groups already presenting uniform singular status profiles
    if not df_to_process.empty:
        group_status_counts = df_to_process.groupby('Full Location Movex Id')['Status'].nunique()
        uniform_status_groups = group_status_counts[group_status_counts == 1].index
        df_to_process = df_to_process[~df_to_process['Full Location Movex Id'].isin(uniform_status_groups)]
        
    # Execute structural cascading evaluation logic per structural node grouping
    if not df_to_process.empty:
        processed_data = df_to_process.groupby('Full Location Movex Id', group_keys=False).apply(_check_parent_status_consistency)
        if 'Full Location Movex Id' not in processed_data.columns:
            processed_data = processed_data.reset_index()
    else:
        processed_data = pd.DataFrame()
        
    # Filter output layout properties
    selected_cols = [
        'Installed Product ID', 'right sts', 'Status', 'sts equals parent?', 'Date Installed',
        'Full Location Movex Id', 'Product: Product Code', 'Customer/Device Acceptance Date', 
        'Quantity', 'Account Region', 'Created Date', 'Date Shipped'
    ]
    
    cols_processed = [c for c in selected_cols if c in processed_data.columns]
    final_output = processed_data[cols_processed].copy() if not processed_data.empty else pd.DataFrame(columns=selected_cols)
    
    if not final_output.empty and 'Date Installed' in final_output.columns:
        final_output['Date Installed'] = final_output['Date Installed'].dt.date
        
    # Export the final single targeted spreadsheet layout to Excel
    final_output.to_excel(result_name, index=False)
    
    return final_output

# --- PRIVATE HELPER FUNCTIONS ---

def find_locations_with_duplicated_fathers(df, product_hierarchy):
    """[PRIVATE] Returns an array of full location movex id containing the locations that have duplicated parent product codes based on the product hierarchy"""
    def get_all_codes_with_children(hierarchy_node):
        results = []
        if isinstance(hierarchy_node, list):
            for item in hierarchy_node:
                results.extend(get_all_codes_with_children(item))
        elif isinstance(hierarchy_node, dict):
            if hierarchy_node.get("Children"):
                if hierarchy_node.get("Product Code"):
                    results.append(hierarchy_node["Product Code"])
                results.extend(get_all_codes_with_children(hierarchy_node["Children"]))
        return list(set(results))
    
    complete_list = get_all_codes_with_children(product_hierarchy)
    duplicated = df[
        df['Product: Product Code'].isin(complete_list) & 
        df.duplicated(subset=['Full Location Movex Id', 'Product: Product Code'], keep=False)
    ]
    ids_unicos = duplicated['Full Location Movex Id'].unique().astype(str)
    return ids_unicos

def filter_df(df, full_location_movex_id_list):
    """[PRIVATE] Filters a df based on a list of full location movex id"""
    return df.groupby('Full Location Movex Id').filter(lambda group: group.name not in full_location_movex_id_list)

def filter_df_by_hierarchy_json(df, json_path):
    """[PRIVATE] Loads a hierarchy JSON file from a specific path, extracts product names and sibex names, and filters rows."""
    with open(json_path, 'r', encoding='utf-8') as file:
        hierarchy_json = json.load(file)

    allowed_sibex_names = set()
    allowed_product_names = set()

    def extract_names_recursively(node_list):
        nodes = node_list if isinstance(node_list, list) else [node_list]
        for node in nodes:
            if isinstance(node, dict):
                if node.get("Sibex Name"):
                    allowed_sibex_names.add(str(node["Sibex Name"]).strip().upper())
                if node.get("Product Name"):
                    allowed_product_names.add(str(node["Product Name"]).strip().upper())
                children = node.get("Children", [])
                if children:
                    extract_names_recursively(children)

    extract_names_recursively(hierarchy_json)
    df_sibex_upper = df['Sibex Name'].astype(str).str.strip().str.upper()
    df_product_upper = df['Product: Product Name'].astype(str).str.strip().str.upper()

    mask_sibex = df_sibex_upper.isin(allowed_sibex_names)
    mask_product = df_product_upper.isin(allowed_product_names)
    return df[mask_sibex | mask_product].copy()

def filter_hierarchy_by_allowed_codes(hierarchy_node, allowed_codes):
    """[PRIVATE] Recursively filters a hierarchy tree, keeping only nodes within the allowed_codes set."""
    if isinstance(hierarchy_node, list):
        filtered_list = []
        for item in hierarchy_node:
            filtered_item = filter_hierarchy_by_allowed_codes(item, allowed_codes)
            if filtered_item:
                filtered_list.append(filtered_item)
        return filtered_list if filtered_list else None
        
    elif isinstance(hierarchy_node, dict):
        node_code = str(hierarchy_node.get("Product Code", "")).strip()
        children = hierarchy_node.get("Children", [])
        filtered_children = []
        if children:
            res_children = filter_hierarchy_by_allowed_codes(children, allowed_codes)
            if res_children:
                filtered_children = res_children if isinstance(res_children, list) else [res_children]
        
        if node_code in allowed_codes or filtered_children:
            new_node = hierarchy_node.copy()
            if filtered_children:
                new_node["Children"] = filtered_children
            elif "Children" in new_node:
                del new_node["Children"]
            return new_node
    return None

def clear_parent_master_columns(df):
    """[PRIVATE] Clears columns and sets the type to object to ensure compatibility with string IDs."""
    df = df.copy()
    df['Parent: Installed Product'] = None
    df['Master Device: Installed Product'] = None
    return df

def populate_master_device_by_hierarchy(df, all_mosaiq_list):
    """[PRIVATE] Identifies the Root Parent. If the product is one of the root codes, sets Parent/Master to None."""
    df = df.copy()
    hierarchy_code_map = {}
    root_nodes = all_mosaiq_list if isinstance(all_mosaiq_list, list) else [all_mosaiq_list]
    root_codes = {item.get("Product Code").strip() for item in root_nodes if isinstance(item, dict) and item.get("Product Code")}

    def extract_hierarchy_recursively(product_list):
        nodes = product_list if isinstance(product_list, list) else [product_list]
        for item in nodes:
            if isinstance(item, dict):
                parent_code = item.get("Product Code")
                children = item.get("Children", [])
                for child in children:
                    if isinstance(child, dict):
                        child_code = child.get("Product Code")
                        if child_code and parent_code:
                            hierarchy_code_map[child_code.strip()] = parent_code.strip()
                        if child.get("Children"):
                            extract_hierarchy_recursively([child])

    extract_hierarchy_recursively(all_mosaiq_list)
    
    def find_root_code(product_code):
        current_code = str(product_code).strip()
        while current_code in hierarchy_code_map:
            current_code = hierarchy_code_map[current_code]
        return current_code

    def assign_hierarchy_logic(group):
        local_id_map = {
            str(code).strip(): idx 
            for code, idx in zip(group['Product: Product Code'], group['Installed Product ID'])
        }
        
        def process_row(row):
            current_code = str(row['Product: Product Code']).strip()
            if current_code in root_codes:
                return pd.Series([None, None])
            
            parent_code = hierarchy_code_map.get(current_code)
            parent_id = local_id_map.get(parent_code) if parent_code else None
            root_code = find_root_code(current_code)
            master_id = local_id_map.get(root_code) if root_code != current_code else None
            return pd.Series([parent_id, master_id])

        group[['Parent: Installed Product', 'Master Device: Installed Product']] = group.apply(process_row, axis=1)
        return group

    return df.groupby('Full Location Movex Id', group_keys=False).apply(assign_hierarchy_logic).reset_index(drop=True)

def populate_parent_by_hierarchy_dict(df, hierarchy_data):
    """[PRIVATE] Populates the Parent column by checking the hierarchy dictionary."""
    df = df.copy()
    hierarchy_code_map = {}
    
    def extract_hierarchy_recursively(product_list):
        nodes = product_list if isinstance(product_list, list) else [product_list]
        for item in nodes:
            if isinstance(item, dict):
                parent_code = item.get("Product Code")
                children = item.get("Children", [])
                for child in children:
                    if isinstance(child, dict):
                        child_code = child.get("Product Code")
                        if child_code and parent_code:
                            hierarchy_code_map[child_code.strip()] = parent_code.strip()
                        if child.get("Children"):
                            extract_hierarchy_recursively([child])

    extract_hierarchy_recursively(hierarchy_data)
    
    def assign_parent_id(group):
        local_id_map = {
            str(code).strip(): idx 
            for code, idx in zip(group['Product: Product Code'], group['Installed Product ID'])
        }
        
        def find_best_ancestor(row):
            current_code = str(row['Product: Product Code']).strip()
            master_id = row['Master Device: Installed Product']
            temp_code = current_code
            while temp_code in hierarchy_code_map:
                parent_code = hierarchy_code_map[temp_code]
                if parent_code in local_id_map:
                    return local_id_map[parent_code]
                temp_code = parent_code
            return master_id

        group['Parent: Installed Product'] = group.apply(find_best_ancestor, axis=1)
        return group

    return df.groupby('Full Location Movex Id', group_keys=False).apply(assign_parent_id).reset_index(drop=True)

def _check_parent_status_consistency(group):
    """[PRIVATE] Resolves tracking state alignment cascading loops rules inside local dataframe node groups."""
    parent_dict = group.set_index('Installed Product')['Parent: Installed Product'].to_dict()
    status_dict = group.set_index('Installed Product')['Status'].to_dict()
    latest_acceptance_date = group['Customer/Device Acceptance Date'].min()

    # Rule 2: Explicit alignment check for MOSAIQ and MOSAIQ Software components
    mosaiq_rows = group[group['Sibex Name'] == 'MOSAIQ']
    mosaiq_sw_rows = group[group['Sibex Name'] == 'MOSAIQ Software']
    
    if not mosaiq_rows.empty and not mosaiq_sw_rows.empty:
        sts_mosaiq = mosaiq_rows.iloc[0]['Status']
        sts_mosaiq_sw = mosaiq_sw_rows.iloc[0]['Status']
        
        if sts_mosaiq == sts_mosaiq_sw:
            group['right sts'] = sts_mosaiq
            group['sts equals parent?'] = group.apply(lambda r: "True" if r['Status'] == r['right sts'] else "False", axis=1)
            group['Date Installed'] = pd.NaT
            mask = (group['sts equals parent?'] == "False") & (group['right sts'] == "Installed")
            group.loc[mask, 'Date Installed'] = latest_acceptance_date
            return group

    # Rule 3: Evaluation using majority density vote versus root item state profile
    majority_sts = group['Status'].mode()[0]
    roots = group[group['Parent: Installed Product'].isna()]
    
    if not roots.empty:
        root_sts = roots.iloc[0]['Status'] 
        if root_sts != majority_sts:
            group['right sts'] = majority_sts
            group['sts equals parent?'] = group.apply(lambda r: "True" if r['Status'] == r['right sts'] else "False", axis=1)
            group['Date Installed'] = pd.NaT
            mask = (group['sts equals parent?'] == "False") & (group['right sts'] == "Installed")
            group.loc[mask, 'Date Installed'] = latest_acceptance_date
            return group

    # Rule 4: Sequential top-down root ancestor resolution algorithm (Fallback resolution)
    def _get_ultimate_root_status(item_id):
        curr = item_id
        visited = set()
        while pd.notna(curr) and curr in parent_dict and curr not in visited:
            visited.add(curr)
            parent = parent_dict[curr]
            if pd.isna(parent) or parent not in parent_dict:
                break
            curr = parent
        return status_dict.get(curr, status_dict.get(item_id))

    group['right sts'] = group['Installed Product'].apply(_get_ultimate_root_status)
    group['sts equals parent?'] = group.apply(lambda r: "True" if r['Status'] == r['right sts'] else "False", axis=1)
    group['Date Installed'] = pd.NaT
    mask = (group['sts equals parent?'] == "False") & (group['right sts'] == "Installed")
    group.loc[mask, 'Date Installed'] = latest_acceptance_date
    
    return group
