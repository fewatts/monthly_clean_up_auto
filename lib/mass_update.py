"""MASS UPDATE"""

import pandas as pd
import json

def clean_quantity(df):
    """Returns a DF containing the ready to go spreadsheet"""
    result = df[(df['Quantity'] < 0) | (df['Quantity'].isna())].copy()
    result['Quantity'] = result['Quantity'].fillna(1)
    result.loc[result['Quantity'] < 0, 'Quantity'] *= -1
    return result

def find_locations_with_duplicated_fathers(df, product_hierarchy):
    """Returns an array of full location movex id containing the locations that have duplicated parent product codes based on the product hierarchy"""
    def get_all_codes_with_children(hierarchy_node):
        results = []
        
        # If it's a list, iterate through its elements
        if isinstance(hierarchy_node, list):
            for item in hierarchy_node:
                results.extend(get_all_codes_with_children(item))
                
        # If it's a dictionary, safely check for children
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
    """Filters a df based on a list of full location movex id"""
    return df.groupby('Full Location Movex Id').filter(lambda group: group.name not in full_location_movex_id_list)

def filter_df_by_hierarchy_json(df, json_path):
    """
    Loads a hierarchy JSON file from a specific path, extracts all product 
    names and sibex names, and filters the DataFrame to keep only matching rows.
    """
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

    filtered_df = df[mask_sibex | mask_product].copy()
    return filtered_df

def clear_parent_master_columns(df):
    """
    Clears columns and sets the type to object to ensure compatibility with string IDs.
    """
    df = df.copy()
    df['Parent: Installed Product'] = None
    df['Master Device: Installed Product'] = None
    return df

def populate_master_device_by_hierarchy(df, all_mosaiq_list):
    """
    Identifies the Root Parent. If the product is one of the root codes, 
    sets Parent and Master Device to None.
    """
    df = df.copy()
    
    hierarchy_code_map = {}
    
    # Safely extract root codes handling both list and dictionary structures at the root level
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
    """
    Populates the Parent column by checking the hierarchy dictionary.
    """
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

def clean_product_hierarchy(df, product_hierarchy, result_name):
    """Cleans the installed product hierarchy columns based on the product hierarchy"""
    
    # 1. Find IDs with duplicated fathers and remove them from processing
    duplicated_father_ids = find_locations_with_duplicated_fathers(df, product_hierarchy)
    IPs_with_no_duplicated_father_filtered = filter_df(df, duplicated_father_ids)
    
    # 2. Run the cleanup and population chain based on the JSON structure
    IPs_Included_in_setup_cleaned = clear_parent_master_columns(IPs_with_no_duplicated_father_filtered)
    IPs_Included_in_setup_cleaned = populate_master_device_by_hierarchy(IPs_Included_in_setup_cleaned, product_hierarchy)
    IPs_Included_in_setup_cleaned = populate_parent_by_hierarchy_dict(IPs_Included_in_setup_cleaned, product_hierarchy)

    columns_to_keep = ['Installed Product ID', 'Parent: Installed Product', 'Master Device: Installed Product', 
                       'Sibex Name', 'Product: Product Code', 'Full Location Movex Id']
    IPs_Included_in_setup_cleaned = IPs_Included_in_setup_cleaned[columns_to_keep].copy()

    # 3. Map current base, replacing old values with the correct IDs
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

    # 4. Compare original mapped base vs current correction applied
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

    affected_ids = df_comparison.loc[df_comparison['Line affected?'] == 'Yes', 'Installed Product ID']
    correction_filtered = IPs_Included_in_setup_cleaned[IPs_Included_in_setup_cleaned['Installed Product ID'].isin(affected_ids)].copy()

    df_comparison['Diferença_Parent'] = (
        df_comparison['Parent: Installed Product'].astype(str) + 
        " -> " + 
        df_comparison['Parent_Correct'].astype(str)
    )

    # 5. Export generated reports
    df_comparison.to_excel('lines_affected_in_hierarchy_clean_up.xlsx', index=False)
    correction_filtered.to_excel(result_name, index=False)
    
    return correction_filtered