"""MANUAL"""

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
