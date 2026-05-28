"""MASS UPDATE"""

def clean_quantity(df):
    """Returns a DF containing the ready to go spreadsheet"""
    result = df[(df['Quantity'] < 0) | (df['Quantity'].isna())].copy()
    result['Quantity'] = result['Quantity'].fillna(1)
    result.loc[result['Quantity'] < 0, 'Quantity'] *= -1
    return result

def convert_to_excel(df, result_name):
    """Given a name, converts the data frame to excel"""
    return df.to_excel(result_name, index=False)