import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def set_all_seeds(seed_value=42):
    """Set all random seeds to ensure reproducibility"""
    random.seed(seed_value)
    np.random.seed(seed_value)
    
def generate_policy_data(num_policies=1000, seed=42):
    """Generate sample policy data for Secure20 Term Life Insurance"""
    # Set all random seeds
    set_all_seeds(seed)
    
    # Product specifications
    ENTRY_AGE = 30
    TERMS = [5, 10, 15]
    BASE_SA = 10000
    PREMIUM_RATE = 0.05  # 5% of SA
    
    # Fix the current date to ensure reproducibility
    current_date = datetime(2025, 4, 16)  # Using the context date
    
    # Generate policy numbers (format: S20TLXXXXX)
    policy_numbers = [f'S20TL{str(i).zfill(5)}' for i in range(1, num_policies + 1)]
    
    # Calculate birth dates for entry age 30
    birth_year = current_date.year - ENTRY_AGE
    dates_of_birth = [
        datetime(birth_year, 
                random.randint(1, 12), 
                random.randint(1, 28))
        for _ in range(num_policies)
    ]
    
    # Generate policy terms
    terms = np.random.choice(TERMS, num_policies)
    
    # Generate sum assured (multiples of 10000)
    sum_assured = np.array([BASE_SA * random.randint(1, 10) for _ in range(num_policies)])
    
    # Generate policy issue dates (within last 2 years)
    purchase_dates = [
        current_date - timedelta(days=random.randint(1, 2*365))
        for _ in range(num_policies)
    ]
    
    # Calculate annual premiums (5% of SA)
    annual_premiums = sum_assured * PREMIUM_RATE
    
    # Generate policy status with controlled randomization
    status_random = np.random.random(num_policies)
    policy_status = ['Death Claim' if x < 0.05 else 'In Force' for x in status_random]
    
    # Create DataFrame
    df = pd.DataFrame({
        'Policy_Number': policy_numbers,
        'Date_of_Birth': dates_of_birth,
        'Entry_Age': ENTRY_AGE,
        'Purchase_Date': purchase_dates,
        'Policy_Term': terms,
        'Premium_Payment_Term': terms,  # Same as policy term
        'Sum_Assured': sum_assured,
        'Annual_Premium': annual_premiums,
        'Premium_Payment_Timing': 'Beginning of Year',
        'Policy_Status': policy_status,
        'Underwriting_Class': 'Ultimate Mortality',
        'Surrender_Value': 'None (Pure Term Plan)',
        'Reserve_Basis': 'Prospective Reserve'
    })
    
    # Add derived columns
    df['Expiry_Date'] = df.apply(
        lambda x: x['Purchase_Date'] + timedelta(days=int(x['Policy_Term']*365)), 
        axis=1
    )
    
    # For death claims, generate random death dates between purchase and current date
    death_dates = []
    for idx, row in df.iterrows():
        if row['Policy_Status'] == 'Death Claim':
            max_days = (current_date - row['Purchase_Date']).days
            if max_days > 0:  # Ensure positive days
                death_date = row['Purchase_Date'] + timedelta(
                    days=random.randint(1, max_days)
                )
            else:
                death_date = current_date
            death_dates.append(death_date)
        else:
            death_dates.append(None)
    df['Death_Date'] = death_dates
    
    # Format dates
    df['Date_of_Birth'] = pd.to_datetime(df['Date_of_Birth']).dt.strftime('%Y-%m-%d')
    df['Purchase_Date'] = pd.to_datetime(df['Purchase_Date']).dt.strftime('%Y-%m-%d')
    df['Expiry_Date'] = pd.to_datetime(df['Expiry_Date']).dt.strftime('%Y-%m-%d')
    df['Death_Date'] = pd.Series(death_dates).dt.strftime('%Y-%m-%d')
    
    return df

def save_to_excel(df, filename='Secure20_Term_Life_Data.xlsx'):
    """Save the DataFrame to Excel with proper formatting"""
    # Create Excel writer object
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Write the data
        df.to_excel(writer, sheet_name='Policy_Data', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Policy_Data']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        money_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })
        
        date_format = workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'border': 1
        })
        
        text_format = workbook.add_format({
            'border': 1
        })
        
        # Set column widths and formats
        columns = {
            'A': ['Policy_Number', 15, text_format],
            'B': ['Date_of_Birth', 12, date_format],
            'C': ['Entry_Age', 10, text_format],
            'D': ['Purchase_Date', 12, date_format],
            'E': ['Policy_Term', 10, text_format],
            'F': ['Premium_Payment_Term', 10, text_format],
            'G': ['Sum_Assured', 15, money_format],
            'H': ['Annual_Premium', 15, money_format],
            'I': ['Premium_Payment_Timing', 20, text_format],
            'J': ['Policy_Status', 12, text_format],
            'K': ['Death_Date', 12, date_format],
            'L': ['Expiry_Date', 12, date_format],
            'M': ['Underwriting_Class', 15, text_format],
            'N': ['Surrender_Value', 20, text_format],
            'O': ['Reserve_Basis', 15, text_format]
        }
        
        # Apply formats
        for col, (_, width, cell_format) in columns.items():
            worksheet.set_column(f'{col}:{col}', width, cell_format)
        
        # Format headers
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, header_format)
        
        # Add table with built-in autofilter
        worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
            'style': 'Table Style Medium 2',
            'columns': [{'header': col} for col in df.columns]
        })
        
        # Freeze panes
        worksheet.freeze_panes(1, 0)

if __name__ == "__main__":
    # Generate sample data with fixed seed
    policy_data = generate_policy_data(num_policies=1000, seed=42)
    
    # Save to Excel with formatting
    save_to_excel(policy_data)
    
    print("Generated 100 policy records with seed=42")
    print("\nSample Records:")
    print(policy_data[['Policy_Number', 'Entry_Age', 'Policy_Term', 
                      'Sum_Assured', 'Annual_Premium', 'Policy_Status']].head())