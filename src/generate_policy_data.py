"""Generate sample policy data for actuarial analysis"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
from config import *

class PolicyDataGenerator:
    def __init__(self, num_policies=DEFAULT_NUM_POLICIES, seed=RANDOM_SEED):
        """Initialize the generator with number of policies and random seed"""
        self.num_policies = num_policies
        self.seed = seed
        self.current_date = datetime(2025, 4, 16)  # Fixed date for reproducibility
        self._set_seeds()
        
    def _set_seeds(self):
        """Set all random seeds for reproducibility"""
        random.seed(self.seed)
        np.random.seed(self.seed)
    
    def _generate_policy_numbers(self):
        """Generate unique policy numbers"""
        return [f'S20TL{str(i).zfill(5)}' for i in range(1, self.num_policies + 1)]
    
    def _generate_dates_of_birth(self):
        """Generate birth dates for fixed entry age"""
        birth_year = self.current_date.year - ENTRY_AGE
        return [
            datetime(birth_year, random.randint(1, 12), random.randint(1, 28))
            for _ in range(self.num_policies)
        ]
    
    def _generate_purchase_dates(self):
        """Generate policy purchase dates within specified range"""
        max_days = PURCHASE_DATE_RANGE_YEARS * 365
        return [
            self.current_date - timedelta(days=random.randint(1, max_days))
            for _ in range(self.num_policies)
        ]
    
    def _generate_terms(self):
        """Generate policy terms"""
        return np.random.choice(TERMS, self.num_policies)
    
    def _generate_sum_assured(self):
        """Generate sum assured amounts"""
        return np.array([BASE_SA * random.randint(1, 10) for _ in range(self.num_policies)])
    
    def _calculate_premiums(self, sum_assured):
        """Calculate annual premiums"""
        return sum_assured * PREMIUM_RATE
    
    def _generate_policy_status(self):
        """Generate policy status with controlled probabilities"""
        return np.random.choice(
            list(STATUS_PROBABILITIES.keys()),
            self.num_policies,
            p=list(STATUS_PROBABILITIES.values())
        )
    
    def _generate_death_dates(self, purchase_dates, policy_status):
        """Generate death dates for policies with death claims"""
        death_dates = []
        for purchase_date, status in zip(purchase_dates, policy_status):
            if status == 'Death Claim':
                max_days = (self.current_date - purchase_date).days
                if max_days > 0:
                    death_date = purchase_date + timedelta(
                        days=random.randint(1, max_days)
                    )
                else:
                    death_date = self.current_date
                death_dates.append(death_date)
            else:
                death_dates.append(None)
        return death_dates
    
    def generate(self):
        """Generate complete policy dataset"""
        # Generate basic data
        policy_numbers = self._generate_policy_numbers()
        dates_of_birth = self._generate_dates_of_birth()
        purchase_dates = self._generate_purchase_dates()
        terms = self._generate_terms()
        sum_assured = self._generate_sum_assured()
        annual_premiums = self._calculate_premiums(sum_assured)
        policy_status = self._generate_policy_status()
        death_dates = self._generate_death_dates(purchase_dates, policy_status)
        
        # Create DataFrame
        df = pd.DataFrame({
            'Policy_Number': policy_numbers,
            'Date_of_Birth': dates_of_birth,
            'Entry_Age': ENTRY_AGE,
            'Purchase_Date': purchase_dates,
            'Policy_Term': terms,
            'Premium_Payment_Term': terms,
            'Sum_Assured': sum_assured,
            'Annual_Premium': annual_premiums,
            'Premium_Payment_Timing': 'Beginning of Year',
            'Policy_Status': policy_status,
            'Death_Date': death_dates,
            'Underwriting_Class': 'Ultimate Mortality',
            'Surrender_Value': 'None (Pure Term Plan)',
            'Reserve_Basis': 'Prospective Reserve'
        })
        
        # Calculate expiry dates
        df['Expiry_Date'] = df.apply(
            lambda x: x['Purchase_Date'] + timedelta(days=int(x['Policy_Term']*365)), 
            axis=1
        )
        
        # Format dates
        for col in ['Date_of_Birth', 'Purchase_Date', 'Expiry_Date', 'Death_Date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
        
        return df

def save_to_excel(df, filename):
    """Save the DataFrame to Excel with proper formatting"""
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Policy_Data', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Policy_Data']
        
        # Define formats
        formats = {
            'header': workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#D9E1F2',
                'border': 1
            }),
            'money': workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1
            }),
            'date': workbook.add_format({
                'num_format': 'yyyy-mm-dd',
                'border': 1
            }),
            'text': workbook.add_format({
                'border': 1
            })
        }
        
        # Column definitions with widths and formats
        columns = {
            'A': ['Policy_Number', 15, 'text'],
            'B': ['Date_of_Birth', 12, 'date'],
            'C': ['Entry_Age', 10, 'text'],
            'D': ['Purchase_Date', 12, 'date'],
            'E': ['Policy_Term', 10, 'text'],
            'F': ['Premium_Payment_Term', 10, 'text'],
            'G': ['Sum_Assured', 15, 'money'],
            'H': ['Annual_Premium', 15, 'money'],
            'I': ['Premium_Payment_Timing', 20, 'text'],
            'J': ['Policy_Status', 12, 'text'],
            'K': ['Death_Date', 12, 'date'],
            'L': ['Expiry_Date', 12, 'date'],
            'M': ['Underwriting_Class', 15, 'text'],
            'N': ['Surrender_Value', 20, 'text'],
            'O': ['Reserve_Basis', 15, 'text']
        }
        
        # Apply formats
        for col, (_, width, format_name) in columns.items():
            worksheet.set_column(f'{col}:{col}', width, formats[format_name])
        
        # Format headers
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, formats['header'])
        
        # Add table with built-in autofilter
        worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
            'style': 'Table Style Medium 2',
            'columns': [{'header': col} for col in df.columns]
        })
        
        # Freeze panes
        worksheet.freeze_panes(1, 0)

if __name__ == "__main__":
    # Generate data
    generator = PolicyDataGenerator()
    policy_data = generator.generate()
    
    # Save to Excel
    output_path = '../outputs/Secure20_Term_Life_Data.xlsx'
    save_to_excel(policy_data, output_path)
    
    print(f"Generated {len(policy_data)} policy records")
    print(f"Data saved to: {output_path}")
    print("\nSample Records:")
    print(policy_data[['Policy_Number', 'Entry_Age', 'Policy_Term', 
                      'Sum_Assured', 'Annual_Premium', 'Policy_Status']].head())