import pandas as pd 
import os
import re
from numpy import setdiff1d
from PathManager import locationManager as lm
from ErrorLogging import error_popup
import openpyxl

class MasterReferenceUpdater:
    def __init__(self, current_path = ''):
        self.FAILED = 'Failures:' #length == 9 for check
        self.lm = lm(current_path) 
        self.ref_path = self.lm.get_reference_path()
        self.input_path = self.lm.get_input_path()
        self.archive_path = self.lm.get_archive_path()
        self.master_ref_df = self.load_master_ref()
        self.new_customer_df = None 
    
    def append_FAILED(self, msg):
        self.FAILED = self.FAILED + '\n' + msg
        
    def load_master_ref(self):
        """Load existing master reference into dataframe."""
        try:
            return pd.read_excel(self.ref_path, sheet_name='CustomerList', engine='openpyxl')
        except:
            self.append_FAILED('Failed to load Master Ref CustomerList')
            return False

    def load_master_price_ref(self):
        """Load existing master reference Price Sheet into dataframe"""
        try:
            return pd.read_excel(self.ref_path, sheet_name='PriceSheet', engine='openpyxl')
        except:
            self.append_FAILED('Failed to load Master Ref PriceSheet')
            return False
    
    def load_newest_files(self):
        """Get the new files containing the new customer list and load it into a dataframe."""
        input_files = os.listdir(self.input_path)
        raw_pattern = 'customer'
        for input_file in input_files:
            if re.search(raw_pattern, input_file):
                if '.csv' in input_file:
                    self.new_customer_df = pd.read_csv(os.path.join(self.input_path, input_file))
                else:
                    self.new_customer_df = pd.read_excel(os.path.join(self.input_path, input_file), engine='openpyxl')
                return True
        return False
    
    def compare_new_to_existing_reference(self):
        """Identify customers that currently don't exist in the master reference."""
        old_customers = list(self.master_ref_df['Record ID'].unique())
        temp_customers = list(self.new_customer_df['Record ID'].unique())
        new_customers = setdiff1d(temp_customers, old_customers)
        temp = self.new_customer_df.copy()
        temp['KEEP'] = temp['Record ID'].apply(lambda x: True if x in new_customers else False)
        result = temp[temp['KEEP']==True]
        result.drop(['KEEP'], axis=1, inplace=True)
        self.new_customer_df = result

    def overwrite_old_PLN_Nos(self):
        taken_PLN_Nos = list(self.new_customer_df['PLN Stock Lens Account Number'].unique())
        self.master_ref_df['KEEP'] = self.master_ref_df['PLN Stock Lens Account Number'].apply(lambda x: True if x not in taken_PLN_Nos else False)
        self.master_ref_df = self.master_ref_df[self.master_ref_df['KEEP']]
        self.master_ref_df.drop(columns=['KEEP'], inplace=True)
    
    def update_reference(self):
        """Add new customers into existing reference."""
        self.master_ref_df = self.master_ref_df.append(self.new_customer_df)
        return True
    
    def save_reference(self):
        """Save master reference."""
        price_reference = self.load_master_price_ref()
        if True  in [self.master_ref_df.empty, price_reference.empty]:
            self.append_FAILED('Failed to save Master reference.')
            return False
        writer = pd.ExcelWriter(self.ref_path, engine='xlsxwriter')
        self.master_ref_df.to_excel(writer, index=False, sheet_name='CustomerList')
        price_reference.to_excel(writer, index=False, sheet_name='PriceSheet')
        writer.save()
        return True
    
    def RUN(self):
        new_customers_found = self.load_newest_files()
        if new_customers_found:
            self.compare_new_to_existing_reference()
            self.overwrite_old_PLN_Nos() #Make sure solution is acceptable
            self.update_reference()
            self.save_reference()
        else:
            print('No new customers found.') 
        if len(self.FAILED)>9:
            error_popup(self.FAILED)

if __name__ == '__main__':
    pass