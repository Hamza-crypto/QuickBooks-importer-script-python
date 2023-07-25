from cv2 import dft
import pandas as pd
import sys
import os
from datetime import date, datetime
from calendar import monthrange, month_name
import re
import tkinter as tk
import openpyxl
import traceback
from MasterReferenceUpdater import MasterReferenceUpdater
class SanityCheck:
    
    def __init__(self):
        """
        Does the checks for Master Reference to prevent errors from getting into the invoice generator.
        Desired Features (copied from customer message)
        1. If there's a customer that has a Pivotal Group number, but no Stock Lens account number? (Or vis versa). So basically if there was a blank box in either of the columns for this info
        2. If there's a duplicate Pivotal Group number identified
        3. If there's a duplicate Stock lens account number identified.
        """
        self.current_location = os.path.dirname(sys.executable)#os.path.dirname(os.path.abspath(__file__))
        reference_path = os.path.join(self.current_location, 'MasterReference.xlsx')
        self.price_reference = pd.read_excel(reference_path, sheet_name='PriceSheet', engine='openpyxl')
        self.customer_list = pd.read_excel(reference_path, sheet_name='CustomerList', engine='openpyxl')
        self.Passed = False
        self.LOG = pd.DataFrame(columns = ['Description', 'Location'])

    def append_LOG(self, description, location):
        """Add a line to the error log describing the issue with Master Reference."""
        temp = pd.DataFrame()
        temp['Description'] = [description]
        temp['Location'] = [location]
        self.LOG = self.LOG.append(temp)

    def check_for_duplicates(self, column):
        """Check to see if column contains duplicates."""
        def id_duplicate(l):
            duplicates = [str(i) for i in l if l.count(i)>1]
            return ' '.join(list(set(duplicates)))
        if len(self.customer_list[column].unique()) != len(list(self.customer_list[column])):
            duplicates = id_duplicate(list(self.customer_list[column]))
            self.append_LOG(f'Duplicate Values in {column}', 'CustomerList, MasterReference')
            self.append_LOG(f'--->{column} duplicates are:', duplicates)
            return False
        else:
            return True

    def check_for_missing(self, column):
        """Check to see if column contains any NaN"""
        nulls = self.customer_list[column].isnull().values.any()
        if nulls:
            self.append_LOG(f'Missing Values in {column}', 'CustomerList, MasterReference')
            return False
        else:
            return True

    def all_prices_present(self):
        """Check to see if price reference chart is fully filled out."""
        nulls = self.price_reference['Retail'].isnull().values.any()    
        if nulls:
            self.append_LOG('Missing Values', 'PriceSheet, MasterReference')
            return False
        else:
            return True 
        
    def run_check(self):
        """Run full suite of checks for reference sheet. """
        check_1 = self.check_for_duplicates('PLN Stock Lens Account Number')
        check_2 = self.check_for_duplicates('Pivotal Account No.')
        check_3 = self.check_for_missing('PLN Stock Lens Account Number')
        check_4 = self.check_for_missing('Pivotal Account No.')
        check_5 = self.all_prices_present()
        all_checks  = [check_1, check_2, check_3, check_4, check_5]
        if False in all_checks:
            self.LOG.to_csv(os.path.join(self.current_location, 'REFERENCE_ERROR.csv'), index=False)
            return False
        else:
            return True

class ReportGenerator:
    def __init__(self):
        """Handles all the file locations. Note: This script will only work with files in the same directory as it."""
        self.current_location = os.path.dirname(sys.executable)#os.path.dirname(os.path.abspath(__file__))
        self.LensImport = None
        self.ShippingImport = None
        self.DiscountImport = None
        self.LensReturnsCredits = None
        self.SummarySheet = None
        self.SummaryOverviewSheet = None
        self.TaxSheet = None
        self.invoice_number_counter = 1
        self.invoice_Found = False
        self.SOMO_Disc = 0
        reference_path = os.path.join(self.current_location, 'MasterReference.xlsx')
        self.customer_list = pd.read_excel(reference_path, sheet_name='CustomerList', engine='openpyxl')
        self.price_reference = pd.read_excel(reference_path, sheet_name='PriceSheet', engine='openpyxl')
        self.save_location = os.path.join(self.current_location, 'Output')
        self.raw_invoice_path = self.get_raw_invoice()
        self.raw_invoice = pd.read_csv(self.raw_invoice_path)
        self.check_missing_DropShipNo()
        self.create_customer_suffix_key()
        self.now =self.get_month_of_invoice()

    def check_missing_DropShipNo(self):
        missings = self.raw_invoice[self.raw_invoice['DropShipNo'].isna()]
        if not missings.empty:
            missings_str = ', '.join(list(missings['PONo'].astype('str')))
            self.error_popup(f'Missing values PONo {missings_str}')

    def get_raw_invoice(self):
        """Searches the input folder for a raw invoice."""
        input_path = os.path.join(self.current_location, 'Input')
        input_files = os.listdir(input_path)
        raw_pattern = 'H00241'
        for input_file in input_files:
            if re.search(raw_pattern, input_file):
                self.invoice_Found = True
                return os.path.join(input_path, input_file)
            else:
                pass
        return ''

    def error_report(self, _msg, location):
        self.error_popup(_msg)
        self.LOG = pd.DataFrame()
        self.LOG['Description'] = [_msg]
        self.LOG['Location'] = [location]
        self.LOG.to_csv(os.path.join(self.current_location, 'ERROR_DETAILS.csv'), index=False)

    def error_popup(self, msg):
        """Super simple pop-up to indicate an error has occured."""
        popup = tk.Tk()
        popup.wm_title("!")
        label = tk.Label(popup, text=msg)
        label.pack(side="top", fill="x", pady=10)
        B1 = tk.Button(popup, text="Okay", command = popup.destroy)
        B1.pack()
        popup.mainloop()

    def create_customer_suffix_key(self):
        """Strips the Full Account Number down to the suffix. This is used by the supplier to ID customers."""
        def get_suffix_num(account_num):
            return account_num[-5:].replace('-', '').lstrip('0')
        self.customer_list['SuffixNum'] = self.customer_list['PLN Stock Lens Account Number'].apply(lambda x: get_suffix_num(x))

    def get_month_of_invoice(self):
        """Get which month this invoice is for."""
        months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        for m in months:
            if re.search(m, self.raw_invoice_path):
                this_month = months.index(m)+1
            year_pattern = "2\d\d\d"
            this_year = int(re.findall(year_pattern, self.raw_invoice_path)[0])
        return date(this_year, this_month, 1)

    def create_output_name(self):
        """Creates a filename in the format of 'Invoice Import 1 (Mmm YYYY).xlsx'"""
        this_year = str(self.now.year)
        this_month = month_name[self.now.month][:3]
        return f'Invoice Import 1 ({this_month} {this_year}).xlsx'

    def get_Dropship(self, suffix_num):
        if self.customer_list[self.customer_list['SuffixNum']==str(suffix_num)].empty:
            self.error_report(f'Customer number (DropShipNo) {suffix_num} not found in Customer list', 'MasterReference, Customer List')
        return self.customer_list[self.customer_list['SuffixNum']==str(suffix_num)]['PLN Stock Lens Account Number'].iloc[0]

    def get_Pivotal_Account(self, suffix_num):
        suffix_num = int(suffix_num)
        if self.customer_list[self.customer_list['SuffixNum']==str(suffix_num)].empty:
            self.error_report(f'Customer number {suffix_num} not found in Customer list', 'MasterReference, Customer List')
        return self.customer_list[self.customer_list['SuffixNum']==str(suffix_num)]['Pivotal Account No.'].iloc[0]
    
    def get_New_Unit_Price(self, upc):
        """Sets new price based on Barcode"""
        upc = int(upc)
        if self.price_reference[self.price_reference['UPC']==upc].empty:
            self.error_report(f'Barcode {upc} not found in Price Sheet', 'MasterReference, PriceSheet ')
        return self.price_reference[self.price_reference['UPC']==upc]['Retail'].iloc[0]

    def get_Category_Name(self, upc):
        """Sets category name based on Barcode"""
        upc = int(upc)
        if self.price_reference[self.price_reference['UPC']==upc].empty:
            self.error_report(f'Barcode {upc} not found in Price Sheet', 'MasterReference, PriceSheet ')
        return self.price_reference[self.price_reference['UPC']==upc]['Lens'].iloc[0]
        
            
    def get_sum_ShipAmount(self, customers):
        values = []
        for customer in customers:
            values.append(self.LensImport[self.LensImport['Pivotal Account']==customer]['ShipAmount'].sum())
        return values

    def get_sum_NewShipAmount(self, customers):
        values = []
        for customer in customers:
            values.append(self.LensImport[self.LensImport['Pivotal Account']==customer]['NewShipAmount'].sum())
        return values

    def get_sum_Freight(self, customers):
        values = []
        for customer in customers:   
            values.append(self.ShippingImport[self.ShippingImport['Pivotal Account']==str(customer)]['Freight'].sum())
        return values


    def generate_Lens_Import(self):
        """
        Due Date = 'Net 15'
        To Be emailed = False
        Print Later = False
        Dropship = H00241-/d/d/d/d/d; Given is \d* from supplier (see DropshipNo), use lookups to get this 
        Pivotal Account = /d/d/d/d/dA; get from lookup from Dropship
        DropShipNo = number given by supplier
        OrderID = S/d/d/d/d/d/d; get from Raw Invoice Data
        ShipDate = m/d/y; get from Raw Invoice Data
        Item = 'SOMO Stock'; get from Raw Invoice Data
        ItemName = str, get from Raw Invoice Data
        ShipQty = int; get from Raw Invoice Data
        UnitPrice = price; get from Raw Invoice Data
        ShipAmount = price; get from Raw Invoice Data
        NewUnit$ = price; use lookup from price table
        NewShipAmount = ShipQty * NewUnit$
        """
        df = pd.DataFrame()
        length = self.raw_invoice.shape[0]
        df['Due Date'] = ['Net 15'] * length
        df['To Be emailed'] = [False] * length
        df['Print Later'] = [False] * length
        df['Dropship'] = self.raw_invoice['DropShipNo'].apply(lambda x: self.get_Dropship(x))
        df['Pivotal Account'] = self.raw_invoice['DropShipNo'].apply(lambda x: self.get_Pivotal_Account(x))
        df['DropShipNo'] = self.raw_invoice['DropShipNo']
        df['OrderID'] = self.raw_invoice['OrderID']
        df['ShipDate'] =self.raw_invoice['ShipDate']
        df['Item'] = ['SOMO Stock'] * length 
        df['ItemName'] = self.raw_invoice['ItemName']
        df['ShipQty'] = self.raw_invoice['ShipQty']
        df['UnitPrice'] = self.raw_invoice['UnitPrice']
        df['ShipAmount'] = self.raw_invoice['ShipAmount']
        df['NewUnit$'] = self.raw_invoice['Barcode'].apply(lambda x: self.get_New_Unit_Price(x))
        df['NewShipAmount'] = df['ShipQty'] * df['NewUnit$']
        df['ShipAmount'] = df['ShipAmount'].astype('float32')
        # df['NewShipAmount'] = df['NewShipAmount'].astype('float32')
        df['UPC'] = self.raw_invoice['Barcode']    
        df['Category'] = self.raw_invoice['Barcode'].apply(lambda x: self.get_Category_Name(x))  
        
        #Custom Formatting
        df['NewUnit$'] = df['NewUnit$'].round(2)
        df['NewShipAmount'] = df['NewShipAmount'].round(2)
        df['UPC'] = df['UPC'].astype(str).str.zfill(10)
    
        self.LensImport = df

    
    def generate_Shipping_Import(self):
        """
        Pivotal Account = /d/d/d/d/dA; get from lookup from Dropship
        Due Date = 'Net 15'
        To Be emailed = False
        Print Later = False
        Dropship = H00241-/d/d/d/d/d; Given is \d* from supplier (see DropshipNo), use lookups to get this
        OrderID = S/d/d/d/d/d/d; get from Raw Invoice Dat
        ShipDate = m/d/y; get from Raw Invoice Data
        Item = 'Shipping'
        ShipVia = str; get from Raw Invoice Data
        Freight = price; get from Raw Invoice Data

        Special Note: in Shipping Import tab only display where Freight is NOT zero.
        """
        df = pd.DataFrame()
        length = self.raw_invoice.shape[0]
        df['Pivotal Account'] = self.raw_invoice['DropShipNo'].apply(lambda x: self.get_Pivotal_Account(x))
        df['Due Date'] = ['Net 15'] * length
        df['To Be emailed'] = [False] * length
        df['Print Later'] = [False] * length
        df['Dropship'] = self.raw_invoice['DropShipNo'].apply(lambda x: self.get_Dropship(x))
        df['OrderID'] = self.raw_invoice['OrderID']
        df['ShipDate'] =self.raw_invoice['ShipDate']
        df['Item'] = ['Shipping']*length
        df['ShipVia'] = self.raw_invoice['ShipVia']
        df['Freight'] = self.raw_invoice['Freight']
        df['Freight'] = df['Freight'].round(2)

        df = df[df['Freight']!=0]
        self.ShippingImport = df

    def generate_Discount_Import(self):
        """
        Pivotal Account No. = /d/d/d/d/dA; get from lookup from Dropship
        Due Date = 'Net 15'
        To Be emailed = False
        Print Later = False
        Item = 'Stock Discount'
        Invoice # = TBD; Wants a unique invoice number - rough D{MMDDYY}seqint
        Description = '5% Legacy Discount'
        Invoice Date = last date of the current month
        ShipAmount = sum of all supplier prices from one customer; get from Lens Import
        NewShipAmount = sum of all retail prices from one customer; get from Lens Import
        Discount = NewShipAmount * .05
        Total Amount Owed = NewShipAmount - Discount

        Special Notes: only show non zeros, copy most from Lens Import
        """
        def generate_invoice_number():
            monthnum = str(self.now.month)
            if len(monthnum)<2:
                monthnum = '0' + monthnum
            yearnum  =str(self.now.year)[-2:]
            finalday = str(monthrange(self.now.year, self.now.month)[1])
            finaldigit = str(self.invoice_number_counter)
            while len(finaldigit)<4:
                finaldigit = '0' + finaldigit
            self.invoice_number_counter+=1
            return 'D'+monthnum+yearnum+finalday+finaldigit
        self.LensImport = self.LensImport[self.LensImport.DropShipNo.notna()]
        self.LensImport['DropShipNo'] = self.LensImport['DropShipNo'].apply(lambda x: str(int(x)) if isinstance(x, float) else x)
        #TODO prolly the spot to fix this
        discount_customer_list = self.customer_list[self.customer_list['Stock Lens 5% Discount']=='Yes']
        discount_customers = list(discount_customer_list['SuffixNum'].unique())
        all_customers = list(self.LensImport['DropShipNo'].unique())

        # all_customers = [str(int(c)) if isinstance(c, float) else c for c in all_customers]

        customers = [str(i) for i in all_customers if str(i) in discount_customers]
        length = len(customers)
        customers = [self.get_Pivotal_Account(i) for i in customers]
        df = pd.DataFrame()
        df['Pivotal Account No.'] = customers
        df['Due Date'] = ['Net 15'] * length
        df['To Be emailed'] = [False] * length
        df['Print Later'] = [False] * length
        df['Item'] = ['Stock Discount']*length
        df['Invoice #'] = [generate_invoice_number() for i in range(length)]
        df['Description'] = [r'5% Legacy Discount'] * length
        df['Invoice Date'] = [date(self.now.year, self.now.month, monthrange(self.now.year, self.now.month)[1])] * length
        df['ShipAmount'] = self.get_sum_ShipAmount(customers)
        df['NewShipAmount'] = self.get_sum_NewShipAmount(customers)
        df['Discount'] =  df['NewShipAmount']*.05
        df['Discount'] = df['Discount'].apply(lambda x: round(x, 2))
        df['Discount'] = df['Discount'].round(2)
        df['Total Amount Owed'] = round(df['NewShipAmount']-df['Discount'], 2)
        self.DiscountImport = df
    
    def generate_Summary_Sheet(self):
        """
        Pivotal #
        Freight = Freight
        ShipAmount = Total supplier amount for each customer
        NewShipAmount = Total retail amount for each customer
        Discount = Discount from Discount Import for relevant customers
        Total Charged = Freight + NewShipAmount - Discount
        """
        this_months_customers = list(self.LensImport['DropShipNo'].unique())
        all_customers = list(self.customer_list['SuffixNum'].unique())
        customers_who_didnt_purchase = [str(i) for i in all_customers if str(i) not in this_months_customers]
        customer_by_pivotal_account_no = [self.get_Pivotal_Account(i) for i in this_months_customers]
        df = pd.DataFrame()
        df['Pivotal #'] = customer_by_pivotal_account_no
        df['DropShipNo'] = this_months_customers
        df['Freight']  = self.get_sum_Freight(customer_by_pivotal_account_no)
        df['ShipAmount'] = self.get_sum_ShipAmount(customer_by_pivotal_account_no)
        df['NewShipAmount'] = self.get_sum_NewShipAmount(customer_by_pivotal_account_no)
        df['Discount'] = [self.DiscountImport[self.DiscountImport['Pivotal Account No.']==i]['Discount'].sum() for i in customer_by_pivotal_account_no]
        df['Total Charged'] = round(df['Freight'] + df['NewShipAmount'] - df['Discount'], 2)
        temp = pd.DataFrame()
        temp['Pivotal #'] = [self.get_Pivotal_Account(i) for i in customers_who_didnt_purchase]
        temp['DropShipNo'] = [int(i) for i in customers_who_didnt_purchase]
        df = df.append(temp)
        self.SummarySheet = df

    def generate_Summary_Overview(self):
        # df = pd.DataFrame()
        # df['Description'] = ['Total Pivotal Invoiced', 
        #                      'Retail Invoiced to Members', 
        #                      'Retail 5% Discount', 
        #                      'Net Retail Invoiced', 
        #                      '', 
        #                      'Net Profit']
        # df['Value'] = [self.SummarySheet['ShipAmount'].sum(),
        #                self.SummarySheet['NewShipAmount'].sum(),
        #                self.SummarySheet['Discount'].sum(),
        #                self.SummarySheet['NewShipAmount'].sum()-self.SummarySheet['Discount'].sum(),
        #                '',
        #                self.SummarySheet['NewShipAmount'].sum()-self.SummarySheet['Discount'].sum()-self.SummarySheet['ShipAmount'].sum()]
        # self.SummaryOverviewSheet = df
        #Updated Section
        df = pd.DataFrame()
        df['Description'] = ['Retail Lens Invoiced', 
                             'Shipping Costs', 
                             'Retail 5% Discount', 
                             'Total Invoiced', 
                             '',
                             '',
                             '', 
                             'Pivotal Invoiced',
                             'SOMO Disc', 
                             'Shipping Costs', 
                             'Tax',
                             'Total Pivotal Cost',
                             'Total Return Credits',
                             '',
                             '',
                             '',
                             'Net Profit']
        total_invoiced = self.SummarySheet['NewShipAmount'].sum() + self.SummarySheet['Freight'].sum() - self.SummarySheet['Discount'].sum()
        total_cost = self.SummarySheet['ShipAmount'].sum()+self.SOMO_Disc + self.SummarySheet['Freight'].sum()+self.TaxSheet['NewShipAmount'].sum()
        net_profit = total_invoiced-total_cost+self.LensReturnsCredits['NewShipAmount'].sum()
        df['Value'] = [self.SummarySheet['NewShipAmount'].sum(),
                       self.SummarySheet['Freight'].sum(),
                       self.SummarySheet['Discount'].sum(),
                       total_invoiced,
                       '',
                       '',
                       '',
                       self.SummarySheet['ShipAmount'].sum(),
                       self.SOMO_Disc, 
                       self.SummarySheet['Freight'].sum(),
                       self.TaxSheet['NewShipAmount'].sum(), #TODO Double check this
                       total_cost, #Total Pivotal Cost
                       (-1*self.LensReturnsCredits['NewShipAmount'].sum()),
                       '',
                       '',
                       '',
                       net_profit]
        self.SummaryOverviewSheet = df

    def generate_Lens_Returns_Credits(self):
        """
        Generates a sheet covering every negative sale indicating a return amount.
        """
        df = self.LensImport[self.LensImport['ShipQty']<0].copy()
        self.LensImport = self.LensImport[self.LensImport['ShipQty']>=0] # Removing negative LensImport ShipQty from Lens Import
        df['Item'] = ['Rtn Credit'] * df.shape[0]
        df['Positive ShipQ'] = df['ShipQty']*-1
        df['Positive New Ship Amount'] = df['NewShipAmount'] *-1
        self.LensReturnsCredits = df

    def remove_discount(self):
        """Removes Discount given for Pivotal denoted by DropShipNo 0 for use in the summary"""
        discount_row = self.raw_invoice[self.raw_invoice['DropShipNo']==0]
        self.SOMO_Disc = discount_row['TotalAmount'].sum()
        self.raw_invoice = self.raw_invoice[self.raw_invoice['DropShipNo']!=0]


    def archive_inputs(self):
        """Moves the files in Input to Archive upon completion of the script."""
        input_path = os.path.join(self.current_location, 'Input')
        archive_location = os.path.join(self.current_location, 'Archive')
        input_files = os.listdir(input_path)
        for input_file in input_files:
            old_location = os.path.join(input_path, input_file)
            new_location = os.path.join(archive_location, input_file)
            os.rename(old_location, new_location)

    def divide_Lens_Import(self):
        """Takes the Lens Import sheet and breaks it into chunks of no more than 5000 rows so that QuickBooks can handle it."""
        number_of_sheets = self.LensImport.shape[0]//5000 + 1
        return [[f'Lens Import {i+1}', self.LensImport[i*5000:(i+1)*5000]] for i in range(number_of_sheets)]

    def generate_Tax_Sheet(self):
        temp = self.raw_invoice[self.raw_invoice['Tax']!=0].copy()
        df = pd.DataFrame()
        length = temp.shape[0]
        df['Due Date'] = ['Net 15'] * length
        df['To Be emailed'] = [False] * length
        df['Print Later'] = [False] * length
        df['Dropship'] = temp['DropShipNo'].values
        df['Dropship'] = df['Dropship'].apply(lambda x: self.get_Dropship(x))
        df['Pivotal Account'] = temp['DropShipNo'].values
        df['Pivotal Account'] = df['Pivotal Account'].apply(lambda x: self.get_Pivotal_Account(x))
        df['DropShipNo'] = temp['DropShipNo'].values
        df['OrderID'] = temp['OrderID'].values
        df['ShipDate'] =temp['ShipDate'].values
        df['Item'] = ['SOMO Stock'] * length 
        df['ItemName'] = ['Tax'] * length
        df['NewShipAmount'] = temp['Tax'].values
        df['NewShipAmount'] = df['NewShipAmount'].round(2)
        self.TaxSheet = df

    def generate_csv(self):
        """Full process of generating each sheet then writing it to an excel file."""
        self.remove_discount()
        self.generate_Lens_Import()
        self.generate_Shipping_Import()
        self.generate_Discount_Import()
        self.generate_Lens_Returns_Credits()
        self.generate_Tax_Sheet()
        self.generate_Summary_Sheet()
        self.generate_Summary_Overview()
        output_name = self.create_output_name()
        writer = pd.ExcelWriter(os.path.join(self.save_location, output_name), engine='xlsxwriter')
        list_of_Lens_Import_chunks = self.divide_Lens_Import()
        if len(list_of_Lens_Import_chunks)>1:
            for chunk in list_of_Lens_Import_chunks:
                chunk[1].to_excel(writer, index=False, sheet_name=chunk[0])
        else:
            self.LensImport.to_excel(writer, index=False, sheet_name='Lens Import')
        self.TaxSheet.to_excel(writer, index=False, sheet_name='Taxes')
        self.ShippingImport.to_excel(writer, index=False, sheet_name='Shipping Import')
        self.DiscountImport.to_excel(writer, index=False, sheet_name='Discount Import')
        self.LensReturnsCredits.to_excel(writer, index=False, sheet_name='Lens Returns Credits')
        self.SummarySheet.to_excel(writer, index=False, sheet_name='Summary Details')
        self.SummaryOverviewSheet.to_excel(writer, index=False, sheet_name='Summary Overview')
        writer.save()
        self.archive_inputs()


def main(_path = ''):
    if not _path:
        _path = os.path.dirname(sys.executable)#os.path.dirname(os.path.abspath(__file__))
 # Currnet loc
    def error_popup(msg):
        """Super simple pop-up to indicate an error has occured."""
        popup = tk.Tk()
        popup.wm_title("!")
        label = tk.Label(popup, text=msg)
        label.pack(side="top", fill="x", pady=10)
        B1 = tk.Button(popup, text="Okay", command = popup.destroy)
        B1.pack()
        popup.mainloop()
 
    #First, update the reference sheet
    #umr = UpdateMasterReference()
    #umr.add_new_customers_to_MasterReference()
    umr = MasterReferenceUpdater(_path) 
    umr.RUN()

    #Next, run the sanity check
    sc = SanityCheck()
    passed_checks = sc.run_check()

    #Finally generate the invoice plus summaries
    if passed_checks:
        rg = ReportGenerator()
        if rg.invoice_Found:
            rg.generate_csv()
        else:
            error_popup('Failed to run. No invoice found in input folder.')
    else:
        error_popup('Failed to run. One or more tests failed. See REFERENCE_ERROR for details.')
if __name__ == '__main__':
    print('Running')
    try:
        main() #TODO remove
    except:
        print(traceback.format_exc())
        with open(r'.\TRACEBACK.txt', 'w') as f:
            f.write(f'{traceback.format_exc()}')


