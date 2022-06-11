from csv import writer
from sys import argv, exit 
import os
from datetime import date 
import pandas as pd 
import re
from xlsxwriter.utility import xl_rowcol_to_cell

def get_sales_csv():

    #Check command line paramters were provided
    if len(argv) >= 2:
        sales_csv = argv[1]

        #check whether csv path is an existing file
        if os.path.isfile(sales_csv):
            return sales_csv
        
        else:
            print('Error: CSV file does not exist')
            exit('Script execution aborted')

    else:
        print('Error: No CSV file path provided')
        exit('Script execution aborted')

def get_order_dir(sales_csv):
    
    # Get directory path of sales data CSV file
    sales_dir = os.path.dirname(sales_csv)

    # Determine orders directory name Orders_YYYY_MM_DD
    todays_date = date.today().isoformat()
    order_dir_name = 'Orders_' + todays_date 
    

    # Build full path of orders directory 
    order_dir = os.path.join(sales_dir, order_dir_name)

    
    # Make orders directory if it does not already exist
    if not os.path.exists(order_dir):
        os.makedirs(order_dir)

    return order_dir

def split_sales_into_orders(sales_csv, order_dir):

    # Read data from the sales data CSV into dataframe 
    sales_df = pd.read_csv(sales_csv)   

    # Insert new column for total price
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE']) 

    # Drop unwanted columns
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'COUNTRY', 'POSTAL CODE'], inplace=True)

    for order_id, order_df in sales_df.groupby('ORDER ID'):

        # Drop order id column 
        order_df.drop(columns=['ORDER ID'], inplace=True)

        # Sort the order by item number 
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        #Add grand total row at bottom
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        # Determine the save path of the order file
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)

        # Save the order information to an excel spreadsheet
        sheet_name = 'Order #' + str(order_id)        
        #order_df.to_excel(order_file_path, index=False, sheet_name=sheet_name)
        
        # Formatting output
        writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name] 

        money_fmt = workbook.add_format ({'num_format': '$#,##0.##'})
        worksheet.set_column('A:A', 11)
        worksheet.set_column('B:B', 13) 
        worksheet.set_column('C:E', 15)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 30)
        worksheet.set_column('F:G', 12, money_fmt)

        writer.save()

sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv) 
split_sales_into_orders(sales_csv, order_dir)