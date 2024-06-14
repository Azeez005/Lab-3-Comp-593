import sys
import os
from datetime import date
import pandas as pd


def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)
    print(sales_csv)
    print(orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    if len(sys.argv) < 2:
        print("Error, Missing Csv file path")
        sys.exit(1)

    # Check whether provide parameter is valid path of file
    csv_path = sys.argv[1]
    if not os.path.isfile(sys.argv[1]):
        print('Error: Invaild CSV file path')
        sys.exit(1)

    return sys.argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_csv_path = os.path.abspath(sales_csv) 
    sales_csv_dir = os.path.dirname(sales_csv_path)

    # Determine the name and path of the directory to hold the order data files
    current_date = date.today().isoformat()
    orders_folder = f"Orders_{current_date}"
    orders_dir = os.path.join(sales_csv_dir, orders_folder)

    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir 

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)

    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df['TOTAL PRICE'] = sales_df['ITEM QUANTITY'] * sales_df['PRICE EACH']

    # Remove columns from the DataFrame that are not needed
    sales_df.drop(['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], axis=1, inplace=True)

    # Group the rows in the DataFrame by order ID
    grouped_orders = sales_df.groupby('ORDER ID')

    for order_id, order_data in grouped_orders:
        # Remove the "ORDER ID" column
        order_data = order_data.drop('ORDER ID', axis=1)

        # Sort the items by item number
        order_data = order_data.sort_values(by='ITEM NUMBER')

        # Append a "GRAND TOTAL" row
        grand_total_row = pd.DataFrame([['GRAND TOTAL', '', '', '', '', '', '', order_data['TOTAL PRICE'].sum()]],columns=order_data.columns)
        order_data = pd.concat([order_data, grand_total_row])

        # Determine the file name and full path of the Excel sheet
        excel_file_name = f"{order_id}.xlsx"
        excel_file_path = os.path.join(orders_dir, excel_file_name)

        # Export the data to an Excel sheet
        order_data.to_excel(excel_file_path, index=False)

        # Format the Excel sheet (you can add formatting as per your requirements)
        # For example, you can set the column width
        writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
        order_data.to_excel(writer, sheet_name='Sheet1', index=False)
        excelsheet = writer.sheets['Sheet1']
        excelsheet.set_column('A:H', 15)  # Adjust column width as needed
        writer.save()
    pass

if __name__ == '__main__':
    main()