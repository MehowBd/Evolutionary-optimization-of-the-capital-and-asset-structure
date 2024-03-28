import openpyxl
import pandas as pd
import glob
import os
import csv

def import_companies_balance_sheet():
    #get path to files in folder
    path_to_files = './Companies_Balance_sheet/*.xlsx'
    excel_files = glob.glob(path_to_files)
    data_dict = dict()

    for file_path in excel_files:
        sheet_name = 'YS'

        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the worksheet by name
        sheet = workbook[sheet_name]   

        cell_ranges = ['C1:C26', 'X1:AB26']

        # Initialize an empty list to store the cell values

        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the worksheet by name
        sheet = workbook[sheet_name]

        cell_values = []
        for row in sheet[cell_ranges[1]]:
            row_values = [cell.value for cell in row]
            if row_values:
                cell_values.append(row_values)

        # Extract index values
        index_values = []
        for row in sheet[cell_ranges[0]]:
            for cell in row:
                if cell.value is not None:
                    index_values.append(cell.value)

        # Convert the list of lists to a DataFrame
        df = pd.DataFrame(cell_values, index=index_values, columns=['2017', '2018', '2019', '2020', '2021'])
        
        # Add Dataframe for scructure into dataframe of all structures
        filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
        data_dict[filename_without_extension] = df

    csv_file = './Data/companies_balance_sheet.csv'
    with open(csv_file, 'w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=list(data_dict.keys()))
        
        # Write the header
        writer.writeheader()
        
        # Write the data
        writer.writerow(data_dict)
    return data_dict