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


file_paths = ["Rocznik_2014__GR.xls", "Rocznik_2015__GR.xls", "Rocznik_2016__GR.xls", "Rocznik_2017_GR.xls", "Rocznik_2018_GR.xls", "Rocznik_2019_GR.xls",
              "Rocznik_2020_GR.xls", "Rocznik_2021_GR.xls", "Rocznik_2022_GR.xls"]


def create_helper_dicts():

    files_to_tab_names = dict()
    headers = dict()
    for ind, path in enumerate(file_paths):
        files_to_tab_names[path] = dict()
        headers[path] = dict()

        if ind <= 3:
            files_to_tab_names[path]["wartosci_akcji"] = "Tab 13"
            files_to_tab_names[path]["najwyzsze_sesyjne_obroty"] = "Tab 14"
            files_to_tab_names[path]["stopy_zwrotu"] = "Tab 15"
            files_to_tab_names[path]["najwyzsze_stopy_zwrotu"] = "Tab 16"
        else:
            files_to_tab_names[path]["wartosci_akcji"] = "Tab 8"
            files_to_tab_names[path]["najwyzsze_sesyjne_obroty"] = "Tab 9"
            files_to_tab_names[path]["stopy_zwrotu"] = "Tab 10"
            files_to_tab_names[path]["najwyzsze_stopy_zwrotu"] = "Tab 11"
        headers[path]["wartosci_akcji"] = [3]
    
    return files_to_tab_names, headers


def read_excel_files(path):
    excel_data = {}
    files_to_tab_names, headers = create_helper_dicts()
    for file in os.listdir(path):
        if file.endswith(".xls"):
            year = file.split('_')[1]  # Extracting year from file name
            excel_data[year] = {}
            xls = pd.ExcelFile(os.path.join(path, file))
            for sheet_name in xls.sheet_names:

              if sheet_name == files_to_tab_names[file]["wartosci_akcji"]:
                excel_data[year]["wartosci_akcji"] = pd.read_excel(xls, sheet_name=sheet_name, header=headers[file]["wartosci_akcji"], usecols = range(11))

    print(excel_data.keys())
    return excel_data
excel_data = read_excel_files("./Market_Value/")


def separate_data_by_company(excel_data):

    company_dfs = {}
    for year, sheets in excel_data.items():
        for sheet_name, df in sheets.items():
            company_col = None
            if 'Lp./ No' in df.columns:
              df.drop(columns = ['Lp./ No'])
            if 'Unnamed: 9' in df.columns: #delete empty/irrelevent columns
              df.drop(columns = ['Unnamed: 9'])
            for col in df.columns: #Company column named differently in several files
                if 'Spółka/ Company' in col:
                    company_col = col
                    break
                elif 'Akcje/ Shares' in col:
                    company_col = col
                    break
                elif 'Spółka/Company' in col:
                    company_col = col
                    break
                elif 'Akcje/Shares' in col:
                    company_col = col
                    break
                elif 'Spółka / Company' in col:
                    company_col = col
                    break
                elif 'Akcje / Shares' in col:
                    company_col = col
                    break

            if company_col is None:
                raise ValueError(f"Company column not found in the DataFrame, sheet_name={sheet_name}, year = {year}")

            for company, group in df.groupby(by=company_col):
                if company not in company_dfs:
                    # Initialize DataFrame with columns corresponding to each year
                    company_dfs[company] = pd.DataFrame(columns=excel_data.keys())
                # Add data for the company and year, transposing the group
                company_dfs[company][year] = group.drop(columns=[company_col]).T
    for company, df in company_dfs.items():
        company_dfs[company] = df.sort_index(axis=1)
    return company_dfs
