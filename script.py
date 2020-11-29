import os
import pandas as pd


class ExcelCompare:
    def __init__(self):
        self.master = f"{os.path.dirname(os.path.abspath(__file__))}/FRS  MASTER FILE WITH HEARING DATES C & I COUNTY COURT (1) (1).xlsx"
        self.format_file = f"{os.path.dirname(os.path.abspath(__file__))}/11_20_2020 BRWD FJ RAW.xlsx"
        self.master_case_numbers = []

    def query_data_and_update_contents_excel_file(self):
        file = pd.read_excel(self.format_file)
        master = pd.read_excel(self.master)
        file_case_num = file[['Case #']]
        df = pd.DataFrame(columns=file.columns)
        for index, case_num in file_case_num.iterrows():
            case_number = case_num['Case #']
            update_cred = master.loc[master['Case Number'] == case_number]
            if update_cred.empty:
                continue
            print('Updating Case Number ....' + case_number)
            file_row_content = file.loc[file['Case #'] == case_number]
            df.loc[index, 'Unnamed: 0'] = file_row_content['Unnamed: 0'].values[0]
            df.loc[index, 'Unnamed: 0.1'] = file_row_content['Unnamed: 0.1'].values[0]
            df.loc[index, 'Unnamed: 0.1.1'] = file_row_content['Unnamed: 0.1.1'].values[0]
            df.loc[index, 'Unnamed: 0.1.1.1'] = file_row_content['Unnamed: 0.1.1.1'].values[0]
            df.loc[index, 'ClerkFileNumber'] = file_row_content['ClerkFileNumber'].values[0]
            df.loc[index, 'Date'] = file_row_content['Date'].values[0]
            df.loc[index, 'Plaintiff'] = file_row_content['Plaintiff'].values[0]
            df.loc[index, 'Type'] = file_row_content['Type'].values[0]
            df.loc[index, 'Amount $'] = update_cred['Code'].values[0]
            df.loc[index, 'Case #'] = case_number
            df.loc[index, 'Def First'] = file_row_content['Def First'].values[0]
            df.loc[index, 'Def MI'] = file_row_content['Def MI'].values[0]
            df.loc[index, 'Def Last'] = file_row_content['Def Last'].values[0]
            df.loc[index, 'Mailing Address'] = update_cred['Address'].values[0]
            df.loc[index, 'City'] = update_cred['City'].values[0]
            df.loc[index, 'ST'] = update_cred['ST'].values[0]
            df.loc[index, 'Zip'] = update_cred['Zip Code'].values[0]

        with pd.ExcelWriter(self.format_file, mode='w') as writer:
            file.to_excel(writer)


if __name__ == '__main__':
    main = ExcelCompare()
    main.query_data_and_update_contents_excel_file()


