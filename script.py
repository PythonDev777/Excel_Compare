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
        for index, case_num in file_case_num.iterrows():
            case_number = case_num['Case #']
            update_cred = master.loc[master['Case Number'] == case_number]
            if update_cred.empty:
                continue
            print('Updating Case Number ....' + case_number)
            address = update_cred['Address'].values[0]
            city = update_cred['City'].values[0]
            zip_code = update_cred['Zip Code'].values[0]
            state = update_cred['ST'].values[0]
            amount = update_cred['Code'].values[0]
            file.loc[index, 'Mailing Address'] = address
            file.loc[index, 'City'] = city
            file.loc[index, 'ST'] = state
            file.loc[index, 'Zip'] = zip_code
            file.loc[index, 'Amount $'] = amount

        with pd.ExcelWriter(self.format_file, mode='w') as writer:
            file.to_excel(writer)


if __name__ == '__main__':
    main = ExcelCompare()
    main.query_data_and_update_contents_excel_file()


