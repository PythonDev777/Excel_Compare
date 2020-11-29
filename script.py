import os
import pandas as pd


class ExcelCompare:
    def __init__(self):
        self.master = f"{os.path.dirname(os.path.abspath(__file__))}/FRS  MASTER FILE WITH HEARING DATES C & I COUNTY COURT.xlsx"
        self.format_file = f"{os.path.dirname(os.path.abspath(__file__))}/11_20_2020 BRWD FJ RAW.xlsx"

    def query_data_and_update_contents_excel_file(self):
        file = pd.read_excel(self.format_file)
        master = pd.read_excel(self.master)
        file_case_num = file[['Case #']]
        df = ''
        for index, case_num in file_case_num.iterrows():
            case_number = case_num['Case #']
            update_cred = master.loc[master['Case Number'] == case_number]
            file_row_content = file.loc[file['Case #'] == case_number]
            date = pd.to_datetime(file_row_content["Date"]).dt.strftime("%m-%d-%Y")
            file.loc[index, 'Date'] = date.values[0]
            if update_cred.empty:
                continue
            print('Updating Case Number ....' + case_number)
            file.loc[index, 'Amount $'] = update_cred['Code'].values[0]
            file.loc[index, 'Case #'] = case_number
            file.loc[index, 'Mailing Address'] = update_cred['Address'].values[0]
            file.loc[index, 'City'] = update_cred['City'].values[0]
            file.loc[index, 'ST'] = update_cred['ST'].values[0]
            file.loc[index, 'Zip'] = update_cred['Zip Code'].values[0]
            df = file.sort_values(['Mailing Address', 'City', 'ST', 'Zip'], na_position='last')

        with pd.ExcelWriter('test2.xlsx', mode='w') as writer:
            df.to_excel(writer)


if __name__ == '__main__':
    main = ExcelCompare()
    main.query_data_and_update_contents_excel_file()


