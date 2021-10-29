import os

import pandas as pd


class Exel:
    # def __init__(self, data, main_key):
    #     self.data = data
    #     self.main_key = main_key

    def change_data(self, data):
        self.data = data

    def change_main_key(self, main_key):
        self.main_key = main_key

    def add_several_values(self, values, name: str):
        keys = list(self.data.keys())
        p_key = []
        for k in keys:
            if name in k:
                p_key.append(k)
        if len(values) >= len(p_key):
            for i in range(len(p_key) + 1, len(values) + 1):
                self.data[f'{name}{i}'] = [None] * (len(self.data[self.main_key]) - 1)
            for i in range(len(values)):
                self.data[f'{name}{i + 1}'].append(values[i])
        else:
            for i in range(len(values)):
                if f'{name}{i + 1}' not in self.data:
                    self.data[f'{name}{i + 1}'] = [None] * (len(self.data[self.main_key]) - 1)
                self.data[f'{name}{i + 1}'].append(values[i])
            for i in range(len(values) + 1, len(p_key) + 1):
                if f'{name}{i}' not in self.data:
                    self.data[f'{name}{i}'] = [None] * (len(self.data[self.main_key]) - 1)
                self.data[f'{name}{i}'].append(None)

    def add_key_value(self, key, value):
        keys = list(self.data.keys())
        if key in keys:
            difference = len(self.data[self.main_key]) - len(self.data[key])
            if difference > 1:
                self.data[key].extend([None] * (difference - 1))
            self.data[key].append(value)
        else:
            self.data[key] = [None] * (len(self.data[self.main_key]) - 1)
            self.data[key].append(value)

    def check_data(self):
        keys = list(self.data.keys())
        main_len = len(self.data[self.main_key])
        for key in keys:
            key_len = len(self.data[key])
            diff = main_len - key_len
            if diff > 0:
                self.data[key].extend([None] * diff)

    def write_exel(self, path):
        self.check_data()
        df = pd.DataFrame(self.data)
        df.to_excel(path, index=False)
        return os.path.abspath(path)
        # writer = pd.ExcelWriter(path)
        # df.to_excel(writer, sheet_name='1', index=False)
        # for column in df:
        #     column_width = max(df[column].astype(str).map(len).max(), len(column))
        #     col_idx = df.columns.get_loc(column)
        #     writer.sheets['1'].set_column(col_idx, col_idx, column_width)
        # writer.save()

    @staticmethod
    def beautiful_exel(open_path, save_path, sheet_name='Sheet1', bt_type='max'):
        df = pd.read_excel(open_path)
        writer = pd.ExcelWriter(save_path)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        for column in df:
            if bt_type == 'max':
                column_width = max(df[column].astype(str).map(len).max(), len(column))
            else:
                column_width = len(column) + 5
            col_idx = df.columns.get_loc(column)
            writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width)
        writer.save()
        return os.path.abspath(save_path)

    @staticmethod
    def concat_files(files: list, final_name: str, drop_duplicates: bool = False, subset=''):
        con_files = []
        for file in files:
            con_files.append(pd.read_excel(file))
        df = pd.concat(con_files, ignore_index=True)
        if drop_duplicates:
            df = df.drop_duplicates(subset=subset, keep='first')
        df.to_excel(final_name, index=False)
        return os.path.abspath(final_name)
