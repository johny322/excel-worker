import datetime
import io
import os
import re
from pathlib import Path
from typing import Optional, Union, List

import pandas as pd

from excel_worker.types import BeautifulType, OptionsEngineKwargs, DropDuplicatesSettings, DEFAULT_SHEET_NAME

DATETIME_FORMAT = '%d_%m_%Y_%H_%M'


class Excel:
    def __init__(self):
        self.data: Optional[dict] = None
        self.main_key: Optional[str] = None
        self.writer: Optional[pd.ExcelWriter] = None
        self.sheet_name = DEFAULT_SHEET_NAME
        self.datetime_format = DATETIME_FORMAT

    def setup(self, main_key: str, data: dict) -> None:
        self.main_key = main_key
        self.data = data

    def cleanup(self) -> None:
        self.data = {
            self.main_key: []
        }

    def set_data(self, data) -> None:
        self.data = data

    def set_main_key(self, main_key) -> None:
        self.main_key = main_key

    def get_df(self) -> pd.DataFrame:
        """
        DataFrame из self.data
        """
        self.check_data()
        df = pd.DataFrame(self.data)
        return df

    def get_df_bytes(self, sheet_name: str = None, beautiful_type=BeautifulType.max_length) -> bytes:
        if sheet_name is None:
            sheet_name = self.sheet_name

        self.check_data()
        df = pd.DataFrame(self.data)
        towrite = io.BytesIO()

        writer = pd.ExcelWriter(towrite)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        self.__beautiful_df(df=df, writer=writer, beautiful_type=beautiful_type, sheet_name=sheet_name)

        writer.save()

        towrite.seek(0)
        return towrite.read()

    def get_df_bytes_io(self, sheet_name: str = None, beautiful_type=BeautifulType.max_length) -> io.BytesIO:
        if sheet_name is None:
            sheet_name = self.sheet_name

        self.check_data()
        df = pd.DataFrame(self.data)
        towrite = io.BytesIO()

        writer = pd.ExcelWriter(towrite)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        self.__beautiful_df(df=df, writer=writer, beautiful_type=beautiful_type, sheet_name=sheet_name)

        writer.save()

        towrite.seek(0)
        return towrite

    def _get_now_file_name(self, path: Path, datetime_format: str = None) -> Path:
        if not datetime_format:
            datetime_format = self.datetime_format
        now = datetime.datetime.now().strftime(datetime_format)
        file_parent = path.parent
        file_stem = path.stem
        file_suffix = path.suffix

        result_path = Path(file_parent) / f'{file_stem}_{now}{file_suffix}'
        return result_path

    @staticmethod
    def _handle_path(path: Union[Path, str]):
        if not isinstance(path, Path):
            path = Path(path)
        return path.absolute()

    def create_writer(self, path: Optional[Path, str]) -> None:
        """
        Создает writer для работы с листами экселя

        :param path: путь для сохранения будущего файла
        """
        self.writer = pd.ExcelWriter(path)

    def writer_save(self) -> None:
        """
        Сохраняет файл с помощью writer по пути path из метода self.create_writer(path)
        """
        self.writer.save()

    def add_to_sheet(self, sheet_name: str) -> None:
        """
        Создает df из self.data и добавляет его на лист sheet_name.
        После добавления надо задать начальное значение self.data для сохранения на другие листы.

        :param sheet_name: имя листа в экселе
        :return:
        """
        self.check_data()
        df = pd.DataFrame(self.data)
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)

    @staticmethod
    def create_hyperlink(url: str, text: str) -> str:
        """
        форматирование строки для создания гипперсылки в файле excel
        """
        return f'=HYPERLINK("{url}", "{text}")'

    def add_several_values(self, name: str, values: list) -> None:
        """
        Добавляет значения из массива в столбцы с названием и порядковым номером.

        Пример:

        add_several_values('name', ['value1', 'value2', 'value3'])
        Добавится в столбцы name1 - value1, name2 - value2, name3 - value3, если их не было, они создадутся.

        :param name: имя столбца для добавления
        :param values: массив значений
        """
        keys = list(self.data.keys())
        p_key = []
        for k in keys:
            pattern = name + r'\d{1,}'
            match = re.match(pattern, k)
            if match:
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

    def add_key_value(self, key: Union[str, int, float, bool, None], value: Union[str, int, float, bool, None]) -> None:
        """
        Добавляет значение в колонку по ключу.

        :param key: ключ для колонки
        :param value: значение для добавлния
        """
        keys = list(self.data.keys())
        if key in keys:
            difference = len(self.data[self.main_key]) - len(self.data[key])
            if difference > 1:
                self.data[key].extend([None] * (difference - 1))
            # if difference < 0:
            #     self.add_several_values([value], key)
            # if difference - 1 < 0 and key != self.main_key:
            #     self.add_several_values([value], key)
            self.data[key].append(value)
        else:
            self.data[key] = [None] * (len(self.data[self.main_key]) - 1)
            self.data[key].append(value)

    def check_data(self) -> None:
        """
        Проверяет данные, чтобы в каждом столбце было одинаковое количество значений, если не так, то добавит None.

        """
        keys = list(self.data.keys())
        main_len = len(self.data[self.main_key])
        for key in keys:
            key_len = len(self.data[key])
            diff = main_len - key_len
            if diff > 0:
                self.data[key].extend([None] * diff)
            if diff < 0:
                self.data[key] = self.data[key][:diff]

    def save_excel(self, path: Union[Path, str], beautiful_type=BeautifulType.max_length, now_date=True,
                   ignore_urls=False, sheet_name: str = None, df: pd.DataFrame = None) -> Path:
        """
        Сохраняет файл по пути.

        """
        path = self._handle_path(path)
        if now_date:
            path = self._get_now_file_name(path)
        self.check_data()
        if df is None:
            df = pd.DataFrame(self.data)

        if sheet_name is None:
            sheet_name = self.sheet_name
        engine_kwargs = OptionsEngineKwargs(ignore_urls=ignore_urls).engine_kwargs

        writer = pd.ExcelWriter(path, engine_kwargs=engine_kwargs)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        self.__beautiful_df(df=df, writer=writer, beautiful_type=beautiful_type, sheet_name=sheet_name)

        writer.save()
        return path

    @staticmethod
    def __beautiful_df(df: pd.DataFrame, writer: pd.ExcelWriter, beautiful_type: BeautifulType,
                       sheet_name: str) -> None:
        if beautiful_type != BeautifulType.no_beautiful:
            for column in df:
                if beautiful_type == BeautifulType.max_length:
                    column_width = max(df[column].astype(str).map(len).max(), len(column))
                elif beautiful_type == BeautifulType.column_length:
                    column_width = len(column) + 5
                else:
                    column_width = 40
                col_idx = df.columns.get_loc(column)
                writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width)

    def beautiful_excel(self, open_path, save_path, beautiful_type=BeautifulType.max_length, sheet_name: str = None):
        """

        """
        if sheet_name is None:
            sheet_name = self.sheet_name
        df = pd.read_excel(open_path)
        writer = pd.ExcelWriter(save_path)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        self.__beautiful_df(df, writer, beautiful_type, sheet_name)

        writer.save()
        return os.path.abspath(save_path)

    @staticmethod
    def concat_files(files: List[Union[Path, str]], final_name: Union[Path, str],
                     drop_duplicates_settings: DropDuplicatesSettings = DropDuplicatesSettings,
                     beautiful_type=BeautifulType.max_length,
                     ignore_urls: bool = False,
                     sheet_name: str = None
                     ) -> Path:
        con_files = []
        for file in files:
            con_files.append(pd.read_excel(file))
        final_name = Excel._handle_path(final_name)
        if sheet_name is None:
            sheet_name = DEFAULT_SHEET_NAME

        df = pd.concat(con_files, ignore_index=True)
        if drop_duplicates_settings.drop_duplicates:
            df = df.drop_duplicates(subset=drop_duplicates_settings.subset, keep=drop_duplicates_settings.keep)
        engine_kwargs = OptionsEngineKwargs(ignore_urls=ignore_urls).engine_kwargs

        writer = pd.ExcelWriter(final_name, engine_kwargs=engine_kwargs)

        df.to_excel(writer, sheet_name=sheet_name, index=False)

        Excel.__beautiful_df(df, writer, beautiful_type, sheet_name)

        writer.save()

        return final_name

    @staticmethod
    def drop_duplicates(read_path: Union[Path, str], save_path: Union[Path, str] = None,
                        drop_duplicates_settings: DropDuplicatesSettings = DropDuplicatesSettings,
                        ignore_urls=False) -> Path:

        read_path = Excel._handle_path(read_path)
        save_path = Excel._handle_path(save_path)
        if not save_path:
            file_parent = read_path.parent
            file_stem = read_path.stem
            file_suffix = read_path.suffix

            save_path = Path(file_parent) / f'{file_stem}_no_dub{file_suffix}'

        engine_kwargs = OptionsEngineKwargs(ignore_urls=ignore_urls).engine_kwargs
        writer = pd.ExcelWriter(save_path, engine_kwargs=engine_kwargs)

        df = pd.read_excel(read_path)
        df = df.drop_duplicates(subset=drop_duplicates_settings.subset, keep=drop_duplicates_settings.keep)
        df.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)
        writer.save()

        return save_path
