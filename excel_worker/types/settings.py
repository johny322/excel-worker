from dataclasses import dataclass, field
from enum import Enum
from typing import Literal

DEFAULT_SHEET_NAME = 'Sheet1'


class OptionsEngineKwargs:
    def __init__(self, ignore_urls=False):
        self._engine_kwargs = {
            'options': {}
        }
        if ignore_urls:
            self._engine_kwargs['options']['strings_to_urls'] = False

    @property
    def engine_kwargs(self):
        return self._engine_kwargs


class DropDuplicatesOptions:
    def __init__(self, drop_duplicates: bool = True, subset: str = '', keep: Literal['first', 'last', False] = 'first'):
        pass


@dataclass
class DropDuplicatesSettings:
    drop_duplicates: bool = False
    subset: str = ''  # название столбца
    keep: Literal['first', 'last', False] = 'first'


class BeautifulType(str, Enum):
    """
    типы форматирования файла по ширине столбцов
    """
    no_beautiful = 'none'  # не форматировать файл
    column_length = 'column_length'  # форматировать по ширине названия столбца
    max_length = 'max_length'  # форматировать столбец по максимальной ширине значения (включая название столбца)


class Settings:

    @property
    def kwargs(self):
        return self.__dict__.copy()


@dataclass
class BeautifulSettings(Settings):
    pass


@dataclass
class ExcelSettings:
    pass
