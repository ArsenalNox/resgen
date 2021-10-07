import xlsxwriter
import os

from xlsxwriter.worksheet import Worksheet


def create_new_report_directory(dirName:str)->None:
    os.mkdir(f'results/{dirName}')

def create_xlsx_table(worksheet_name:str):
    """
    Созадёт файл таблицы 
    возвращает таблицу и лист 
    Формат таблицы:
                                             | Тип вопроса 1 | ....
    Муниципалитет | Общий процент выполнения | % вопроса 1 | % 2 ...
    Код школы | % выполнения школы | % вопроса 1 | ...
    """
    print(worksheet_name+'.xlsx')
    xlsxWorkBook  = xlsxwriter.Workbook(worksheet_name+'.xlsx')
    xlsxWorkSheet = xlsxWorkBook.add_worksheet()

    return xlsxWorkBook, xlsxWorkSheet


def write_header_info(worksheet: Worksheet, cursor_row:int, data:dict)->int:
    """
    Записывает основную информацию об тесте в 
    """

    return cursor_row

def write_munipal_info():
    pass 


def write_school_info():
    pass


def write_result_info():
    pass


def write_formulas():
    pass


if __name__ == '__main__':
    create_new_report_directory('test')
