import xlsxwriter
import os
import sys 

from xlsxwriter.workbook  import Workbook
from xlsxwriter.worksheet import Worksheet
from typing               import Tuple

def create_new_report_directory(dirName:str) -> None:
    os.mkdir(f'results/{dirName}')

def create_xlsx_table(worksheet_name:str) -> Tuple[Workbook, Worksheet]:
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


def write_header_info(worksheet: Worksheet, cursor_row:int, data:dict) -> Tuple[int, list]:
    """
    Записывает основную информацию о генерируемой вызгрузке 
    """

    worksheet.write(cursor_row, 0, data['test_name'])
    worksheet.write(cursor_row, 1, data['gen_date'])
    working_col = 3
    start_row = cursor_row
    start_col = working_col
    
    for q_info in data['test_question_data']:
        worksheet.write(cursor_row, working_col, data['test_question_data'][q_info]['q_type'])
        working_col+=1 
    cursor_row += 2
    

    return cursor_row, [1,2,3]


def write_munipal_info(worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[dict, int]:
    """
    Записывает информацию о муниципалитете 
    """

    worksheet.write(cursor_row, 0, f'Муниципалитет {data[0]}')
    cursor_row+=1
    
    return {"mun_cells_for_formula": ['A1','B1','C1','D1']}, cursor_row

def write_school_info():
    pass


def write_result_info():
    pass


def write_formulas():
    pass


if __name__ == '__main__':
    sys.exit()
