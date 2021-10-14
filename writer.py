import xlsxwriter
import os
import sys
from xlsxwriter.utility   import xl_rowcol_to_cell 

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


def write_header_info(worksheet: Worksheet, cursor_row:int, data:dict) -> Tuple[int, list, int]:
    """
    Записывает основную информацию о генерируемой вызгрузке 
    """

    worksheet.write(cursor_row, 0, data['test_name'])
    worksheet.write(cursor_row, 1, data['gen_date'])
    working_col = 3

    start_row = cursor_row
    start_col = working_col
    cell = xl_rowcol_to_cell(start_row, start_col)

    q_num = 0
    for q_info in data['test_question_data']:
        worksheet.write(cursor_row, working_col, data['test_question_data'][q_info]['q_type'])
        working_col+=1 
        q_num+=1

    cursor_row += 2

    return cursor_row, [cell], q_num


def write_munipal_info(worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[dict, int]:
    """
    Записывает информацию о муниципалитете 
    """

    worksheet.write(cursor_row, 0, f'Муниципалитет {data["mcode"]}')

    cell = xl_rowcol_to_cell(cursor_row, 1)
    cursor_row+=1
        
    return {"mun_cells_formula_start": [cell]}, cursor_row

def write_school_info(workbook: Workbook ,worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[dict, int]:
    """
    Записывает информацию о школе 
    """
    format_borders = workbook.add_format()
    format_borders.set_bottom(6)
    format_borders.set_top(6)
    format_borders.set_bold()

    worksheet.write(cursor_row, 1, 'Код школы')
    worksheet.write(cursor_row, 2, data['s_code'])
    worksheet.write(cursor_row+1, 1, 'Итого ответов')
    worksheet.write(cursor_row+2, 1, 'Правильных ответов')

    cell = xl_rowcol_to_cell(cursor_row, 2)
    
    worksheet.set_row(cursor_row, cell_format=format_borders)
    cursor_row+=3
    
    return {"mun_cells_formula_start": [cell]}, cursor_row


def write_result_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[dict, int]:
    """
    Записывает информацию о результате ученика вместе с его классом и именем 
    """
    cursor_col = 1
    worksheet.write(cursor_row, cursor_col-1, data['test_uid'])
    worksheet.write(cursor_row, cursor_col,   data['class'])
    worksheet.write(cursor_row, cursor_col+1, data['name'])
    cursor_col+=1

    writen_results = []
    cell_start = xl_rowcol_to_cell(cursor_row, cursor_col+1)

    for answer in data['answers']: #Writing answer given by student
        writen_results.append(answer)
        worksheet.write(cursor_row, cursor_col+int(answer), data['answers'][answer])
    
    for i in range(1, data['q_num']+1, +1): #Checking for questions without an answer
        if i not in writen_results:
            worksheet.write(cursor_row, cursor_col+int(i), 0)

    cell_end = xl_rowcol_to_cell(cursor_row, cursor_col+data['q_num'])

    #Writing conditional formatting
    green_format = workbook.add_format({'bg_color': '00dd00'})
    red_format   = workbook.add_format({'bg_color': 'dd0000'})
    
    worksheet.conditional_format(f'{cell_start}:{cell_end}', {"type":     'cell', 
                                                              "criteria": '==',
                                                              "value":    '0',
                                                              "format":   red_format})

    worksheet.conditional_format(f'{cell_start}:{cell_end}', {"type":     'cell', 
                                                              "criteria": '==',
                                                              "value":    '1',
                                                              "format":   green_format})

    cursor_row += 1

    return {"mun_cells_formula_start": ['A1']}, cursor_row


def write_class_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:list):
    """
    Записывает информацию о классе и размечает место под статистику класса 
    """
    class_fromat = workbook.add_format()
    class_fromat.set_top(2)
    class_fromat.set_bottom(2)


    worksheet.write(cursor_row,   2, f'Класс {data[0][2]}, % правильных ответов')
    worksheet.write(cursor_row+1, 2, 'Правильных ответов')
    worksheet.write(cursor_row+2, 2, 'Неправильных ответов')
    
    cells = {}
    cells['class_percent']   = xl_rowcol_to_cell(cursor_row, 2)
    cells['class_correct']   = xl_rowcol_to_cell(cursor_row+1, 2)
    cells['class_incorrect'] = xl_rowcol_to_cell(cursor_row+2, 2)

    worksheet.set_row(cursor_row, cell_format=class_fromat)

    cursor_row+=3 


    return cells, cursor_row

def write_formulas():
    pass


if __name__ == '__main__':
    sys.exit()
