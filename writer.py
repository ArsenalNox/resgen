import xlsxwriter
import os
import sys
from xlsxwriter.utility   import xl_cell_to_rowcol, xl_rowcol_to_cell 

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


def write_munipal_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[dict, int]:
    """
    Записывает информацию о муниципалитете 
    """
    
    format_bold = workbook.add_format()
    format_bold.set_bold()
    
    worksheet.write(cursor_row, 1, f'Муниципалитет {data["mcode"]}, % правильных ответов')
    worksheet.set_row(cursor_row, None, format_bold)

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

    format_bold = workbook.add_format()
    format_bold.set_bold()
    format_bold.set_bottom(8)
    format_bold.set_top(8)

    format_background = workbook.add_format()
    format_background.set_bg_color('ffc000')

    worksheet.write(cursor_row, 1, 'Код школы')
    worksheet.write(cursor_row, 2, data['s_code'])
    worksheet.write(cursor_row+1, 1, 'Итого ответов', format_background)
    worksheet.write(cursor_row+2, 1, 'Правильных ответов', format_background)

    cell = xl_rowcol_to_cell(cursor_row, 2)
    worksheet.set_row(cursor_row,   None, format_bold)
    worksheet.set_row(cursor_row+1, None, None, {'level': 2})
    worksheet.set_row(cursor_row+2, None, None, {'level': 2})
    cursor_row+=3
    
    return {"mun_cells_formula_start": [cell]}, cursor_row


def write_result_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[str, int]:
    """
    Записывает информацию о результате ученика вместе с его классом и именем 
    """
    cursor_col = 1
    worksheet.write(cursor_row, cursor_col-1, data['test_uid'])
    worksheet.write(cursor_row, cursor_col,   data['class'])
    worksheet.write(cursor_row, cursor_col+1, data['name'])
    
    worksheet.set_row(cursor_row,   None, None, {'level': 2})

    cursor_col+=1

    writen_results = []
    cell_start = str(xl_rowcol_to_cell(cursor_row, cursor_col+1))

    for answer in data['answers']: #Writing answer given by student
        writen_results.append(answer)
        worksheet.write(cursor_row, cursor_col+int(answer), data['answers'][answer])
    
    for i in range(1, data['q_num']+1, +1): #Checking for questions without an answer
        if i not in writen_results:
            worksheet.write(cursor_row, cursor_col+int(i), 0)

    cell_end = xl_rowcol_to_cell(cursor_row, cursor_col+data['q_num'])

    #Writing formatting
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

    return cell_start, cursor_row


def write_class_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:list)->Tuple[dict, int]:
    """
    Записывает информацию о классе и размечает место под статистику класса 
    """
    class_fromat = workbook.add_format()
    class_fromat.set_top(7)
    class_fromat.set_bottom(7)

    worksheet.write(cursor_row,   2, f'Класс {data[0][2]}, % правильных ответов')
    worksheet.write(cursor_row+1, 2, 'Правильных ответов')
    worksheet.write(cursor_row+2, 2, 'Неправильных ответов')

    worksheet.set_row(cursor_row, None, class_fromat, {'level': 1})
    worksheet.set_row(cursor_row+1, None, None, {'level': 2})
    worksheet.set_row(cursor_row+2, None, None, {'level': 2})
    
    cells = {}
    cells['class_percent']   = xl_rowcol_to_cell(cursor_row, 3)
    cells['class_correct']   = xl_rowcol_to_cell(cursor_row+1, 3)
    cells['class_incorrect'] = xl_rowcol_to_cell(cursor_row+2, 3)


    cursor_row+=3 


    return cells, cursor_row

def write_class_formula(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict):
    cell_format = workbook.add_format()
    cell_format.set_num_format(9)

    cell1, cell2 = data['data_cells'][0], data['data_cells'][1]
 
    rowC, colC = xl_cell_to_rowcol(data['class_correct'])
    rowI, colI = xl_cell_to_rowcol(data['class_incorrect'])
    rowP, colP = xl_cell_to_rowcol(data['class_percent'])
   
    for i in range(0, data['q_num'], +1):

        data['data_cells'][0] = increment_cell_col(cell1, i)
        data['data_cells'][1] = increment_cell_col(cell2, i)

        worksheet.write(rowP, colP+i, f'={increment_cell_col(data["class_correct"], i)}/{increment_cell_col(data["class_incorrect"], i)}', cell_format)
        worksheet.write(rowC, colC+i, f'=COUNTIF({data["data_cells"][0]}:{data["data_cells"][1]}, "1")')
        worksheet.write(rowI, colI+i, f'=COUNTIF({data["data_cells"][0]}:{data["data_cells"][1]}, "0")')

    return


def increment_cell_col(cell:str, increment:int) -> str:
    row, col = xl_cell_to_rowcol(cell)
    col += increment

    inc = str(xl_rowcol_to_cell(row, col))

    return inc


if __name__ == '__main__':
    sys.exit()
