from pprint import pprint
import xlsxwriter
import os
import sys
import getter
from xlsxwriter.utility   import xl_cell_to_rowcol, xl_rowcol_to_cell 

from xlsxwriter.workbook  import Workbook
from xlsxwriter.worksheet import Worksheet
from typing               import Tuple


tree_leves = {
    'school': 4,
    'school_classes': 3,
    'class_results':  2
}


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
    #worksheet.write(cursor_row+1, 1, '% выполнения')
    #worksheet.write(cursor_row+2, 1, 'Всего правильных')
    #worksheet.write(cursor_row+3, 1, 'Всего неправильных')

    working_col = 3

    start_row = cursor_row
    start_col = working_col
    cell = xl_rowcol_to_cell(start_row, start_col)

    q_num = 0
    q_max = len(data['test_question_data'])
    for q_info in data['test_question_data']:
        if q_max == q_num+1:
            break
        worksheet.write(cursor_row, working_col, data['test_question_data'][q_info]['q_type'])
        working_col+=1 
        q_num+=1

    cursor_row += 4

    return cursor_row, [cell], q_num


def write_munipal_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[str, int]:
    """
    Записывает информацию о муниципалитете 
    """
    
    format_bold = workbook.add_format()
    format_bold.set_bold()
    
    worksheet.write(cursor_row, 1, f'Муниципалитет {data["mcode"]}, % правильных ответов')
    worksheet.write(cursor_row+1, 1, f'Всего правильных')
    worksheet.write(cursor_row+2, 1, f'Всего неправильных')
    worksheet.set_row(cursor_row+1, None, None, {'level': 2, 'collapsed': True})
    worksheet.set_row(cursor_row+2, None, None, {'level': 2, 'collapsed': True})
    worksheet.set_row(cursor_row, None, format_bold)

    cell = str(xl_rowcol_to_cell(cursor_row, 1))
    cursor_row+=3
        
    return cell, cursor_row


def write_school_info(workbook: Workbook ,worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[str, int]:
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

    worksheet.write(cursor_row, 1, f'Код школы, {data["s_code"]}')
    worksheet.write(cursor_row, 2, 'процент правильных')
    worksheet.set_row(cursor_row,   None, format_bold, {'level': 1, 'collapsed': True})
    worksheet.write(cursor_row+1, 1, 'Итого ответов', format_background)
    worksheet.write(cursor_row+2, 1, 'Правильных ответов', format_background)
    
    cell = str(xl_rowcol_to_cell(cursor_row, 2))
    worksheet.set_row(cursor_row+1, None, None, {'level': 2, 'collapsed': True})
    worksheet.set_row(cursor_row+2, None, None, {'level': 2, 'collapsed': True})
    cursor_row+=3
    
    return cell, cursor_row


def write_result_info(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict) -> Tuple[str, int]:
    """
    Записывает информацию о результате ученика вместе с его классом и именем 
    """
    cursor_col = 1
    #worksheet.write(cursor_row, cursor_col-1, data['test_uid'])
    worksheet.write(cursor_row, cursor_col,   data['class'])
    worksheet.write(cursor_row, cursor_col+1, data['name'])
    
    worksheet.set_row(cursor_row,   None, None, {'level': 3, 'collapsed': True})

    cursor_col+=1

    writen_results = []
    cell_start = str(xl_rowcol_to_cell(cursor_row, cursor_col+1))

    for answer in data['answers']: #Writing answer given by student
        if data['q_num'] < int(answer):
            continue
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

    worksheet.set_row(cursor_row, None, class_fromat, {'level': 2, 'collapsed': True})
    worksheet.set_row(cursor_row+1, None, None, {'level': 3, 'collapsed': True})
    worksheet.set_row(cursor_row+2, None, None, {'level': 3, 'collapsed': True})
    
    cells = {}
    cells['class_percent']   = xl_rowcol_to_cell(cursor_row, 3)
    cells['class_correct']   = xl_rowcol_to_cell(cursor_row+1, 3)
    cells['class_incorrect'] = xl_rowcol_to_cell(cursor_row+2, 3)


    cursor_row+=3 


    return cells, cursor_row


def write_class_formula(workbook: Workbook, worksheet: Worksheet, cursor_row: int, data:dict):
    """
    Записывает формулы в ячейки классов 
    """
    cell_format = workbook.add_format()
    cell_format.set_num_format(9)

    cell1, cell2 = data['data_cells'][0], data['data_cells'][1]
 
    rowC, colC = xl_cell_to_rowcol(data['class_correct'])
    rowI, colI = xl_cell_to_rowcol(data['class_incorrect'])
    rowP, colP = xl_cell_to_rowcol(data['class_percent'])
   
    for i in range(0, data['q_num'], +1):

        data['data_cells'][0] = increment_cell_col(cell1, i)
        data['data_cells'][1] = increment_cell_col(cell2, i)
        
        cell_c_answ = increment_cell_col(data["class_correct"], i)
        cell_w_answ = increment_cell_col(data["class_incorrect"], i)

        worksheet.write(rowP, colP+i, f'=IFERROR({cell_c_answ}/({cell_w_answ}+{cell_c_answ}), 1)', cell_format)
        worksheet.write(rowC, colC+i, f'=COUNTIF({data["data_cells"][0]}:{data["data_cells"][1]}, "1")')
        worksheet.write(rowI, colI+i, f'=COUNTIF({data["data_cells"][0]}:{data["data_cells"][1]}, "0")')

    worksheet.conditional_format(
            f'{data["class_percent"]}:{increment_cell_col(data["class_percent"], data["q_num"])}',
            {
                "type": '3_color_scale',
                "min_color": 'red',
                "mid_color": 'yellow',
                "max_color": 'green',
                "mid_value": '50%',
                "max_value": '100%',
                "min_value": '0%',
                "min_type": 'num',
                "max_type": 'num',
                "mid_type": 'num'
                }
        )


    return rowP, colP 


def write_school_formula(workbook: Workbook, worksheet: Worksheet, data:dict):
    """
    Записывает формулы в ячейки школ 
    """
        
    cell_format = workbook.add_format()
    cell_format.set_num_format(9)

    for i in range(1, data['q_num']+1, +1):
        worksheet.write(
                increment_cell_col(data['school_cell'], i), 
                f'=IFERROR(AVERAGE({",".join(increment_cell_col_array(data["classes"], i-1))}), 0)',
                cell_format)

        worksheet.conditional_format(
                increment_cell_col(data['school_cell'], i),
                {
                "type": '3_color_scale',
                "min_color": 'red',
                "mid_color": 'yellow',
                "max_color": 'green',
                "min_value": '0%',
                "mid_value": '50%',
                "max_value": '100%',
                "min_type": 'num',
                "mid_type": 'num',
                "max_type": 'num'
                }
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['school_cell'], 1), i),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['classes'], 1), i-1))}"
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['school_cell'], 2), i),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['classes'], 2), i-1))}"
            )

    return data['school_cell']


def write_munipal_formula(workbook: Workbook, worksheet: Worksheet, data:dict):

    cell_format = workbook.add_format()
    cell_format.set_num_format(9)

    for i in range(1, data['q_num']+1, +1):
        worksheet.write(
                increment_cell_col(data['mun'], i+1), 
                f'=IFERROR(AVERAGE({",".join(increment_cell_col_array(data["school_data"], i))}), 0)',
                cell_format)

        worksheet.conditional_format(
                increment_cell_col(data['mun'], i+1),
                {
                "type": '3_color_scale',
                "min_color": 'red',
                "mid_color": 'yellow',
                "max_color": 'green',
                "mid_value": '50%',
                "max_value": '100%',
                "min_value": '0%',
                "min_type": 'num',
                "max_type": 'num',
                "mid_type": 'num'
                }
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['mun'], 1), i+1),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['school_data'], 1), i))}"
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['mun'], 2), i+1),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['school_data'], 2), i))}"
            )

    return data['mun']


def write_final_formula(workbook: Workbook, worksheet: Worksheet, data:dict):

    cell_format = workbook.add_format()
    cell_format.set_num_format(9)

    for i in range(0, data['q_num'], +1):
        worksheet.write(
                increment_cell_col(data['start'], i), 
                f'=IFERROR(AVERAGE({",".join(increment_cell_col_array(data["cells"], i+2))}), 0)',
                cell_format)
        
        worksheet.conditional_format(
                increment_cell_col(data['start'], i),
                {
                    "type": '3_color_scale',
                    "min_color": 'red',
                    "mid_color": 'yellow',
                    "max_color": 'green',
                    "mid_value": '50%',
                    "max_value": '100%',
                    "min_value": '0%',
                    "min_type": 'num',
                    "mid_type": 'num',
                    "max_type": 'num'
                    }
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['start'], 1), i),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['cells'], 1), i+2))}"
            )
        
        worksheet.write(
            increment_cell_col(increment_cell_row(data['start'], 2), i),
            f"={'+'.join(increment_cell_col_array(increment_cell_row_array(data['cells'], 2), i+2))}"
            )

    return 


def increment_cell_col(cell:str, increment:int) -> str:
    row, col = xl_cell_to_rowcol(cell)
    col += increment

    inc = str(xl_rowcol_to_cell(row, col))

    return inc


def increment_cell_row(cell:str, increment:int) -> str:
    row, col = xl_cell_to_rowcol(cell)
    row += increment
    
    inc = str(xl_rowcol_to_cell(row, col))

    return inc


def increment_cell_row_array(cells:list, increment:int) -> list[str]:
    arr_new = []
    for cell in cells:
        arr_new.append(increment_cell_row(cell, increment))

    return arr_new


def increment_cell_col_array(cells:list, increment:int) -> list[str]:
    arr_new = []
    for cell in cells:
        arr_new.append(increment_cell_col(cell, increment))

    return arr_new


def write_module_questions(workbook: Workbook, module_id:int):
    #Создать новый лист в файле, записать туда все вопросы модуля 
    sheet = workbook.add_worksheet('Вопросы')
    questions = getter.get_module_questions(module_id)
    cursor_row = 1
    sheet.write(0,0, "Номер вопроса")
    sheet.write(0,1, "Тип вопроса")
    sheet.write(0,2, "Текст вопроса")
    sheet.write(0,3, "Вариант")
    sheet.write(0,4, "Ответ 1")
    sheet.write(0,5,"Ответ 2")
    sheet.write(0,6,"Ответ 3")
    sheet.write(0,7,"Ответ 4")
    sheet.write(0,8, "Правильный ответ")
    
    for questions in questions:
            column = 0            
            for question_content in questions:
                sheet.write(cursor_row,column,question_content)
                column+=1 
            cursor_row +=1        
    return 


if __name__ == '__main__':
    sys.exit()

