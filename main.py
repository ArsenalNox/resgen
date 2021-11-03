import time
from mysql.connector.utils import print_buffer

from xlsxwriter.utility import xl_rowcol_to_cell
import helper
import getter
import writer 
import datetime
import os 

from pprint import pprint

from getter import g_requests_count

script_start_time = time.time()
print('')


#create new report directory
print('Creating new directory for results...')
working_directory_name = datetime.datetime.now().strftime("%d_%m_%Y-%H%M")
writer.create_new_report_directory(working_directory_name)
working_directory_name = 'results/' + working_directory_name
print(f'generating results to {working_directory_name}')

print('Getting subjects...')
for subject in getter.get_subjects(): #Iterating over all subjects 

    print(f'getting subjects tests {subject}\n')
    for module in getter.get_subjects_tests(subject[0]): #Iterating over all modules of subject


        cursor_row = 0 #Pointer to current workign row 

        #create new xlsx table for this subject
        worksheet_name = str(module[3]).replace(' ', '_')
        xltable, xlsheet= writer.create_xlsx_table(working_directory_name+'/'+worksheet_name)

        qnum, q_info = getter.get_question_test_data(module[0])

        header_info = { #Information which goes into header row in excel table 
                "test_name": f'Название теста: {module[3]}',
                "gen_date":  datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                "test_question_data": helper.refine_question_data(q_info)
        }

        cursor_row, cells_with_students_results, q_num = writer.write_header_info(xlsheet, cursor_row, header_info)
        
        writer.write_module_questions(xltable, module[0])

        print(f'Getting results for module: {module}\n\n')

        cells_final = []

        for munipal in getter.get_all_munipals(): #get munipal list 
            print('Getting munipal information..') 
            if len(getter.get_schools_by_mo_in_results(munipal[0])) == 0:
                print('Munipal is empty, skipping...')
                continue 

            cells_munipal_level, cursor_row = writer.write_munipal_info(xltable, xlsheet, cursor_row, {'mcode': munipal[0]}) #write munipal information 
            #^ remember the cells for generating formulas for statistics 
            cells_munipal_schools = []
            isSchoolListEmpty = True #Flag for checking if returned list is empty, i.e. if munipal has not participated in testing 
            for school in getter.get_schools_by_mo_in_results(munipal[0]): #get active schools of current munipal 
                print(f'    Writing school {school[2]}')
                isSchoolListEmpty = False

                if len(getter.get_classes_of_school_by_test(school[0], module[0])) == 0:
                   continue

                cells_school_level, cursor_row = writer.write_school_info(xltable, xlsheet, cursor_row, {'s_code': school[4]})
                cells_school_classes = []
                school_result_count = 0
                cells_class = {}

                for s_class in getter.get_classes_of_school_by_test(school[0], module[0]): #Get all id's of active classes
                    #print(f'Writing class {getter.get_class_info(s_class[0])}')
                    #Write class info and prepare cells for formulas
                    cells_class, cursor_row = writer.write_class_info(xltable, xlsheet, cursor_row, getter.get_class_info(s_class[0]))
                    

                    isFirstWrite = True
                    firstCell = ''
                    lastCell  = ''

                    #get all results ids by class, school and module
                    for sresult in getter.get_students_resutls_of_school_and_class_by_test(school[0], s_class[0], module[0]):  #get test_uid's 
                        #get every student's result 
                        #write student class, name, student result, remembering the cells by number of quenstion in module 
                        student_result_data = getter.get_test_results_by_uid(sresult[0])
                        answers             = helper.refine_student_answers(student_result_data)
                        student_info        = getter.get_student_detailed_info(student_result_data[0][7])
                            
                       #print(f'\n\n\n{sresult}\n{getter.get_student_detailed_info(student_result_data[0][6])}\n\n\n')
                       #print(f'{s_class}')
                       #print(f'{getter.get_class_info(s_class[0])}')
                       #for qer in student_result_data:
                       #    print(f'\n{qer}\n')
                        
                        cells, cursor_row = writer.write_result_info(xltable, xlsheet, cursor_row, {
                            "test_uid": sresult[0],
                            "class":    getter.get_class_info(s_class[0])[0][2], 
                            "name":     student_info[0][3]+' '+student_info[0][4]+' '+student_info[0][5],
                            "answers":  answers,
                            "q_num":    q_num
                            })

                        if isFirstWrite:
                            firstCell = cells
                            isFirstWrite = False
                        lastCell = cells
                        school_result_count += 1
                    #write formulas for class statistic
                    cells_class['q_num']      = q_num
                    cells_class['data_cells'] = [firstCell, lastCell]
                    ccell1, ccell2 = writer.write_class_formula(xltable, xlsheet, 1, cells_class) #get cells with classes stat 
                    cells_school_classes.append(xl_rowcol_to_cell(ccell1, ccell2))

                #write school stat, get cells for munipal stat 
                cells_munipal_school = writer.write_school_formula(xltable, xlsheet, {
                    'school_cell': cells_school_level,
                    'classes':     cells_school_classes,
                    'q_num':       q_num
                    })
                cells_munipal_schools.append(cells_munipal_school)
                print('Done!')

            #write munipal stat, get cells for 
            cell_final = writer.write_munipal_formula(xltable, xlsheet, {
                "q_num":       q_num,
                "school_data": cells_munipal_schools,
                "mun":         cells_munipal_level
                })
            print(f'Done writing munipal {munipal[1]}')
            cells_final.append(cell_final)

        writer.write_final_formula(xltable, xlsheet, {
        "cells": cells_final,
        "start": "D2",
        "q_num": q_num
        })
        xltable.close() #close file

print(f'script run time: {time.time()-script_start_time}s')
print(f'Requests: {g_requests_count}')