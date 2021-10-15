import time
import helper
import getter
import writer 
import datetime

from pprint import pprint

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

    print('getting subjects tests\n')
    for module in getter.get_subjects_tests(subject[0]): #Iterating over all modules of subject
        print(f'Getting results for module: {module}\n\n')
        
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
        

        print('Getting munipal information..') 
        for munipal in getter.get_all_munipals(): #get munipal list 
            cells, cursor_row = writer.write_munipal_info(xltable, xlsheet, cursor_row, {'mcode': munipal[0]}) #write munipal information 
            #^ remember the cells for generating formulas for statistics 
            
            isSchoolListEmpty = True #Flag for checking if returned list is empty, i.e. if munipal has not participated in testing 
            for school in getter.get_schools_by_mo_in_results(munipal[0]): #get active schools of current munipal 
                isSchoolListEmpty = False

                if len(getter.get_classes_of_school_by_test(school[0], module[0])) == 0:
                   continue

                cells_school_level, cursor_row = writer.write_school_info(xltable, xlsheet, cursor_row, {'s_code': school[4]})
                school_result_count = 0;
                cells_class = {}

                for s_class in getter.get_classes_of_school_by_test(school[0], module[0]): #Get all id's of active classes
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
                    writer.write_class_formula(xltable, xlsheet, 1, cells_class)
                    
                print(f'Starting cell for class {firstCell}, last cell {lastCell}, writing in cell {cells_class}')
                print(f'Results in school {school[2]}: {school_result_count}')
            print('\n')

        #write formulas for school
        #write formulas for munipals 
        #write formulas for unified stats 
        xltable.close() #close file 
        break
    break
