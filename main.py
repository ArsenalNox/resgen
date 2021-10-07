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
working_directory_name = datetime.datetime.now().strftime("%d_%m_%Y-%H%m")
writer.create_new_report_directory(working_directory_name)
working_directory_name = 'results/' + working_directory_name
print(f'generating results to {working_directory_name}')

print('Getting subjects...')
for subject in getter.get_subjects(): #Iterating over all subjects 

    print('getting subjects tests\n')
    for module in getter.get_subjects_tests(subject[0]): #Iterating over all modules of subject
        print(f'Getting results for module: {module}\n\n')
        
        is_empty_results = False
        cursor_row = 0
        #create new xlsx table for this subject
        worksheet_name = str(module[3]).replace(' ', '_')
        xltable, xlsheet= writer.create_xlsx_table(working_directory_name+'/'+worksheet_name)
        qnum, q_info = getter.get_question_test_data(module[0])
        pprint(q_info)
        header_info = {
                "test_name": f'Название теста: {module[1]}',
                "gen_date":  f'',
                "test_question_data": helper.refine_question_data(q_info)
                }
        writer.write_header_info(xlsheet, cursor_row, header_info)
        
        print('Getting munipal information..') #get munipal list 
        for munipal in getter.get_all_munipals():
            #write munipal information 
            #remember the cells for generating formulas for statistics 
            #check if munipal participated in testing, if not - dont remember 
            
            #get active schools of current munipal 
            print('Getting mo\'s schools')
            isSchoolListEmpty = True
            for school in getter.get_schools_by_mo_in_results(munipal[0]):
                isSchoolListEmpty = False
                pprint(school)
                #write student result, remembering the cells by number of quenstion in module 
                for sresult in getter.get_students_resutls_of_school_by_test(school[0], module[0]):
                    print(sresult)
                
            if isSchoolListEmpty:
                print(f'Munipal {munipal[0]} ({munipal[1]}) has not participated in testing')

            print('\n')

        #write result to file, close it 
        xltable.close()
        if is_empty_results:
            print('No results were found')
