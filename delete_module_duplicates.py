import mysql.connector
import config 

sqlConnection = mysql.connector.connect(
    host= config.host,
    user= config.user,
    database= config.database,
    password= config.password
    )

cursor = sqlConnection.cursor()

def get_students():
    global cursor

    sql = """
    SELECT 
        DISTINCT(student) 
    FROM results;
    """
    cursor.execute(sql)
    return cursor.fetchall()


def get_student_modules(student:int): 
    global cursor
    sql = f"""
    SELECT 
        DISTINCT(mid)
    FROM results
    WHERE student = {student}
    """
    cursor.execute(sql)
    return cursor.fetchall()


def get_student_results(student:int, mid:int):
    global cursor
    sql = f"""
    SELECT 
        DISTINCT(test_uid)
    FROM results
    WHERE 
        student = {student}
        AND 
        mid = {mid}
    """
    cursor.execute(sql)
    return cursor.fetchall()


def get_content(test_uid:str):
    global cursor
    sql = f"""
    SELECT 
        *
    FROM results
    WHERE 
        test_uid = '{test_uid}'
    """
    cursor.execute(sql)
    return cursor.fetchall()


def delete_results(test_uid:str):
    global cursor
    sql = f"""
    DELETE FROM results WHERE test_uid = '{test_uid}'
    """
    print(sql)
    cursor.execute(sql)

for student in get_students():
    print(f'Student {student[0]}')
    for s_mid in get_student_modules(student[0]):
        
        if not len(get_student_results(student[0], s_mid[0]))>1:
            continue

        results = []
        highest_count = 0
        for s_result in get_student_results(student[0], s_mid[0]):
            correct_count = 0
            for content in get_content(s_result[0]):
                if content[4] == 1:
                    correct_count+=1 

            if correct_count > highest_count:
                highest_count = correct_count

            results.append({'uid': s_result[0], 'count': correct_count})

        print(results)
        best_results = [result for result in results if result['count']==highest_count]

        do_not_delete = best_results[0]

        for res in results:
            if res['uid'] == do_not_delete['uid']:
                continue
            delete_results(str(res['uid']))
    print('')
