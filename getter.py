import mysql.connector
from retry_decorator.retry_decorator import retry
import config

g_requests_count = 0 

sqlConnection = mysql.connector.connect(
        host=config.host,
        user=config.user,
        password=config.password,
        database=config.database
        )

@retry(Exception, tries=3, timeout_secs=1)
def create_database_connection():
    global sqlConnection, g_requests_count
    """
    Создает подключение к базе данных 

    :return: соединение, курсор 
    """
    return sqlConnection, sqlConnection.cursor()


def get_subjects()->list:
    global g_requests_count
    g_requests_count+=1 
    """
    Получает список всех предметов из базы данных 
    """

    sqlCon, cur = create_database_connection()
    sql = """
    SELECT 
        *
    FROM subjects; 
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_all_munipals():
    global g_requests_count
    g_requests_count+=1 
    """
    Получает список всех муниципалитетов, которые 
    присутствовали на тестировании 
    """
    
    sqlCon, cur = create_database_connection()
    sql = """
    SELECT 
        *
    FROM munipals
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_active():
    global g_requests_count
    g_requests_count+=1 
    """
    Получает список всех муниципалитетов, которые 
    присутствовали на тестировании 
    """
    
    
    sqlCon, cur = create_database_connection()
    sql = """
    SELECT 
        DISTINCT(munipals.id)
    FROM results
    LEFT JOIN main ON results.sid=main.id 
    LEFT JOIN munipals ON main.mo=munipals.id;
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_schools_by_mo(mo_id:int):
    global g_requests_count
    g_requests_count+=1 
    """
    Получает список всех школ конкретного муниципалитета 
    """
    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        *
    FROM main 
    WHERE 
        main.mo = {mo_id}
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_schools_by_mo_in_results(mo_id:int):
    """
    Получает список всех школ конкретного муниципалитета,
    которые присутствуют в таблице результатов 
    """
    global g_requests_count
    g_requests_count+=1 
    cur = create_database_connection()[1]
    sql = f"""
    SELECT 
        * 
    FROM main
    WHERE id in 
    (
        SELECT 
            DISTINCT(main.id)
        FROM results 
        LEFT JOIN main ON results.sid=main.id
        WHERE main.mo = {mo_id}
    )
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_subjects_tests(sbjid:int)->list:
    global g_requests_count
    g_requests_count+=1 
    """
    Получает список модулей у предмета 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        * 
    FROM modules 
    WHERE subject = {sbjid}
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_students_resutls_of_school_and_class_by_test(sid:int, cid:int, testid:int):
    global g_requests_count
    g_requests_count+=1 
    """
    Получает результаты школы по модулю 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
            distinct(test_uid)
    FROM results 
    LEFT JOIN students ON results.student=students.id
    LEFT JOIN classes ON students.cid=classes.id
    WHERE 
        results.sid = {sid} AND 
        results.mid = {testid} AND 
        classes.id  = {cid}
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_test_results_by_uid(uid:str):
    global g_requests_count
    g_requests_count+=1 
    """
    Получает резульаты конкретного модуля 
    """
    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT
        *
    FROM results
    LEFT JOIN modules_questions ON results.qid=modules_questions.id 
    WHERE test_uid = '{uid}'
    ORDER BY q_num
    """
    cur.execute(sql)
    result = cur.fetchall()

    return result


def get_question_test_data(mid:int):
    global g_requests_count
    g_requests_count+=1 
    """
    Получает информацию о вопросах теста 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT MAX(q_num) FROM modules_questions WHERE mid = {mid};
    """
    cur.execute(sql)
    qnum = cur.fetchall() #number of questions in module 
    
    sql = f"""
    SELECT 
        modules_questions.id,
        modules_questions.q_num,
        modules_questions.q_variant,
        modules_questions.answ1,
        modules_questions.answ2,
        modules_questions.answ3,
        modules_questions.answ4,
        modules_questions.correct_answ,
        question_types.name as 'q_type_name'
    FROM modules_questions
    LEFT JOIN question_types ON modules_questions.q_type=question_types.id
    WHERE mid = {mid}
    """
    cur.execute(sql)
    result = cur.fetchall()
    return qnum, result


def get_student_detailed_info(sid:int):
    global g_requests_count
    g_requests_count+=1 
    """
    Получить информацию о студенте 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        *
    FROM students 
    LEFT JOIN classes ON students.cid = classes.id 
    WHERE students.id = {sid}
    """

    cur.execute(sql)
    result = cur.fetchall()
    return result

def get_classes_of_school_by_test(sid:int, testid:int):
    global g_requests_count
    g_requests_count+=1 
    """
    Получает все классы, которые решали данный тест от школы 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        distinct(classes.id)
    FROM results 
    LEFT JOIN students ON results.student=students.id
    LEFT JOIN classes  ON students.cid=classes.id
    WHERE 
        results.sid = {sid} AND 
        results.mid = {testid}
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_class_info(cid:int) -> list:
    global g_requests_count
    g_requests_count+=1 
    """
    Получает все классы, которые решали данный тест от школы 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        *
    FROM classes 
    WHERE 
        classes.id = {cid} 
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


def get_module_questions(mid:int):
    global g_requests_count
    g_requests_count+=1 
    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        modules_questions.q_num, 
        question_types.name, 
        modules_questions.q_text, 
        modules_questions.q_variant,
        modules_questions.answ1,
        modules_questions.answ2,
        modules_questions.answ3,
        modules_questions.answ4, 
        modules_questions.correct_answ

    from question_types 
    LEFT JOIN modules_questions on modules_questions.q_type = question_types.id

    WHERE 
        mid = {mid}
    
    """
    cur.execute(sql)
    result = cur.fetchall()
    return result


if __name__ == '__main__':
    print(get_subjects())
    for subj in get_subjects():
        print(subj[0])
        print(get_subjects_tests(subj[0]))
