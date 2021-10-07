import mysql.connector
import config

def create_database_connection():
    """
    Создает подключение к базе данных 

    :return: соединение, курсор 
    """
    
    sqlConnection= mysql.connector.connect(
            host=config.host,
            user=config.user,
            password=config.password,
            database=config.database
            )

    return sqlConnection, sqlConnection.cursor()


def get_subjects()->list:
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
    sqlCon.close()
    return result


def get_all_munipals():
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
    sqlCon.close()
    return result


def get_active():
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
    sqlCon.close()
    return result


def get_schools_by_mo(mo_id:int):
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
    sqlCon.close()
    return result


def get_schools_by_mo_in_results(mo_id:int):
    """
    Получает список всех школ конкретного муниципалитета,
    которые присутствуют в таблице результатов 
    """

    sqlCon, cur = create_database_connection()
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
    sqlCon.close()
    return result


def get_subjects_tests(sbjid:int)->list:
    """
    Получает список модулей у предмета 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        * 
    FROM modules 
    WHERE id = {sbjid}
    """
    cur.execute(sql)
    result = cur.fetchall()
    sqlCon.close()
    return result


def get_students_resutls_of_school_by_test(sid:int, testid:int):
    """
    Получает результаты школы по модулю 
    """

    sqlCon, cur = create_database_connection()
    sql = f"""
    SELECT 
        * 
    FROM results 
    LEFT JOIN students ON results.student=students.id
    LEFT JOIN classes ON students.cid=classes.id
    WHERE 
        results.sid = {sid} AND 
        results.mid = {testid}
    
    """
    cur.execute(sql)
    result = cur.fetchall()
    sqlCon.close()
    return result


def get_question_test_data(mid:int):
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
    LEFT JOIN question_types ON modules_questions.q_type=question_types.id;
    """
    cur.execute(sql)
    result = cur.fetchall()
    sqlCon.close()
    return qnum, result

