from logging import currentframe
import mysql.connector
import config 

sqlConnection = mysql.connector.connect(
    host= config.host,
    user= config.user,
    database= config.database,
    password= config.password
    )

cursor = sqlConnection.cursor()

sql = """
SELECT 
    DISTINCT(student) 
FROM results;
"""

cursor.execute(sql)

result = cursor.fetchall()

for student_id in result:
    sql = f"""
    SELECT 
        DISTINCT(test_uid)
    FROM results 
    WHERE student = {student_id[0]}
    """
    cursor.execute(sql)
    tests_id = cursor.fetchall()
    student_tests = []
    if len(tests_id) > 1:
        for test_id in tests_id:
            sql = f"""
            SELECT 
                qid,
                a_given,
                isCorrect,
                mid,
                sid,
                student
            FROM results
            WHERE test_uid = '{test_id[0]}'
            """
            cursor.execute(sql)
            
            test_content = cursor.fetchall()
            student_tests.append({
                'test_uid': test_id[0],
                'content':  test_content
                })

        print(student_tests)
