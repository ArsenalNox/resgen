import pprint
pp = pprint.PrettyPrinter(indent=4, depth=22, width=90)

def refine_question_data(data):
    """
    Обрабатывает данные о вопросах теста 
    """

    new_data = {}
    for question in data:
        if question[1] not in new_data:
            new_data[question[1]] = {
                    "q_type": question[8]
                    }

    return new_data


def refine_student_answers(data):
    """
    Делает массив с пронумерованными ответами ученика
    """

    #Сначала нужно выбрать все ответы в которых желательно присутвует сам ответ 

    data = [answ for answ in data if answ[3] != '']

    answer_data = {}

    for answer in data:
        if answer[14] not in answer_data:
            answer_data[answer[14]] = answer[4]

    return answer_data 
