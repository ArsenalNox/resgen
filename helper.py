def refine_question_data(data):
    """
    Обрабатывает данные о вопросах теста 
    """

    new_data = {}

    for question in data:
        if data[1] not in new_data:
            print('Adding new question...')
            new_data[question[1]] = {
                    "q_type": question[8]
                    }
    print(new_data)
    return new_data
