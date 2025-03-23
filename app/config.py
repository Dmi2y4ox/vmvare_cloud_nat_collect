import os

# Получаем базовую директорию (где находится config.py)
base_dir = os.path.dirname(os.path.abspath(__file__))

# Файл загруженный из VMware Cloud
#input_file = os.path.join(base_dir, '..', 'excel', 'data.txt')

#Файл для тестового запуска
input_file = os.path.join(base_dir, '..', 'test', 'data.txt')

# Файл, в который будет записана информация
output_file = os.path.join(base_dir, '..', 'excel', 'Inform.xlsx')

#Функция выбора имени облака в листе output_file
#Предполагается, что название листов в Excel файле соответствуют названию ваших облаков
def get_cloud_name():
    orgselect = input("С какого облака собрана информация?\n1.cloud1\n2.cloud2\n3.cloud3\n4.cloud4\n")
    if orgselect == "1":
        return 'cloud1'
    elif orgselect == "2":
        return 'cloud2'
    elif orgselect == "3":
        return 'cloud3'
    elif orgselect == "4":
        return 'cloud4'
    else:
        return None