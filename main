import os

from openpyxl import load_workbook
from googletrans import Translator

PATH_TO_XLSX_FILE = ''

tb = load_workbook(PATH_TO_XLSX_FILE)
ts = tb.active
dict_words = {}
translator = Translator()
i = 1

while True:
    # Добавление уже записанных слов в dict_words
    if ts['A'+str(i)].value:
        dict_words[ts['A'+str(i)].value] = ts['B'+str(i)].value
        i += 1
        continue

    new_word = input('Введите новое слово: ').capitalize()

    # Прерывание или продолжение
    if new_word.lower() == 'stop':
        break
    if not new_word:
        print()
        continue

    # Проверка на наличие слова в dict_words
    if new_word in dict_words.keys():
        print(new_word, 'уже есть в списке со значением', f'\033[33m{dict_words[new_word]}\033[0m\n')
        continue

    # Получение верного варианта слова и его значения
    data = translator.translate(new_word, dest='ru', src='en')
    if type(data.extra_data['parsed'][-1]) is list:
        correct_word = data.extra_data['parsed'][-1][0].capitalize()
    else:
        correct_word = data.extra_data['parsed'][-2][-1][0].capitalize()
    value = data.text
    if not value:
        print('Не удалось перевести(\n')
        continue

    # Сравнение правильного слова со словом пользователя
    if new_word != correct_word:
        print('Вы имели ввиду', correct_word)
        new_word = correct_word

    # Проверка на наличие этого слова в dict_words
    if new_word in dict_words.keys():
        print(new_word, 'уже есть в списке со значением', dict_words[new_word], '\n')
        continue

    # Вывод и сохранение
    print(f'\033[34m{new_word}\033[0m', 'имеет значение', f'\033[33m{value}\033[0m')

    save = input('Сохранять? ')
    if not save:
        print()
        continue

    ts['A'+str(i)] = new_word
    ts['B'+str(i)] = value
    tb.save(PATH_TO_XLSX_FILE)
    # os.system('cls')
    print('Успешно сохранено! \n')





