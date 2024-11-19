import docx
import matplotlib.pyplot as plt
from collections import Counter

lion = docx.Document('lion.docx') #берешь произведение
doc = docx.Document() #файл для вывода таблицы
text = [] #бустой лист для того, чтобы добавить в него текст
letters = ['а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 'к', 'л', 'м', 'н', 'о', 'п', 'р', 'с', 'т', 'у', 'ф', 'х', 'ц',
            'ч', 'ш', 'щ', 'ь', 'ы', 'ъ', 'э', 'ю', 'я', 'q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'a', 's', 'd', 'f', 'g', 'h',
            'j', 'k', 'l', 'z', 'x', 'c', 'v', 'b', 'n', 'm', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
#все элементы, которые могут входить в слова или даты
table = doc.add_table(rows = 1, cols = 3) #создаем таблицу с одной строкой и тремя колонками
table.style = 'Table Grid'

hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Слово' #подписываем 3 колонки
hdr_cells[1].text = 'Частота встречи, раз'
hdr_cells[2].text = 'Частота встречи, %'

for paragraph in lion.paragraphs:
    text.append(paragraph.text) #добавлякм весь наш текств раннее созданную переменную

text = '\n'.join(text) #делаем весь текст в одну строку
text_word = text.lower() #в новую переменную добавляем текст только маленькими буквами
text_word = text_word.split() #разбираем его на слова
for i in range(len(text_word)): #создаем цикл, который будет проходить по всем словам
    a = 0 #создаем переменную для подсчета пройденных строк
    print(i)

    word = text_word[i] #переменная будет каждый раз присваиваться новому слову
    if len(word) == 1 and word[0] not in letters: #если длина слова и его состав неподходят под условие
        continue #то сразу пропускаем его
    if word[0] not in letters: #если первый элемент не входит в список
        word = word[1:] #то удаляем его
    if len(word) == 1 and word[0] not in letters: #если длина слова и его состав неподходят под условие
        continue #то сразу пропускаем его
    if word[-1] not in letters: #если последний элемент
        word = word[:-1] #тоже удаляем
    if len(word) == 1 and word[0] not in letters: #если длина слова и его состав неподходят под условие
        continue #то сразу пропускаем его
    if word[-1] not in letters: #и предпоследний 
        word = word[:-1] 
        #это делается для того, чтобы исключить поодобные ситуации " (слово), "

    for row in table.rows: #перебираем каждую строку

        if row.cells[0].text == word: #если слово уже ранее было в первом столбце(в какой-либо из строк)
            row.cells[1].text = str(int(row.cells[1].text) + 1) #то к значению во втором слолбце мы просто добавляем 1
            break #и выходим из этого цикла, чтобы взять новое слово

        else:
            a += 1 #если нету, то проходим все строки до конца

    if a == len(table.rows): #проверяем, если программа прошла все столбцы, то такого слова в таблице нету

        new_row = table.add_row() #мы добавляем новую строку
        new_row.cells[0].text = word #в первую коллонку добавляем само слово
        new_row.cells[1].text = '1' #а во вторую 1, чтобы потом прибавлять к этому значению

for row in table.rows[1:]: #заново проходим все строки кроме первой, чтобы заболнить 3 столбец таблицы
    row.cells[2].text = str(int(row.cells[1].text)/len(text_word)*100) #заполняем его

#фильтруем текст, оставляем только маленькие строчные буквы
filtered_text = ''.join(filter(lambda x: 'а' <= x <= 'я', text.lower()))

#считаем количество каждой буквы
letter_counts = Counter(filtered_text)

#сортируем буквы по алфавиту
sorted_letters = sorted(letter_counts.keys())
sorted_counts = [letter_counts[letter] for letter in sorted_letters]

#создание графика
plt.figure(figsize=(12, 6))
plt.bar(sorted_letters, sorted_counts, color='skyblue')
plt.xlabel('Буквы')
plt.ylabel('Количество')
plt.title('Количество строчных русских букв в тексте')
plt.grid(axis='y')

# Отображаем график
plt.tight_layout()
plt.show()

doc.save('table.docx')