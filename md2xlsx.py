import pandas 
import re
from tkinter import Tk, filedialog

def extract_part_text(md_content, pattern):
    """Извлекает разделы текста между метками частей из Markdown-контента.

    Функция ищет все вхождения заданного шаблона (например, "ЧАСТЬ 1") и извлекает текст между ними.
    Возвращает список словарей с номером части и соответствующим текстом.

    Args:
        md_content: Исходный текст в формате Markdown.
        pattern: Регулярное выражение для поиска разделов (например, r'\\section\*\{ЧАСТЬ\s*(\d)+\}').

    Returns:
        Список словарей вида [{'part': '1', 'unparced_text': '...'}, ...], где:
            - 'part' — номер части (str)
            - 'unparced_text' — текст между метками частей (str)

    Raises:
        re.error: Если передано некорректное регулярное выражение.
    """

    data = []
    while True:
        start_match = re.search(pattern, md_content, re.IGNORECASE)
        if not start_match:
            break  

        start_pos = start_match.end()

        end_match = re.search(pattern, md_content[start_pos:], re.IGNORECASE)
        part_number = re.search(r'(\d)+',
                                md_content[start_match.start():start_match.end()])
        if end_match:
            end_pos = start_pos + end_match.start()
            part_number = re.search(
                r'(\d)+',
                md_content[start_match.start():start_match.end()])
            part = {
                'part': part_number.group(),
                'unparced_text': md_content[start_pos:end_pos].strip()
            }
            data.append(part)
            md_content = md_content[end_pos:]
            continue

        part = {
            'part': part_number.group(),
            'unparced_text': md_content[start_pos:].strip()
        }
        data.append(part)
        break
    return data

def parcer(content):
    """Парсит необработанный текст вопросов из разделов.

    Обрабатывает текст, извлеченный extract_part_text(), выделяя:
    - Номера вопросов (например, "В1")
    - Текст вопросов
    - Ссылки на изображения
    - Подписи к рисункам

    Args:
        content: Список словарей с необработанным текстом (результат extract_part_text()).

    Returns:
        Список словарей с распарсенными данными вида:
        [
            {
                'part': '2', 
                'quest_number': 'В1', 
                'quest': 'Текст вопроса...', 
                'pix': 'http://example.com/image.jpg'
            }, 
            ...
        ]
        Отсутствующие ключи заполняются значением None.
    """

    parced_text = []
    for part_number in content:
        unparced_text = part_number['unparced_text'].split('\n')
        parts = {
            'part': part_number['part']
        }

        for i in range(len(unparced_text)):
            item = unparced_text[i]
            quest_number_match = re.search(r'^[A-ZА-ЯЁ]\d+\s',
                                           item)
            if quest_number_match:
                parts['quest_number'] = quest_number_match.group()[:-1]
                parts['quest'] = item[quest_number_match.end():]
            elif item[0:3] == '![]':
                parts['pix'] = item[4:-1] 
            elif item[:4].lower() == 'Рис.'.lower():
                parts['quest_number'] += f'\n\n{item}'
            elif item != '' or re.search('^.+', item):
                parts['quest_number'] += f'\n\n{item}'
            if i == len(unparced_text) - 1 or re.search(r'^[A-ZА-ЯЁ]\d+\s',
                                                        unparced_text[i+1]):
                parced_text.append(parts)
                parts = {
                    'part': part_number['part']
                }
            
    return parced_text

def parse_md_to_excel(content):
    """Сохраняет распарсенные данные в Excel-файл.

    Преобразует структурированные данные в DataFrame и сохраняет их в output.xlsx.
    Автоматически обрабатывает отсутствующие значения и задает порядок колонок.

    Args:
        content: Список словарей с данными вопросов (результат parcer()).

    Effects:
        Создает файл output.xlsx в текущей директории со следующими колонками:
        - Часть (part)
        - Номер вопроса (quest_number)
        - Вопрос (quest)
        - Рисунок (pix)

    Note:
        Требует установки библиотек pandas и openpyxl.
    """

    all_keys = set()
    for row in content:
        all_keys.update(row.keys())

    # Заполняем отсутствующие ключи значением None
    processed_data = []
    for row in content:
        new_row = {key: row.get(key, None) for key in all_keys}
        processed_data.append(new_row)

    columns_order = ['part', 'quest_number', 'quest', 'pix']


    # # Парсинг и сохранение в .xlsx
    df = pandas.DataFrame(processed_data, columns=columns_order,)
    df.rename(columns={
        'part': 'Часть',
        'quest_number': 'Номер вопроса',
        'quest': 'Вопрос',
        'pix': 'Рисунок'

    }, inplace=True)
    df.to_excel('output.xlsx', index=False, engine='openpyxl')




if __name__=='__main__':
    root = Tk()
    root.withdraw()  # Спрятать окно, если это не нужно

    file_path = filedialog.askopenfile()
    if file_path:
        print(f"Выбранный файл: {str(file_path)}")
    else:
        exit("Выбор папки отменен.")


    with open('92e27331-e7eb-4794-947a-7fe3d2df18cd (1) (1).md', 'r', encoding='utf-8') as f:
        md_content = f.read()

    text = extract_part_text(
        md_content,
        r'\\[a-z]+.*\{\s*ЧАСТЬ\s*(\d)+\s*\}',  # Регулярка для начала части
    )
    parcing_text = parcer(text)

    parse_md_to_excel(parcing_text)

    print("Парсинг завершен. Результат сохранен в output.xlsx")