import string
import random
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
from io import BytesIO
from math import exp


# Генерация уникальных паролей
def generate_passwords(num=10):
    lowercase = string.ascii_lowercase
    uppercase = string.ascii_uppercase
    digits = string.digits
    special_chars = "!@#$%^&*()_+"
    all_chars = [lowercase, uppercase, digits, special_chars]
    passwords = []

    for _ in range(num):
        password_length = random.randint(6, 8)
        password = []

        for char_set in all_chars:
            password.append(random.choice(char_set))

        while len(password) < password_length:
            password.append(random.choice(random.choice(all_chars)))

        random.shuffle(password)
        passwords.append(''.join(password))

    alphabet = lowercase + uppercase + digits + ''.join(special_chars)
    return passwords, alphabet


# Оценка безопасности паролей
def evaluate_security(passwords, alphabet, V=3, T=10):
    T = T * 24 * 60
    results = []
    for password in passwords:
        L = len(password)
        A = len(alphabet)
        N = A ** L
        P = ((V * T) / N)
        results.append((password, L, f"{N:.1e}", f"{P:.1e}"))
    return results


# Создание EXCEL файла с данными и гистограммой
def create_excel(data, alphabet, passwords, probabilities, V=3, T=10):
    df = pd.DataFrame(data, columns=['Пароль', 'Длина', 'Возможные комбинации', 'Вероятность подбора'])
    initial_data = {'V': [V], 'T': [T], 'A': [len(alphabet)], 'L': '6-8'}
    df_initial = pd.DataFrame(initial_data)

    writer = pd.ExcelWriter('password_security_analysis.xlsx', engine='openpyxl')
    df_initial.to_excel(writer, index=False, sheet_name='Исходные Данные')
    df.to_excel(writer, index=False, sheet_name='Безопасность Паролей')

    # Создание вертикальной гистограммы с горизонтальными надписями паролей
    plt.figure(figsize=(12, 8))
    plt.bar(passwords, probabilities, color='skyblue')
    plt.ylabel('Вероятность подбора')
    plt.xlabel('Пароли')
    plt.title('Вероятность подбора паролей')
    plt.xticks(rotation=0)  # Горизонтальное отображение надписей паролей

    # Сохранение гистограммы в поток для добавления в EXCEL
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()

    # Добавление гистограммы в EXCEL файл
    book = writer.book
    sheet = book.create_sheet('Гистограмма')
    img = Image(img_bytes)
    sheet.add_image(img, 'A1')

    writer._save()
    return 'password_security_analysis.xlsx'


# Генерация паролей и оценка их безопасности
passwords, alphabet = generate_passwords()
security_results = evaluate_security(passwords, alphabet)

# Преобразование результатов для гистограммы
probs = [float(result[3].split('e')[0]) * 10 ** int(result[3].split('e')[1]) for result in security_results]

# Создание и сохранение EXCEL файла с гистограммой
excel_path = create_excel(security_results, alphabet, passwords, probs)
print(f"EXCEL файл сохранен: {excel_path}")
