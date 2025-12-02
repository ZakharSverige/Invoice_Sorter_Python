import os
import shutil
import pandas as pd

def sort_invoices_by_courier():
    # --- Настройки ---
    root_data_folder = r"E:\ВВР\Односторонняя" # Основная папка с данными
    courier_data_file = "Курьеры дома 1 вариант.xls - Лист1.csv" # Имя вашего CSV-файла
    unmatched_folder_name = "Несопоставленные_Файлы" # Название папки для несопоставленных файлов

    # --- Загрузка данных курьеров ---
    try:
        courier_df = pd.read_csv(courier_data_file, encoding='utf-8')
        # Создаем словарь для быстрого поиска: 'Андропова д.2' -> 'Иванов Иван'
        # Заменяем " д." на "_" и убираем точки для сопоставления с именами файлов
        courier_mapping = {
            row['Адрес'].replace(' д.', '_').replace('.', ''): row['ФИО']
            for index, row in courier_df.iterrows()
        }
        print("Данные курьеров успешно загружены.")
    except FileNotFoundError:
        print(f"Ошибка: Файл '{courier_data_file}' не найден. Убедитесь, что он находится в той же папке, что и скрипт.")
        return
    except Exception as e:
        print(f"Ошибка при чтении CSV файла: {e}")
        print("Пожалуйста, убедитесь, что файл сохранен в кодировке UTF-8 и имеет правильные заголовки: 'Адрес', 'ФИО'.")
        return

    # --- Создание папки для несопоставленных файлов ---
    unmatched_path = os.path.join(root_data_folder, unmatched_folder_name)
    os.makedirs(unmatched_path, exist_ok=True)
    print(f"Папка для несопоставленных файлов: {unmatched_path}")

    # --- Обработка папок с почтовыми индексами ---
    for postal_code in range(150000, 150065): # От 150000 до 150064 включительно
        folder_path = os.path.join(root_data_folder, str(postal_code))

        if not os.path.exists(folder_path):
            print(f"Предупреждение: Папка '{folder_path}' не найдена. Пропускаем.")
            continue

        print(f"\nОбработка папки: {folder_path}")

        for filename in os.listdir(folder_path):
            if filename.lower().endswith(".pdf"):
                full_file_path = os.path.join(folder_path, filename)
                
                # Извлечение адреса из имени файла (Текущая логика - split)
                # Ожидаем формат: XXX_Улица_Дом.pdf или XXX_Улица_Дом_корп_YY.pdf
                parts = filename.split('_', 1) # Разделяем только по первому подчеркиванию
                if len(parts) < 2:
                    print(f"Пропуск: Некорректное имя файла (нет адреса после первого '_'): {filename}")
                    shutil.move(full_file_path, os.path.join(unmatched_path, filename))
                    continue
                
                # Убираем расширение .pdf и возможные точки
                address_from_filename = parts[1].rsplit('.', 1)[0].replace('.', '')
                
                # Пытаемся найти курьера по адресу из имени файла
                courier_name = courier_mapping.get(address_from_filename)

                if courier_name:
                    target_courier_folder = os.path.join(root_data_folder, courier_name)
                    os.makedirs(target_courier_folder, exist_ok=True)
                    
                    destination_path = os.path.join(target_courier_folder, filename)
                    try:
                        shutil.move(full_file_path, destination_path)
                        print(f"Перемещено '{filename}' в '{courier_name}'")
                    except Exception as e:
                        print(f"Ошибка при перемещении '{filename}': {e}")
                else:
                    print(f"Несопоставлено: Адрес '{address_from_filename}' из файла '{filename}' не найден в данных курьеров.")
                    shutil.move(full_file_path, os.path.join(unmatched_path, filename))
    
    print("\nСортировка завершена!")
    print(f"Файлы, которые не удалось сопоставить, находятся в папке: {unmatched_path}")

# Запуск функции
if __name__ == "__main__":