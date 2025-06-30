import os
import pandas as pd
import django
from django.db.models import Q

# Настройка Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tiketon.settings')
django.setup()

# Импорт модели Scanner после настройки Django
from core.models import Scanner

def import_scanners_from_excel(excel_file_path):
    try:
        # Загрузка Excel файла
        print(f"Загрузка файла: {excel_file_path}")
        df = pd.read_excel(excel_file_path, sheet_name="список сканеров")
        
        # Проверка наличия столбца B
        if len(df.columns) < 2:
            print("Ошибка: В файле нет столбца B")
            return
        
        # Получение данных из столбца B (индекс 1)
        names_column = df.iloc[:, 1].dropna()
        
        # Счетчики
        added_count = 0
        skipped_count = 0
        
        print(f"Найдено {len(names_column)} записей в столбце B")
        
        # Обработка каждого имени
        for full_name in names_column:
            # Пропускаем пустые строки и заголовки
            if not isinstance(full_name, str) or len(full_name.strip()) == 0:
                continue
                
            # Разделение полного имени на имя и фамилию
            name_parts = full_name.strip().split()
            
            if len(name_parts) >= 2:
                first_name = name_parts[0]
                last_name = " ".join(name_parts[1:])
                
                # Проверка, существует ли уже сканер с таким именем и фамилией
                existing_scanner = Scanner.objects.filter(
                    Q(first_name__iexact=first_name) & Q(last_name__iexact=last_name)
                ).first()
                
                if existing_scanner:
                    print(f"Пропуск: {first_name} {last_name} (уже существует)")
                    skipped_count += 1
                else:
                    # Создание нового сканера
                    Scanner.objects.create(
                        first_name=first_name,
                        last_name=last_name
                    )
                    print(f"Добавлен: {first_name} {last_name}")
                    added_count += 1
            else:
                print(f"Пропуск: '{full_name}' (некорректный формат имени)")
                skipped_count += 1
        
        print("\nРезультаты импорта:")
        print(f"Добавлено новых сканеров: {added_count}")
        print(f"Пропущено: {skipped_count}")
        
    except Exception as e:
        print(f"Ошибка при импорте: {str(e)}")

if __name__ == "__main__":
    # Путь к Excel файлу
    excel_file = "сканеры тикетон астана (2).xlsx"
    
    # Проверка существования файла
    if not os.path.exists(excel_file):
        print(f"Ошибка: Файл '{excel_file}' не найден")
    else:
        import_scanners_from_excel(excel_file) 