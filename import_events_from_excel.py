import os
import pandas as pd
import django
from datetime import datetime, timedelta
from django.utils import timezone

# Настройка Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tiketon.settings')
django.setup()

# Импорт моделей после настройки Django
from core.models import Event, User, Scanner, EventParticipant, TeamLeader
from django.contrib.auth.models import Group

def import_events_from_excel(excel_file_path):
    try:
        # Загрузка Excel файла
        print(f"Загрузка файла: {excel_file_path}")
        
        # Получаем все листы в файле
        excel_file = pd.ExcelFile(excel_file_path)
        sheet_names = excel_file.sheet_names
        
        print(f"Найдены листы: {sheet_names}")
        
        # Счетчики
        added_count = 0
        skipped_count = 0
        error_count = 0
        
        # Получаем первого пользователя для created_by (или создаем админа)
        admin_user = User.objects.filter(is_superuser=True).first()
        if not admin_user:
            # Если нет админа, берем первого пользователя
            admin_user = User.objects.first()
            if not admin_user:
                print("Ошибка: Нет пользователей в системе")
                return
        
        # Обработка каждого листа
        for sheet_name in sheet_names:
            print(f"\nОбработка листа: {sheet_name}")
            
            try:
                # Читаем данные листа
                df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
                
                # Пропускаем пустые листы
                if df.empty:
                    print(f"Лист '{sheet_name}' пустой, пропускаем")
                    continue
                
                print(f"Найдено {len(df)} строк в листе '{sheet_name}'")
                
                # Ищем колонки с данными
                # Предполагаем, что первая колонка - это название/дата мероприятия
                first_col = df.columns[0]
                
                # Обрабатываем каждую строку
                for index, row in df.iterrows():
                    if pd.isna(row[first_col]):
                        continue
                    
                    # Извлекаем данные из строки
                    event_data = extract_event_data(row, df.columns)
                    
                    if event_data and event_data['name']:  # Проверяем, что есть название
                        # Ищем ответственного
                        created_by_user = find_user_by_name(event_data['team_leader'])
                        if not created_by_user:
                            created_by_user = admin_user
                        
                        # Проверяем, существует ли уже такое мероприятие
                        existing_event = Event.objects.filter(
                            name=event_data['name'],
                            date=event_data['date']
                        ).first()
                        
                        if existing_event:
                            print(f"Пропуск: {event_data['name']} ({event_data['date']}) - уже существует")
                            skipped_count += 1
                        else:
                            # Создаем новое мероприятие
                            event = Event.objects.create(
                                name=event_data['name'],
                                date=event_data['date'],
                                location=event_data.get('location', ''),
                                description=event_data.get('description', ''),
                                max_scanners=event_data.get('max_scanners', 0),
                                created_by=created_by_user,
                                start_date=event_data.get('start_date'),
                                end_date=event_data.get('end_date'),
                                duration_hours=event_data.get('hours_awarded', 1.0)
                            )
                            print(f"Добавлено: {event_data['name']} ({event_data['date']})")
                            added_count += 1
                            
                            # Добавляем участников мероприятия
                            if event_data['scanners_list']:
                                participants_added = add_event_participants(event, event_data['scanners_list'], event_data['hours_awarded'])
                                print(f"Добавлено участников: {participants_added}")
                    else:
                        error_count += 1
                        
            except Exception as e:
                print(f"Ошибка при обработке листа '{sheet_name}': {str(e)}")
                error_count += 1
        
        print(f"\nРезультаты импорта:")
        print(f"Добавлено мероприятий: {added_count}")
        print(f"Пропущено (уже существуют): {skipped_count}")
        print(f"Ошибок: {error_count}")
        
    except Exception as e:
        print(f"Ошибка при импорте: {str(e)}")

def extract_event_data(row, columns):
    """Извлекает данные мероприятия из строки"""
    try:
        # Ищем колонки по заголовкам
        event_name = ''
        event_date = None
        team_leader = ''
        max_scanners = 0
        hours_awarded = 0.0
        scanners_list = []
        
        # Проходим по всем колонкам и извлекаем данные
        for i, col in enumerate(columns):
            if pd.isna(row[col]):
                continue
                
            cell_value = str(row[col]).strip()
            if not cell_value or cell_value.lower() in ['nan', '']:
                continue
            
            col_name = str(col).lower()
            
            # Название мероприятия
            if any(keyword in col_name for keyword in ['название', 'мероприятие', 'event', 'name']):
                event_name = cell_value
            
            # Дата мероприятия
            elif any(keyword in col_name for keyword in ['дата', 'date']):
                event_date = parse_date(cell_value)
            
            # Ответственный/тимлидер
            elif any(keyword in col_name for keyword in ['ответственный', 'тимлидер', 'leader', 'team']):
                team_leader = cell_value
            
            # Количество сканеров
            elif any(keyword in col_name for keyword in ['кол-во', 'количество', 'count', 'scanners']):
                try:
                    max_scanners = int(float(cell_value))
                except:
                    pass
            
            # Часы
            elif any(keyword in col_name for keyword in ['часы', 'hours', 'час']):
                try:
                    hours_awarded = float(cell_value)
                except:
                    pass
            
            # Список сканеров
            elif any(keyword in col_name for keyword in ['сканеры', 'участники', 'participants', 'scanner']):
                scanners_list = parse_scanners_list(cell_value)
        
        # Если дату не нашли, пробуем извлечь из названия
        if not event_date and event_name:
            event_date = extract_date_from_string(event_name)
        
        # Если дату все еще не нашли, используем текущую
        if not event_date:
            event_date = timezone.now().date()
        
        # Очищаем название от дат
        if event_name:
            event_name = clean_event_name(event_name)
        
        return {
            'name': event_name,
            'date': event_date,
            'team_leader': team_leader,
            'max_scanners': max_scanners,
            'hours_awarded': hours_awarded,
            'scanners_list': scanners_list
        }
        
    except Exception as e:
        print(f"Ошибка при извлечении данных: {str(e)}")
        return None

def parse_date(date_str):
    # Ищем паттерны дат
    date_patterns = [
        r'(\d{2}\.\d{2}\.\d{4})',  # DD.MM.YYYY
        r'(\d{2}/\d{2}/\d{4})',    # DD/MM/YYYY
        r'(\d{4}-\d{2}-\d{2})',    # YYYY-MM-DD
    ]
    
    import re
    for pattern in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            date_str = match.group(1)
            try:
                if '.' in date_str:
                    return datetime.strptime(date_str, '%d.%m.%Y').date()
                elif '/' in date_str:
                    return datetime.strptime(date_str, '%d/%m/%Y').date()
                elif '-' in date_str:
                    return datetime.strptime(date_str, '%Y-%m-%d').date()
            except:
                continue
    
    return None

def parse_scanners_list(scanners_str):
    scanners_list = []
    scanners_str = scanners_str.strip()
    if scanners_str:
        scanners_list = [scanner.strip() for scanner in scanners_str.split(',')]
    
    return scanners_list

def extract_date_from_string(date_str):
    # Ищем паттерны дат
    date_patterns = [
        r'(\d{2}\.\d{2}\.\d{4})',  # DD.MM.YYYY
        r'(\d{2}/\d{2}/\d{4})',    # DD/MM/YYYY
        r'(\d{4}-\d{2}-\d{2})',    # YYYY-MM-DD
    ]
    
    import re
    for pattern in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            date_str = match.group(1)
            try:
                if '.' in date_str:
                    return datetime.strptime(date_str, '%d.%m.%Y').date()
                elif '/' in date_str:
                    return datetime.strptime(date_str, '%d/%m/%Y').date()
                elif '-' in date_str:
                    return datetime.strptime(date_str, '%Y-%m-%d').date()
            except:
                continue
    
    return None
def extract_date_from_data(row, columns):
    """Извлекает дату из данных строки"""
    try:
        # Ищем дату в любой из колонок
        for col in columns:
            if not pd.isna(row[col]):
                cell_value = row[col]
                
                # Если это ячейка с датой
                if isinstance(cell_value, (datetime, pd.Timestamp)):
                    return cell_value.date()
                
                # Если это строка, пытаемся распарсить дату
                if isinstance(cell_value, str):
                    # Ищем паттерны дат
                    date_patterns = [
                        r'(\d{2}\.\d{2}\.\d{4})',  # DD.MM.YYYY
                        r'(\d{2}/\d{2}/\d{4})',    # DD/MM/YYYY
                        r'(\d{4}-\d{2}-\d{2})',    # YYYY-MM-DD
                    ]
                    
                    import re
                    for pattern in date_patterns:
                        match = re.search(pattern, cell_value)
                        if match:
                            date_str = match.group(1)
                            try:
                                if '.' in date_str:
                                    return datetime.strptime(date_str, '%d.%m.%Y').date()
                                elif '/' in date_str:
                                    return datetime.strptime(date_str, '%d/%m/%Y').date()
                                elif '-' in date_str:
                                    return datetime.strptime(date_str, '%Y-%m-%d').date()
                            except:
                                continue
        
        return None
        
    except Exception as e:
        return None

def find_user_by_name(full_name):
    """Ищет пользователя по имени и фамилии"""
    if not full_name:
        return None
    
    name_parts = full_name.strip().split()
    if len(name_parts) >= 2:
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:])
        
        # Ищем пользователя по имени и фамилии
        user = User.objects.filter(
            first_name__iexact=first_name,
            last_name__iexact=last_name
        ).first()
        
        if user:
            print(f"Найден пользователь: {first_name} {last_name}")
            return user
        else:
            print(f"Пользователь не найден: {first_name} {last_name}")
    
    return None

def add_event_participants(event, scanners_list, hours_awarded):
    """Добавляет участников мероприятия"""
    participants_count = 0
    
    for scanner_name in scanners_list:
        if not scanner_name.strip():
            continue
            
        # Ищем сканера по имени
        scanner = find_scanner_by_name(scanner_name.strip())
        if scanner:
            # Проверяем, не участвует ли уже сканер в мероприятии
            existing_participant = EventParticipant.objects.filter(
                event=event,
                volunteer=scanner
            ).first()
            
            if not existing_participant:
                EventParticipant.objects.create(
                    event=event,
                    volunteer=scanner,
                    hours_awarded=hours_awarded,
                    hours_awarded_backup=hours_awarded
                )
                participants_count += 1
                print(f"Добавлен участник: {scanner_name}")
            else:
                print(f"Участник уже существует: {scanner_name}")
        else:
            print(f"Сканер не найден: {scanner_name}")
    
    return participants_count

def find_scanner_by_name(full_name):
    """Ищет сканера по имени и фамилии"""
    if not full_name:
        return None
    
    name_parts = full_name.strip().split()
    if len(name_parts) >= 2:
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:])
        
        # Ищем сканера по имени и фамилии
        scanner = Scanner.objects.filter(
            first_name__iexact=first_name,
            last_name__iexact=last_name
        ).first()
        
        return scanner
    
    return None

def clean_event_name(name):
    """Очищает название мероприятия от лишней информации"""
    import re
    
    # Удаляем даты из названия
    name = re.sub(r'\d{2}\.\d{2}\.\d{4}', '', name)
    name = re.sub(r'\d{2}/\d{2}/\d{4}', '', name)
    name = re.sub(r'\d{4}-\d{2}-\d{2}', '', name)
    
    # Удаляем лишние пробелы и символы
    name = name.strip()
    name = re.sub(r'\s+', ' ', name)
    
    # Ограничиваем длину названия
    if len(name) > 200:
        name = name[:197] + '...'
    
    return name

if __name__ == "__main__":
    # Путь к Excel файлу
    excel_file = "сканеры тикетон астана.xlsx"
    
    # Проверка существования файла
    if not os.path.exists(excel_file):
        print(f"Ошибка: Файл '{excel_file}' не найден")
    else:
        import_events_from_excel(excel_file)
