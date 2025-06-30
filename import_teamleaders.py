import os
import django

# Настройка Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tiketon.settings')
django.setup()

from django.db.models import Q
from django.contrib.auth.models import User, Group
from core.models import TeamLeaderProfile, Scanner, TeamLeader

def create_team_leader(first_name, last_name, email):
    try:
        # Проверка, существует ли уже пользователь с таким email
        existing_user = User.objects.filter(email=email).first()
        if existing_user:
            print(f"Пропуск: {first_name} {last_name} ({email}) - пользователь с таким email уже существует")
            return False
            
        # Создание имени пользователя на основе email
        username = email.split('@')[0]
        
        # Проверка, существует ли пользователь с таким username
        if User.objects.filter(username=username).exists():
            # Добавляем цифру к имени пользователя, если оно занято
            base_username = username
            counter = 1
            while User.objects.filter(username=username).exists():
                username = f"{base_username}{counter}"
                counter += 1
        
        # Создание пользователя
        user = User.objects.create_user(
            username=username,
            email=email,
            password='Freedom2024',  # Временный пароль
            first_name=first_name,
            last_name=last_name
        )
        
        # Добавление пользователя в группу "Тимлидеры"
        team_leaders_group, created = Group.objects.get_or_create(name='Тимлидеры')
        user.groups.add(team_leaders_group)
        
        # Создание профиля тимлидера
        TeamLeaderProfile.objects.create(user=user)
        
        print(f"Добавлен тимлидер (пользователь): {first_name} {last_name} ({email})")
        return True
    
    except Exception as e:
        print(f"Ошибка при добавлении тимлидера (пользователя) {first_name} {last_name}: {str(e)}")
        return False

def check_scanner_exists(first_name, last_name):
    """
    Проверяет, существует ли сканер с указанным именем и фамилией.
    Учитывает возможность того, что имена могут быть на английском или русском.
    """
    # Прямая проверка
    scanner = Scanner.objects.filter(
        Q(first_name__iexact=first_name) & Q(last_name__iexact=last_name)
    ).first()
    
    if scanner:
        return scanner
    
    # Проверка с транслитерацией (примерная)
    # Список соответствий для простой транслитерации
    ru_to_en = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'zh',
        'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
        'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts',
        'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu',
        'я': 'ya', 'ұ': 'u', 'ү': 'u', 'қ': 'k', 'ғ': 'g', 'ә': 'a', 'і': 'i', 'ң': 'n',
        'һ': 'h', 'ө': 'o'
    }
    
    # Функция для транслитерации
    def transliterate(text):
        result = ""
        for char in text.lower():
            result += ru_to_en.get(char, char)
        return result.capitalize()
    
    # Попытка найти сканера с транслитерированным именем
    first_name_en = transliterate(first_name)
    last_name_en = transliterate(last_name)
    
    scanner = Scanner.objects.filter(
        Q(first_name__iexact=first_name_en) & Q(last_name__iexact=last_name_en)
    ).first()
    
    return scanner

def import_team_leaders():
    # Список тимлидеров для импорта (имена и фамилии поменяны местами и переведены на английский)
    team_leaders_data = [
        ("Abeldinova", "Linara", "linaraluntik@gmail.com"),
        ("Aman", "Kuralay", "aman.kuralaj1708@gmail.com"),
        ("Bukharbayeva", "Aiganym", "bukharbaevaaiganym@gmail.com"),
        ("Demchenko", "Elizaveta", "demchenkoelizaveta06@gmail.com"),
        ("Zhanatbekkyzy", "Laila", "lajlazanatbekkyzy@gmail.com"),
        ("Kairat", "Alikhan", "idkrecord18@gmail.com"),
        ("Kim", "Daniil", "danilbip6758@gmail.com"),
        ("Manarbay", "Abylay", "manarbay.ab@nisa.edu.kz"),
        ("Mukhamedzhanova", "Aruzhan", "aruzhan2405@gmail.com"),
        ("Tolegen", "Aisha", "tolegen.a06@icloud.com"),
        ("Zhumagul", "Abylkas", "zhumagul.a@nisa.edu.kz"),
        ("Khaibrakhmanov", "Amir", "amir.haibrahmanov@gmail.com"),
        ("Ospanov", "Akdaulet", "aakdauletospanovv@gmail.com"),
        ("Kairbekova", "Aiganym", "aiganymkairbekova07@gmail.com"),
        ("Khakimova", "Sofia", "sofiya.kha0113@gmail.com"),
        ("Abdilmanov", "Ansar", "1207ansar@gmail.com"),
        ("Kazbek", "Aigerim", "a.kazbek@internet.ru"),
        ("Zhaishylyk", "Asem", "zhaishylykassem@gmail.com"),
        ("Gazizov", "Amir", "gazizov_amir_ga@mail.ru"),
        ("Turganova", "Dzhamilya", "deamaali.17@gmail.com"),
        ("Kenzhebekov", "Batyrkhan", ""),
        ("Albek", "Aday", "adayalbek5@gmail.com"),
        ("Abilmazhinova", "Adel", "adelabilmazhinova8@gmail.com"),
        ("Zhusupova", "Zhangerim", "zusupovazangerim@gmail.com"),
        ("Islyamova", "Alua", "yaktak218@gmail.com")
    ]
    
    added_count = 0
    skipped_count = 0
    scanner_found_count = 0
    
    for first_name, last_name, email in team_leaders_data:
        if email.strip() == "":
            print(f"Пропуск: {first_name} {last_name} - отсутствует email")
            skipped_count += 1
            continue
        
        # Проверка наличия в списке сканеров
        scanner = check_scanner_exists(first_name, last_name)
        if scanner:
            print(f"Найден сканер: {scanner.first_name} {scanner.last_name} - будет добавлен как тимлидер")
            scanner_found_count += 1
        
        # Проверяем, существует ли уже такой тимлидер в новой модели
        existing_teamleader = TeamLeader.objects.filter(
            Q(first_name__iexact=first_name) & Q(last_name__iexact=last_name)
        ).first()
        
        if existing_teamleader:
            print(f"Тимлидер уже существует в новой модели: {first_name} {last_name}")
            
            # Обновляем связь со сканером, если она отсутствует
            if not existing_teamleader.scanner and scanner:
                existing_teamleader.scanner = scanner
                existing_teamleader.save()
                print(f"Обновлена связь со сканером для тимлидера: {first_name} {last_name}")
        else:
            # Создаем запись в новой модели TeamLeader
            new_teamleader = TeamLeader(
                first_name=first_name,
                last_name=last_name,
                email=email
            )
            
            # Если нашли сканера, связываем его с тимлидером
            if scanner:
                new_teamleader.scanner = scanner
            
            # Сохраняем тимлидера (при сохранении автоматически создастся сканер, если его нет)
            new_teamleader.save()
            print(f"Добавлен тимлидер в новую модель: {first_name} {last_name}")
            added_count += 1
        
        # Создаем пользователя-тимлидера, если его еще нет
        success = create_team_leader(first_name, last_name, email)
        if not success:
            print(f"Пользователь-тимлидер уже существует или не удалось создать: {first_name} {last_name}")
    
    print("\nРезультаты импорта тимлидеров:")
    print(f"Добавлено в новую модель: {added_count}")
    print(f"Найдено в списке сканеров: {scanner_found_count}")
    print(f"Пропущено: {skipped_count}")

if __name__ == "__main__":
    import_team_leaders() 