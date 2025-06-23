from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login
from django.contrib.auth.models import User
from django.core.mail import send_mail
from django.conf import settings
from django.http import HttpResponse, JsonResponse, FileResponse
import random
from django.contrib.auth.decorators import login_required, user_passes_test
from .models import Event, Scanner, EventParticipant
from django.utils import timezone
from django.db.models import Q, Sum, Min, Max, FloatField
from django.db.models.functions import Coalesce
import openpyxl
from openpyxl.utils import get_column_letter
from django.contrib import messages
from django.contrib.auth.models import Group
from dateutil import parser as dateparser
from django.core.paginator import Paginator
from openpyxl.styles import Font, PatternFill, Alignment
import os
import io
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import tempfile
from datetime import timedelta

# Проверка доступа (только тимлидеры и админы)
def is_team_leader_or_admin(user):
    return user.is_authenticated and (user.is_team_leader or user.is_staff)

def team_leader_required(view_func):
    return user_passes_test(is_team_leader_or_admin)(view_func)

# Авторизация
def login_request(request):
    if request.method == 'POST':
        name = request.POST.get('name')
        email = request.POST.get('email')
        try:
            user = User.objects.get(email=email)
        except User.DoesNotExist:
            return render(request, 'core/login.html', {'error': 'Только для тимлидеров'})
        code = str(random.randint(100000, 999999))
        request.session['auth_code'] = code
        request.session['auth_email'] = email
        send_mail(
            'Код для входа',
            f'Ваш код: {code}',
            settings.DEFAULT_FROM_EMAIL,
            [email],
            fail_silently=False,
        )
        return redirect('code_verify')
    return render(request, 'core/login.html')

def code_verify(request):
    if request.method == 'POST':
        code = request.POST.get('code')
        if code == request.session.get('auth_code'):
            email = request.session.get('auth_email')
            user = User.objects.get(email=email)
            login(request, user)
            return redirect('events')
        return render(request, 'core/code_verify.html', {'error': 'Неверный код'})
    return render(request, 'core/code_verify.html')

@team_leader_required
def events_list(request):
    events = Event.objects.all().order_by('-date')
    # Поиск по названию
    q = request.GET.get('q', '').strip()
    if q:
        events = events.filter(name__icontains=q)
    # Фильтр по дате
    date_from = request.GET.get('date_from')
    date_to = request.GET.get('date_to')
    if date_from:
        events = events.filter(date__gte=date_from)
    if date_to:
        events = events.filter(date__lte=date_to)
    # Фильтр по тимлидеру
    leader = request.GET.get('leader')
    if leader:
        events = events.filter(leader__first_name__icontains=leader) | events.filter(leader__last_name__icontains=leader)
    # Пагинация
    paginator = Paginator(events, 20)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    return render(request, 'core/events_list.html', {
        'events': page_obj.object_list,
        'page_obj': page_obj,
        'q': q,
        'date_from': date_from,
        'date_to': date_to,
        'leader': leader,
    })

@team_leader_required
def event_create(request):
    scanners = Scanner.objects.all()
    if request.method == 'POST':
        name = request.POST.get('name')
        date = request.POST.get('date')
        volunteers_required = request.POST.get('volunteers_required')
        event = Event.objects.create(
            name=name,
            date=date,
            volunteers_required=volunteers_required,
            leader=request.user
        )
        return redirect('event_edit', event_id=event.id)
    return render(request, 'core/event_create.html', {'scanners': scanners})

@team_leader_required
def event_edit(request, event_id):
    # Проверяем, что пользователь является создателем события или админом
    event = Event.objects.get(id=event_id)
    all_scanners = Scanner.objects.all()
    participants = EventParticipant.objects.filter(event=event).select_related('volunteer')
    
    if request.method == 'POST':
        if request.user.is_staff or event.leader == request.user:
            if 'add_volunteer' in request.POST:
                volunteer_id = request.POST.get('volunteer_id')
                if volunteer_id and not participants.filter(volunteer_id=volunteer_id).exists():
                    EventParticipant.objects.create(event=event, volunteer_id=volunteer_id)
                return redirect('event_edit', event_id=event.id)
            
            if 'save_participants' in request.POST and request.user.is_staff:
                for p in participants:
                    is_late = request.POST.get(f'late_{p.id}') == 'on'
                    late_minutes = int(request.POST.get(f'late_minutes_{p.id}', 0)) if is_late else 0
                    p.is_late = is_late
                    p.late_minutes = late_minutes
                    p.save()
                    
            if 'set_duration' in request.POST and request.user.is_staff:
                duration_hours = float(request.POST.get('duration_hours', 0))
                event.duration_hours = duration_hours
                event.save()
                for p in participants:
                    late_minutes = p.late_minutes or 0
                    # Переводим опоздание в часы и вычитаем из общей длительности
                    late_hours = late_minutes / 60
                    awarded_hours = max(duration_hours - late_hours, 0)
                    p.hours_awarded = awarded_hours
                    p.save()
            
        return redirect('event_edit', event_id=event.id)
    
    return render(request, 'core/event_edit.html', {
        'event': event,
        'all_scanners': all_scanners,
        'participants': participants,
        'is_creator': event.leader == request.user,
        'is_admin': request.user.is_staff
    })

def volunteer_search(request):
    q = request.GET.get('q', '').strip()
    event_id = request.GET.get('event_id')
    exclude_ids = list(EventParticipant.objects.filter(event_id=event_id).values_list('volunteer_id', flat=True)) if event_id else []
    if q:
        scanners = Scanner.objects.filter(
            Q(first_name__icontains=q) | Q(last_name__icontains=q) | Q(email__icontains=q)
        ).exclude(id__in=exclude_ids)[:10]
    else:
        scanners = Scanner.objects.none()
    data = [
        {'id': v.id, 'name': f'{v.first_name} {v.last_name}', 'email': v.email}
        for v in scanners
    ]
    return JsonResponse({'results': data})

def home(request):
    if request.user.is_authenticated:
        return redirect('events')
    else:
        return redirect('login')

def export_event_participants(request, event_id):
    event = Event.objects.get(id=event_id)
    participants = EventParticipant.objects.filter(event=event).select_related('volunteer')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Участники'
    ws.append(['ФИО', 'Email', 'Опоздал (мин)', 'Зачтено (часы:минуты)'])
    for p in participants:
        fio = f'{p.volunteer.first_name} {p.volunteer.last_name}'
        email = p.volunteer.email
        late = p.late_minutes or 0
        hours = int(p.hours_awarded)
        minutes = int(round((p.hours_awarded - hours) * 60))
        time_str = f'{hours}:{minutes:02d}'
        ws.append([fio, email, late, time_str])
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 22
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename=event_{event.id}_participants.xlsx'
    wb.save(response)
    return response

@team_leader_required
def import_excel(request):
    if request.method == 'POST' and request.FILES.get('file'):
        file = request.FILES['file']
        wb = openpyxl.load_workbook(file)
        added = {'scanners': 0, 'leaders': 0, 'events': 0, 'participants': 0}
        # Импорт из всех листов, кроме 'список сканеров'
        for sheet in wb.sheetnames:
            if sheet.lower() == 'список сканеров':
                continue
            ws = wb[sheet]
            current_event = None
            current_leader = None
            current_hours = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                # Если строка с заполненными A–E — это событие
                if all(row[i] for i in range(5)):
                    team_fio = row[0]
                    event_name = row[1]
                    event_date = row[2]
                    volunteers_required = row[3]
                    hours = row[4]
                    # Парсим дату
                    event_date_str = str(event_date).split('-')[0].split()[0] if '-' in str(event_date) else str(event_date)
                    try:
                        parsed_date = dateparser.parse(event_date_str, dayfirst=True, fuzzy=True)
                        event_date_val = parsed_date.date()
                    except Exception:
                        current_event = None
                        continue
                    # Проверка числовых значений
                    try:
                        hours_val = float(hours)
                        volunteers_required_val = int(volunteers_required)
                    except Exception:
                        current_event = None
                        continue
                    # Тимлидер
                    team_parts = str(team_fio).strip().split()
                    team_first = team_parts[0]
                    team_last = ' '.join(team_parts[1:]) if len(team_parts) > 1 else ''
                    team_email = f'{team_first.lower()}.{team_last.lower()}@tiketon.local'
                    group, _ = Group.objects.get_or_create(name='Тимлидеры')
                    leader, created = User.objects.get_or_create(email=team_email, defaults={
                        'username': team_email,
                        'first_name': team_first,
                        'last_name': team_last,
                        'password': User.objects.make_random_password()
                    })
                    if group not in leader.groups.all():
                        leader.groups.add(group)
                    # Мероприятие
                    event, _ = Event.objects.get_or_create(
                        name=event_name,
                        date=event_date_val,
                        leader=leader,
                        defaults={
                            'volunteers_required': volunteers_required_val,
                            'duration_hours': hours_val
                        }
                    )
                    current_event = event
                    current_leader = leader
                    current_hours = hours_val
                    if created:
                        added['events'] += 1
                    if created and leader:
                        added['leaders'] += 1
                else:
                    # Строки с волонтёрами (сканерами) после строки с событием
                    if current_event and row[0]:
                        vol_fio = str(row[0])
                        vol_parts = vol_fio.strip().split()
                        if len(vol_parts) >= 2:
                            vol_first = vol_parts[0]
                            vol_last = ' '.join(vol_parts[1:])
                            vol_email = f'{vol_first.lower()}.{vol_last.lower()}@scanner.local'
                            # Создаём или получаем сканера
                            scanner, created = Scanner.objects.get_or_create(
                                email=vol_email,
                                defaults={
                                    'first_name': vol_first,
                                    'last_name': vol_last
                                }
                            )
                            if created:
                                added['scanners'] += 1
                            # Добавляем участие
                            participant, p_created = EventParticipant.objects.get_or_create(
                                event=current_event,
                                volunteer=scanner,
                                defaults={
                                    'hours_awarded': current_hours or 0
                                }
                            )
                            if p_created:
                                added['participants'] += 1
        # Теперь обрабатываем лист со сканерами, если он есть
        if 'список сканеров' in [s.lower() for s in wb.sheetnames]:
            sheet_name = next(s for s in wb.sheetnames if s.lower() == 'список сканеров')
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    vol_fio = str(row[0])
                    vol_email = row[1] if len(row) > 1 and row[1] else None
                    vol_parts = vol_fio.strip().split()
                    if len(vol_parts) >= 2:
                        vol_first = vol_parts[0]
                        vol_last = ' '.join(vol_parts[1:])
                        if not vol_email:
                            vol_email = f'{vol_first.lower()}.{vol_last.lower()}@scanner.local'
                        # Создаём или получаем сканера
                        scanner, created = Scanner.objects.get_or_create(
                            email=vol_email,
                            defaults={
                                'first_name': vol_first,
                                'last_name': vol_last
                            }
                        )
                        if created:
                            added['scanners'] += 1
        return render(request, 'core/import_excel.html', {'added': added})
    return render(request, 'core/import_excel.html')

def export_all_events(request):
    events = Event.objects.all().order_by('-date')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'События'
    
    # Стили для заголовков
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    
    # Заголовки
    headers = ['Название', 'Дата', 'Тимлидер', 'Треб. волонтёров', 'Длительность (ч)', 'Факт. волонтёров']
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Данные
    row_num = 2
    for event in events:
        participants_count = EventParticipant.objects.filter(event=event).count()
        ws.cell(row=row_num, column=1, value=event.name)
        ws.cell(row=row_num, column=2, value=event.date)
        ws.cell(row=row_num, column=3, value=f"{event.leader.first_name} {event.leader.last_name}")
        ws.cell(row=row_num, column=4, value=event.volunteers_required)
        ws.cell(row=row_num, column=5, value=event.duration_hours)
        ws.cell(row=row_num, column=6, value=participants_count)
        row_num += 1
    
    # Автоподбор ширины столбцов
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=all_events.xlsx'
    wb.save(response)
    return response

@team_leader_required
def generate_certificate(request, participant_id):
    """
    Генерирует благодарственное письмо на основе шаблона PowerPoint
    """
    try:
        participant = EventParticipant.objects.get(id=participant_id)
        event = participant.event
        scanner = participant.volunteer
        
        # Путь к шаблону
        template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
        
        # Открываем шаблон
        prs = Presentation(template_path)
        
        # Полное имя сканера для замены (в верхнем регистре)
        full_name = f"{scanner.first_name} {scanner.last_name}".upper()
        
        # Форматируем часы как округленное целое число
        hours_text = f"{round(participant.hours_awarded)}"
        
        # Заполняем данные
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # Заменяем плейсхолдеры на реальные данные
                            text = run.text
                            
                            # Заменяем NAME на имя сканера
                            if "NAME" in text:
                                run.text = text.replace("NAME", full_name)
                            
                            # Заменяем 00 на количество часов с форматированием
                            if "00" in text:
                                run.text = text.replace("00", hours_text)
                                run.font.name = "Impact"
                                run.font.size = Pt(25)
                            
                            # Заменяем другие плейсхолдеры
                            if "{{EVENT}}" in run.text:
                                run.text = run.text.replace("{{EVENT}}", event.name)
                            if "{{DATE}}" in run.text:
                                run.text = run.text.replace("{{DATE}}", event.date.strftime("%d.%m.%Y"))
                            if "{{LEADER}}" in run.text:
                                run.text = run.text.replace("{{LEADER}}", f"{event.leader.first_name} {event.leader.last_name}")
        
        # Сохраняем результат во временный файл
        temp_file = io.BytesIO()
        prs.save(temp_file)
        temp_file.seek(0)
        
        # Отдаем файл для скачивания
        filename = f"certificate_{scanner.last_name}_{event.name}.pptx"
        response = FileResponse(
            temp_file,
            as_attachment=True,
            filename=filename
        )
        return response
    
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=400)

@team_leader_required
def generate_all_certificates(request, event_id):
    """
    Генерирует благодарственные письма для всех участников мероприятия
    """
    try:
        event = Event.objects.get(id=event_id)
        participants = EventParticipant.objects.filter(event=event)
        
        # Создаем zip-архив с благодарностями
        import zipfile
        temp_zip = io.BytesIO()
        
        with zipfile.ZipFile(temp_zip, 'w') as zipf:
            for participant in participants:
                scanner = participant.volunteer
                
                # Путь к шаблону
                template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
                
                # Открываем шаблон
                prs = Presentation(template_path)
                
                # Полное имя сканера для замены (в верхнем регистре)
                full_name = f"{scanner.first_name} {scanner.last_name}".upper()
                
                # Форматируем часы как округленное целое число
                hours_text = f"{round(participant.hours_awarded)}"
                
                # Заполняем данные
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            text_frame = shape.text_frame
                            for paragraph in text_frame.paragraphs:
                                for run in paragraph.runs:
                                    # Заменяем плейсхолдеры на реальные данные
                                    text = run.text
                                    
                                    # Заменяем NAME на имя сканера
                                    if "NAME" in text:
                                        run.text = text.replace("NAME", full_name)
                                    
                                    # Заменяем 00 на количество часов с форматированием
                                    if "00" in text:
                                        run.text = text.replace("00", hours_text)
                                        run.font.name = "Impact"
                                        run.font.size = Pt(25)
                                    
                                    # Заменяем другие плейсхолдеры
                                    if "{{EVENT}}" in run.text:
                                        run.text = run.text.replace("{{EVENT}}", event.name)
                                    if "{{DATE}}" in run.text:
                                        run.text = run.text.replace("{{DATE}}", event.date.strftime("%d.%m.%Y"))
                                    if "{{LEADER}}" in run.text:
                                        run.text = run.text.replace("{{LEADER}}", f"{event.leader.first_name} {event.leader.last_name}")
                
                # Сохраняем во временный файл
                temp_pptx = io.BytesIO()
                prs.save(temp_pptx)
                temp_pptx.seek(0)
                
                # Добавляем в архив
                filename = f"certificate_{scanner.last_name}_{scanner.first_name}.pptx"
                zipf.writestr(filename, temp_pptx.getvalue())
        
        temp_zip.seek(0)
        
        # Отдаем архив для скачивания
        response = FileResponse(
            temp_zip,
            as_attachment=True,
            filename=f"certificates_{event.name}.zip"
        )
        return response
        
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=400)

@team_leader_required
def scanner_certificates(request):
    """
    Страница поиска сканеров для генерации благодарственных писем
    """
    return render(request, 'core/scanner_certificate.html')

@team_leader_required
def scanner_events_api(request, scanner_id):
    """
    API для получения списка всех мероприятий сканера
    """
    try:
        scanner = get_object_or_404(Scanner, id=scanner_id)
        participations = EventParticipant.objects.filter(volunteer=scanner).select_related('event')
        
        if not participations.exists():
            return JsonResponse({
                "events": [],
                "total_hours": 0,
                "first_date": "",
                "last_date": ""
            })
        
        # Получаем суммарные часы
        total_hours_query = participations.aggregate(
            total_hours=Sum('hours_awarded', output_field=FloatField())
        )
        total_hours = total_hours_query['total_hours'] or 0.0
        
        # Получаем даты первого и последнего мероприятия
        date_range = participations.aggregate(
            first_event=Min('event__date'),
            last_event=Max('event__date')
        )
        first_date = date_range['first_event']
        last_date = date_range['last_event']
        
        # Формируем список мероприятий
        events_list = []
        for p in participations:
            events_list.append({
                'name': p.event.name,
                'date': p.event.date.strftime("%d.%m.%Y"),
                'hours': float(p.hours_awarded)
            })
        
        return JsonResponse({
            "events": events_list,
            "total_hours": float(total_hours),
            "first_date": first_date.strftime("%d.%m.%Y"),
            "last_date": last_date.strftime("%d.%m.%Y")
        })
    
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=400)

@team_leader_required
def generate_scanner_certificate(request, scanner_id):
    """
    Генерирует благодарственное письмо сканеру за все мероприятия, в которых он участвовал
    """
    try:
        scanner = get_object_or_404(Scanner, id=scanner_id)
        
        # Получаем все мероприятия сканера
        participations = EventParticipant.objects.filter(volunteer=scanner).select_related('event')
        
        if not participations.exists():
            return JsonResponse({"error": "У сканера нет участия в мероприятиях"}, status=400)
        
        # Получаем суммарные часы
        total_hours_query = participations.aggregate(
            total_hours=Sum('hours_awarded', output_field=FloatField())
        )
        total_hours = total_hours_query['total_hours'] or 0.0
        
        # Получаем даты первого и последнего мероприятия
        date_range = participations.aggregate(
            first_event=Min('event__date'),
            last_event=Max('event__date')
        )
        first_date = date_range['first_event']
        last_date = date_range['last_event']
        
        # Формируем список мероприятий
        events_list = []
        for p in participations:
            events_list.append({
                'name': p.event.name,
                'date': p.event.date.strftime("%d.%m.%Y"),
                'hours': p.hours_awarded
            })
        
        # Путь к шаблону
        template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
        
        # Проверяем наличие файла
        if not os.path.exists(template_path):
            return JsonResponse({"error": f"Файл шаблона не найден по пути {template_path}"}, status=404)
        
        # Открываем шаблон
        prs = Presentation(template_path)
        
        # Полное имя сканера для замены (КАПС для именения в верхнем регистре)
        full_name = f"{scanner.first_name} {scanner.last_name}".upper()
        
        # Зеленый цвет Freedom Bank
        freedom_green = RGBColor(65, 174, 60)  # Зеленый цвет Freedom Bank
        
        # Период участия
        period_text = f"{first_date.strftime('%d.%m.%Y')} - {last_date.strftime('%d.%m.%Y')}"
        
        # Форматируем суммарные часы как округленное целое число
        total_hours_text = f"{round(total_hours)}"
        
        # Создаем строку списка мероприятий
        events_text = ""
        for event in events_list:
            events_text += f"• {event['name']} ({event['date']}): {round(event['hours'])} ч.\n"
        
        # Обходим все слайды и заменяем текст
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            # Заменяем NAME NAME NAME на имя сканера
                            if "NAME" in run.text:
                                run.text = run.text.replace("NAME", full_name)
                            
                            # Заменяем 00 на суммарное количество часов с форматированием
                            if "00" in run.text:
                                run.text = run.text.replace("00", total_hours_text)
                                run.font.name = "Impact"
                                run.font.size = Pt(25)
                            
                            # Добавляем информацию о периоде участия
                            if "PERIOD" in run.text:
                                run.text = run.text.replace("PERIOD", period_text)
                            
                            # Заменяем EVENT_LIST на список мероприятий
                            if "EVENT_LIST" in run.text:
                                run.text = run.text.replace("EVENT_LIST", events_text)
        
        # Сохраняем результат во временный файл
        temp_file = io.BytesIO()
        prs.save(temp_file)
        temp_file.seek(0)
        
        # Отдаем файл для скачивания
        filename = f"certificate_{scanner.last_name}_{scanner.first_name}.pptx"
        response = FileResponse(
            temp_file,
            as_attachment=True,
            filename=filename
        )
        return response
    
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        return JsonResponse({"error": str(e), "trace": trace}, status=400)

# Добавляем функцию для генерации сертификатов всем сканерам
@team_leader_required
def generate_all_scanner_certificates(request):
    """
    Генерирует благодарственные письма для всех сканеров
    """
    try:
        # Получаем всех сканеров, которые участвовали хотя бы в одном мероприятии
        scanners = Scanner.objects.filter(id__in=EventParticipant.objects.values_list('volunteer', flat=True).distinct())
        
        # Создаем zip-архив с благодарностями
        import zipfile
        temp_zip = io.BytesIO()
        
        with zipfile.ZipFile(temp_zip, 'w') as zipf:
            for scanner in scanners:
                # Получаем все мероприятия сканера
                participations = EventParticipant.objects.filter(volunteer=scanner).select_related('event')
                
                if not participations.exists():
                    continue
                
                # Получаем суммарные часы
                total_hours = participations.aggregate(total_hours=Coalesce(Sum('hours_awarded'), 0))['total_hours']
                
                # Получаем даты первого и последнего мероприятия
                date_range = participations.aggregate(
                    first_event=Min('event__date'),
                    last_event=Max('event__date')
                )
                first_date = date_range['first_event']
                last_date = date_range['last_event']
                
                # Формируем список мероприятий
                events_list = []
                for p in participations:
                    events_list.append({
                        'name': p.event.name,
                        'date': p.event.date.strftime("%d.%m.%Y"),
                        'hours': p.hours_awarded
                    })
                
                # Путь к шаблону
                template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
                
                # Открываем шаблон
                prs = Presentation(template_path)
                
                # Заполняем данные
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            text_frame = shape.text_frame
                            for paragraph in text_frame.paragraphs:
                                for run in paragraph.runs:
                                    text = run.text
                                    
                                    # Заменяем NAME NAME NAME на имя сканера
                                    if "NAME NAME NAME" in text:
                                        run.text = text.replace("NAME NAME NAME", f"{scanner.first_name} {scanner.last_name}")
                                    
                                    # Заменяем 00 на суммарное количество часов с форматированием
                                    if "00" in text:
                                        run.text = text.replace("00", f"{round(total_hours)}")
                                        run.font.name = "Impact"
                                        run.font.size = Pt(25)
                                    
                                    # Добавляем информацию о периоде участия
                                    if "PERIOD" in text:
                                        period_text = f"{first_date.strftime('%d.%m.%Y')} - {last_date.strftime('%d.%m.%Y')}"
                                        run.text = text.replace("PERIOD", period_text)
                                    
                                    # Заменяем EVENT_LIST на список мероприятий
                                    if "EVENT_LIST" in text:
                                        events_text = ""
                                        for event in events_list:
                                            events_text += f"• {event['name']} ({event['date']}): {round(event['hours'])} ч.\n"
                                        run.text = text.replace("EVENT_LIST", events_text)
                
                # Сохраняем во временный файл
                temp_pptx = io.BytesIO()
                prs.save(temp_pptx)
                temp_pptx.seek(0)
                
                # Добавляем в архив
                filename = f"certificate_{scanner.last_name}_{scanner.first_name}.pptx"
                zipf.writestr(filename, temp_pptx.getvalue())
        
        temp_zip.seek(0)
        
        # Отдаем архив для скачивания
        response = FileResponse(
            temp_zip,
            as_attachment=True,
            filename="all_scanner_certificates.zip"
        )
        return response
        
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=400)

@team_leader_required
def debug_template(request):
    """
    Диагностическая функция для анализа содержимого PPTX шаблона
    """
    try:
        template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
        prs = Presentation(template_path)
        
        result = []
        
        for i, slide in enumerate(prs.slides):
            slide_data = {
                "slide_number": i + 1,
                "shapes": []
            }
            
            for j, shape in enumerate(slide.shapes):
                shape_data = {
                    "shape_number": j + 1,
                    "shape_type": str(shape.shape_type),
                    "has_text": hasattr(shape, 'text_frame') and shape.has_text_frame,
                    "text_content": []
                }
                
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    for p_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                        for r_idx, run in enumerate(paragraph.runs):
                            shape_data["text_content"].append({
                                "paragraph": p_idx + 1,
                                "run": r_idx + 1,
                                "text": run.text
                            })
                
                slide_data["shapes"].append(shape_data)
            
            result.append(slide_data)
        
        return JsonResponse({"template_analysis": result})
    
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        return JsonResponse({"error": str(e), "trace": trace}, status=400)
