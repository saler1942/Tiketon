from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login
from django.contrib.auth.models import User
from django.core.mail import send_mail
from django.conf import settings
from django.http import HttpResponse, JsonResponse
import random
from django.contrib.auth.decorators import login_required, user_passes_test
from .models import Event, Volunteer, EventParticipant
from django.utils import timezone
from django.db.models import Q
import openpyxl
from openpyxl.utils import get_column_letter
from django.contrib import messages
from django.contrib.auth.models import Group
from dateutil import parser as dateparser
from django.core.paginator import Paginator
from openpyxl.styles import Font, PatternFill, Alignment

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
    volunteers = Volunteer.objects.all()
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
    return render(request, 'core/event_create.html', {'volunteers': volunteers})

@team_leader_required
def event_edit(request, event_id):
    event = Event.objects.get(id=event_id)
    all_volunteers = Volunteer.objects.all()
    participants = EventParticipant.objects.filter(event=event).select_related('volunteer')
    if request.method == 'POST':
        if 'add_volunteer' in request.POST:
            volunteer_id = request.POST.get('volunteer_id')
            if volunteer_id and not participants.filter(volunteer_id=volunteer_id).exists():
                EventParticipant.objects.create(event=event, volunteer_id=volunteer_id)
            return redirect('event_edit', event_id=event.id)
        if 'save_participants' in request.POST:
            for p in participants:
                is_late = request.POST.get(f'late_{p.id}') == 'on'
                late_minutes = int(request.POST.get(f'late_minutes_{p.id}', 0)) if is_late else 0
                p.is_late = is_late
                p.late_minutes = late_minutes
                p.save()
        if 'set_duration' in request.POST:
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
        'all_volunteers': all_volunteers,
        'participants': participants,
    })

def volunteer_search(request):
    q = request.GET.get('q', '').strip()
    event_id = request.GET.get('event_id')
    exclude_ids = list(EventParticipant.objects.filter(event_id=event_id).values_list('volunteer_id', flat=True)) if event_id else []
    if q:
        volunteers = Volunteer.objects.filter(
            Q(first_name__icontains=q) | Q(last_name__icontains=q) | Q(email__icontains=q)
        ).exclude(id__in=exclude_ids)[:10]
    else:
        volunteers = Volunteer.objects.none()
    data = [
        {'id': v.id, 'name': f'{v.first_name} {v.last_name}', 'email': v.email}
        for v in volunteers
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
        added = {'volunteers': 0, 'leaders': 0, 'events': 0, 'participants': 0}
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
                # Если строка только с ФИО — это участник текущего события
                elif current_event and row[0] and not any(row[1:]):
                    fio = row[0]
                    parts = str(fio).strip().split()
                    first_name = parts[0]
                    last_name = ' '.join(parts[1:]) if len(parts) > 1 else ''
                    email = f'{first_name.lower()}.{last_name.lower()}@tiketon.local'
                    volunteer = Volunteer.objects.filter(email=email).first()
                    if not volunteer:
                        volunteer = Volunteer.objects.create(first_name=first_name, last_name=last_name, email=email)
                    if not EventParticipant.objects.filter(event=current_event, volunteer=volunteer).exists():
                        EventParticipant.objects.create(event=current_event, volunteer=volunteer, hours_awarded=current_hours)
                        added['participants'] += 1
        messages.success(request, f"Импорт завершён. Волонтёров: {added['volunteers']}, Тимлидеров: {added['leaders']}, Мероприятий: {added['events']}, Участников: {added['participants']}")
        return redirect('events')
    return render(request, 'core/import_excel.html')

def export_all_events(request):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Мероприятия'
    # Шапка
    header = ['TEAM', 'События', 'Дата', 'Кол-во', 'Часы']
    ws.append(header)
    for col in range(1, 6):
        ws.cell(row=1, column=col).font = Font(bold=True)
        ws.cell(row=1, column=col).fill = PatternFill(start_color='FFF200', end_color='FFF200', fill_type='solid')
        ws.cell(row=1, column=col).alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(col)].width = 22
    row_idx = 2
    for event in Event.objects.all().order_by('-date'):
        leader = f'{event.leader.first_name} {event.leader.last_name}'
        name = event.name
        date = event.date.strftime('%d.%m.%y')
        count = event.volunteers_required
        hours = event.duration_hours
        # Строка мероприятия
        ws.append([leader, name, date, count, hours])
        for col in range(1, 6):
            ws.cell(row=row_idx, column=col).font = Font(bold=True)
            ws.cell(row=row_idx, column=col).fill = PatternFill(start_color='FFF200', end_color='FFF200', fill_type='solid')
            ws.cell(row=row_idx, column=col).alignment = Alignment(horizontal='center')
        row_idx += 1
        # Участники
        participants = EventParticipant.objects.filter(event=event).select_related('volunteer')
        for p in participants:
            fio = f'{p.volunteer.first_name} {p.volunteer.last_name}'
            ws.append([fio])
            row_idx += 1
        # Пустая строка между мероприятиями
        ws.append([])
        row_idx += 1
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=all_events.xlsx'
    wb.save(response)
    return response
