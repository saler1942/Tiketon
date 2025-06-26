from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login
from django.contrib.auth.models import User, Group
from django.contrib import messages
from django.http import JsonResponse, HttpResponse, FileResponse
from django.db.models import Q, Sum, Min, Max, FloatField
from django.db.models.functions import Coalesce
from django.conf import settings
import random
import string
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, Image, Spacer
from reportlab.lib import colors
from PIL import Image as PILImage, ImageDraw, ImageFont
import tempfile
import subprocess
import zipfile
from django.core.mail import send_mail
from django.contrib.auth.decorators import login_required, user_passes_test
from django.utils import timezone
from dateutil import parser as dateparser
from django.core.paginator import Paginator
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import sys
import platform
import time

from .models import Scanner, Event, EventParticipant

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
        start_date = request.POST.get('start_date')
        end_date = request.POST.get('end_date')
        volunteers_required = request.POST.get('volunteers_required')
        # Если указан период, используем его, иначе обычную дату
        if start_date and end_date:
            event = Event.objects.create(
                name=name,
                date=start_date,  # для обратной совместимости
                start_date=start_date,
                end_date=end_date,
                volunteers_required=volunteers_required,
                leader=request.user
            )
        else:
            event = Event.objects.create(
                name=name,
                date=date,
                start_date=date if date else None,
                end_date=date if date else None,
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
            if 'remove_participant' in request.POST and request.user.is_staff:
                participant_id = request.POST.get('participant_id')
                try:
                    participant = EventParticipant.objects.get(id=participant_id, event=event)
                    participant.delete()
                except EventParticipant.DoesNotExist:
                    pass
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

def save_pptx_as_png(pptx_data):
    """
    Функция-заглушка для работы с PPTX. На данный момент просто возвращает PPTX как есть,
    так как конвертация в PNG требует дополнительных инструментов.
    """
    # Возвращаем оригинальный PPTX
    return pptx_data.getvalue() if isinstance(pptx_data, io.BytesIO) else pptx_data, False

def convert_pptx_to_png(pptx_path):
    """
    Конвертирует PPTX в PNG используя автоматизацию PowerPoint (для Windows)
    или LibreOffice (для других ОС).
    """
    temp_png_path = pptx_path.replace('.pptx', '.png')
    
    if platform.system() == 'Windows':
        try:
            # Пробуем использовать COM-объект PowerPoint с правильной инициализацией
            import comtypes.client
            import pythoncom
            
            # Инициализация COM
            pythoncom.CoInitialize()
            
            # Создаем экземпляр PowerPoint
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = True
            
            # Полный путь к файлу (абсолютный)
            abs_pptx_path = os.path.abspath(pptx_path)
            presentation = powerpoint.Presentations.Open(abs_pptx_path)
            
            # Сохраняем первый слайд как PNG
            slide = presentation.Slides[0]
            abs_png_path = os.path.abspath(temp_png_path)
            slide.Export(abs_png_path, "PNG", 1024, 768)  # Указываем размер
            
            # Закрываем все
            presentation.Close()
            powerpoint.Quit()
            pythoncom.CoUninitialize()  # Освобождаем COM
            
            if os.path.exists(temp_png_path):
                with open(temp_png_path, 'rb') as f:
                    png_data = f.read()
                os.remove(temp_png_path)  # Удаляем временный файл
                return png_data, True
            
        except Exception as e:
            print(f"Ошибка при конвертации через PowerPoint: {e}")
            try:
                # В случае ошибки, попробуем освободить COM
                pythoncom.CoUninitialize()
            except:
                pass
    
    # Если PowerPoint не сработал, попробуем создать скриншот PPTX напрямую
    try:
        from PIL import Image, ImageDraw, ImageFont
        from pptx import Presentation
        
        # Открываем презентацию для анализа
        prs = Presentation(pptx_path)
        
        # Создаем изображение размером со слайд
        width, height = 1920, 1080  # Стандартный размер слайда 16:9
        img = Image.new("RGB", (width, height), (235, 199, 0))  # Желтый фон как на фото
        draw = ImageDraw.Draw(img)
        
        # Добавляем черные полосы сверху и снизу как на фото (кинолента)
        stripe_height = 50
        for i in range(0, width, 150):
            # Верхняя полоса
            draw.rectangle([(i, 0), (i+100, stripe_height)], fill=(30, 30, 30))
            # Нижняя полоса
            draw.rectangle([(i, height-stripe_height), (i+100, height)], fill=(30, 30, 30))
        
        # Извлекаем данные из PPTX
        # Так как у нас нет прямого доступа к информации, которая использовалась для заполнения,
        # мы будем искать текст в слайдах
        
        name = ""
        hours = ""
        event_name = ""
        event_date = ""
        
        # Проходим по всем слайдам и текстовым блокам
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            text = run.text
                            
                            # Здесь очень примерно определяем, что этот текст может быть
                            if len(text) > 0:
                                text = text.strip()
                                if text.isupper() and len(name) == 0 and len(text) > 5:
                                    name = text  # Вероятно это имя (в верхнем регистре)
                                elif text.isdigit() or (len(text) < 3 and text.replace('.', '').isdigit()):
                                    hours = text  # Вероятно это часы
                                elif "20" in text and ("." in text or "/" in text) and len(event_date) == 0:
                                    event_date = text  # Вероятно это дата
                                elif len(text) > 10 and len(event_name) == 0:
                                    event_name = text  # Вероятно это название мероприятия
        
        # Попробуем найти шрифты
        try:
            # В Windows пути к шрифтам
            fonts_dir = os.path.join(os.environ['SystemRoot'], 'Fonts')
            
            # Стандартные шрифты, которые должны быть в большинстве систем
            arial_path = os.path.join(fonts_dir, 'Arial.ttf')
            arialbd_path = os.path.join(fonts_dir, 'Arialbd.ttf')
            impact_path = os.path.join(fonts_dir, 'Impact.ttf')
            
            # Проверяем наличие шрифтов
            if os.path.exists(arialbd_path):
                name_font = ImageFont.truetype(arialbd_path, 60)
            else:
                name_font = ImageFont.load_default()
                
            if os.path.exists(impact_path):
                hours_font = ImageFont.truetype(impact_path, 80)
            else:
                hours_font = ImageFont.load_default()
                
            if os.path.exists(arial_path):
                normal_font = ImageFont.truetype(arial_path, 36)
            else:
                normal_font = ImageFont.load_default()
        except:
            # Если не нашли, используем стандартный
            name_font = ImageFont.load_default()
            hours_font = ImageFont.load_default()
            normal_font = ImageFont.load_default()
        
        # Фон для лучшей видимости
        draw.rectangle([(0, 0), (width, height)], fill="white")
        
        # Рисуем данные на изображении
        try:
            # Попробуем нарисовать билет в левом верхнем углу
            ticket_width, ticket_height = 200, 200
            ticket_x, ticket_y = 80, 150
            
            # Рамка билета
            draw.rectangle(
                [(ticket_x, ticket_y), (ticket_x + ticket_width, ticket_y + ticket_height)],
                outline=(0, 0, 0),
                width=3
            )
            
            # Текст на билете
            small_font = ImageFont.load_default()
            if os.path.exists(arial_path):
                small_font = ImageFont.truetype(arial_path, 20)
                
            draw.text((ticket_x + 10, ticket_y + 10), "TICKET", font=small_font, fill=(0, 0, 0))
            draw.text((ticket_x + 10, ticket_y + 40), "VIP", font=small_font, fill=(0, 0, 0))
            
            # Звезды вокруг билета
            for i in range(8):
                star_x = ticket_x + random.randint(0, ticket_width)
                star_y = ticket_y + random.randint(0, ticket_height)
                star_size = random.randint(10, 20)
                draw.text((star_x, star_y), "★", font=normal_font, fill=(0, 0, 0))
                
        except Exception as e:
            print(f"Ошибка при создании билета: {e}")
        
        # Основной заголовок "CERTIFICAT" большими буквами как на фото
        title = "CERTIFICAT"
        title_font_size = 150
        title_font = hours_font
        if os.path.exists(impact_path):
            try:
                title_font = ImageFont.truetype(impact_path, title_font_size)
            except:
                pass
            
        title_width = title_font.getlength(title) if hasattr(title_font, 'getlength') else len(title) * title_font_size // 2
        title_x = (width - title_width) // 2
        title_y = 100
        
        # Белый текст как на фото
        draw.text((title_x, title_y), title, font=title_font, fill=(255, 255, 255))
        
        # Подзаголовок "OF APPRECIATION"
        subtitle = "OF APPRECIATION"
        subtitle_font_size = 60
        subtitle_font = normal_font
        if os.path.exists(arial_path):
            try:
                subtitle_font = ImageFont.truetype(arial_path, subtitle_font_size)
            except:
                pass
        
        subtitle_width = subtitle_font.getlength(subtitle) if hasattr(subtitle_font, 'getlength') else len(subtitle) * subtitle_font_size // 2
        subtitle_x = width - subtitle_width - 100
        subtitle_y = title_y + title_font_size - 30
        
        draw.text((subtitle_x, subtitle_y), subtitle, font=subtitle_font, fill=(255, 255, 255))
        
        # Текст "This certificate is presented to:"
        presented_text = "This certificate is presented to:"
        presented_font_size = 36
        presented_font = normal_font
        if os.path.exists(arial_path):
            try:
                presented_font = ImageFont.truetype(arial_path, presented_font_size)
            except:
                pass
        
        draw.text((width//2, subtitle_y + 60), presented_text, font=presented_font, fill=(255, 255, 255))
        
        # Имя участника (большим зеленым шрифтом)
        if name:
            name_font_size = 100
            name_font = hours_font
            if os.path.exists(impact_path):
                try:
                    name_font = ImageFont.truetype(impact_path, name_font_size)
                except:
                    pass
            
            name_width = name_font.getlength(name) if hasattr(name_font, 'getlength') else len(name) * name_font_size // 2
            name_x = (width - name_width) // 2
            name_y = subtitle_y + 120
            
            # Зеленый текст как на фото
            draw.text((name_x, name_y), name, font=name_font, fill=(120, 220, 80))
        
        # Блок благодарности (как на скриншоте)
        lines = [
            "We, the Ticketon company, would like to sincerely express our gratitude and",
            "appreciation towards your incredible work and support in organizing our events in 2024 and 2025.",
            "You played important role in organization of each event.",
            "We hope to see you again in upcoming events!"
        ]
        line_h = italic_font.getbbox('Ag')[3] - italic_font.getbbox('Ag')[1] + 6
        for i, line in enumerate(lines):
            x = name_x
            words = line.split(' ')
            for word in words:
                if 'Ticketon' in word:
                    w, h = get_text_size(word, bold_italic_font)
                    draw.text((x, name_y + i * line_h), word, font=bold_italic_font, fill=(255,255,255,255))
                    x += w + italic_font.getbbox(' ')[2]
                else:
                    w, h = get_text_size(word, italic_font)
                    draw.text((x, name_y + i * line_h), word, font=italic_font, fill=(255,255,255,255))
                    x += w + italic_font.getbbox(' ')[2]
        
        # Рисуем стрелку с остриём только слева, справа — ровно
        arrow_w, arrow_height = 540, 110
        arrow = PILImage.new('RGBA', (arrow_w, arrow_height), (0,0,0,0))
        adraw = ImageDraw.Draw(arrow)
        points = [
            (0, arrow_height//2), (40, 0), (arrow_w, 0), (arrow_w, arrow_height), (40, arrow_height)
        ]
        adraw.polygon(points, fill=(76,175,80,255))

        margin = 18
        gap = 14
        gap_hours = 6
        section_w = (arrow_w - 2 * margin - 2 * gap) // 3
        center_y = arrow_height // 2

        # Шрифты для стрелки
        impact_25 = ImageFont.truetype(impact_path, size=25)
        impact_19_9 = ImageFont.truetype(impact_path, size=20)
        impact_18 = ImageFont.truetype(impact_path, size=18)

        # Левая секция: часы
        hours_text = f"{int(round(hours)):02d}"
        hw, hh = get_text_size(hours_text, impact_25)
        hlabel = "hours"
        hlw, hlh = get_text_size(hlabel, impact_18)
        line_w = section_w * 0.8
        total_h = hh + gap + hlh
        base_y = center_y - total_h // 2 - 10
        hours_x = margin + (section_w - hw) // 2
        hours_y = base_y
        adraw.text((hours_x, hours_y), hours_text, font=impact_25, fill=(255,255,255,255))
        line_y = hours_y + hh + gap
        line_x1 = margin + (section_w - line_w) // 2
        line_x2 = line_x1 + line_w
        adraw.line([(line_x1, line_y), (line_x2, line_y)], fill=(255,255,255,255), width=4)
        hlabel_x = margin + (section_w - hlw) // 2
        hlabel_y = line_y + gap_hours
        adraw.text((hlabel_x, hlabel_y), hlabel, font=impact_18, fill=(255,255,255,255))

        # Центральная секция: печать (крупнее, строго по центру секции)
        if os.path.exists(stamp_path):
            stamp = PILImage.open(stamp_path).convert('RGBA')
            stamp_size = int(arrow_height * 0.82)
            stamp = stamp.resize((stamp_size, stamp_size), PILImage.LANCZOS)
            stamp_x = margin + section_w + gap + (section_w - stamp_size)//2
            stamp_y = center_y - stamp_size//2
            arrow.paste(stamp, (stamp_x, stamp_y), stamp)

        # Правая секция: подпись и director (по центру секции, между ними больше расстояния)
        sign_text = "Torgunakova V. K."
        sign_font = impact_19_9
        sign_w, sign_h = get_text_size(sign_text, sign_font)
        dir_text = "director"
        dir_font = impact_18
        dir_w, dir_h = get_text_size(dir_text, dir_font)
        # Центрируем всю группу по вертикали секции
        group_h = sign_h + 12 + dir_h
        group_y = center_y - group_h // 2
        sign_x = margin + 2*section_w + 2*gap + (section_w - sign_w)//2
        sign_y = group_y
        adraw.text((sign_x, sign_y), sign_text, font=sign_font, fill=(0,0,0,255))
        dir_x = margin + 2*section_w + 2*gap + (section_w - dir_w)//2
        dir_y = sign_y + sign_h + 12
        adraw.text((dir_x, dir_y), dir_text, font=dir_font, fill=(0,0,0,255))

        arrow_x = width - arrow_w
        arrow_y = height - arrow_height - 60
        img.paste(arrow, (arrow_x, arrow_y), arrow)

        temp_dir = tempfile.mkdtemp()
        temp_img_path = os.path.join(temp_dir, 'cert.png')
        img.convert('RGB').save(temp_img_path, 'PNG')

        temp_pdf_path = os.path.join(temp_dir, 'certificate.pdf')
        c = canvas.Canvas(temp_pdf_path, pagesize=(width, height))
        c.drawImage(temp_img_path, 0, 0, width=width, height=height)
        c.save()

        with open(temp_pdf_path, 'rb') as f:
            pdf_data = f.read()
        os.remove(temp_img_path)
        os.remove(temp_pdf_path)
        os.rmdir(temp_dir)
        return pdf_data
    except Exception as e:
        print(f"Ошибка при создании PNG: {e}")
    
    # Пробуем через LibreOffice в последнюю очередь
    try:
        if platform.system() == 'Windows':
            # Если у нас есть LibreOffice, попробуем его использовать
            libreoffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            
            soffice_path = None
            for path in libreoffice_paths:
                if os.path.exists(path):
                    soffice_path = path
                    break
                
            if soffice_path:
                # Полный путь к файлу и директории
                abs_pptx_path = os.path.abspath(pptx_path)
                output_dir = os.path.dirname(abs_pptx_path)
                
                cmd = [
                    soffice_path, '--headless', 
                    '--convert-to', 'png', '--outdir', 
                    output_dir, abs_pptx_path
                ]
                subprocess.run(cmd, timeout=30, check=True)
            else:
                return None, False
        else:
            cmd = ['soffice', '--headless', '--convert-to', 'png', 
                  '--outdir', os.path.dirname(pptx_path), pptx_path]
            subprocess.run(cmd, timeout=30, check=True)
        
        # LibreOffice создаст файл с тем же именем, но с расширением .png
        if os.path.exists(temp_png_path):
            with open(temp_png_path, 'rb') as f:
                png_data = f.read()
            os.remove(temp_png_path)  # Удаляем временный файл
            return png_data, True
    except Exception as e:
        print(f"Ошибка при конвертации через LibreOffice: {e}")
    
    # Если все предыдущие попытки не удались, просто используем PPTX
    with open(pptx_path, 'rb') as f:
        pptx_data = f.read()
    return pptx_data, False

def get_certificate_from_template(name, hours, event_name=None, event_date=None, leader_name=None, period=None, events_list=None):
    """
    Создает PPTX на основе шаблона и заполняет данными
    """
    # Путь к шаблону
    template_path = os.path.join(settings.BASE_DIR, 'static', 'templates', 'certificate_template.pptx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Файл шаблона не найден по пути {template_path}")
    
    # Открываем шаблон
    try:
        prs = Presentation(template_path)
    except Exception as e:
        print(f"Ошибка при открытии шаблона: {e}")
        raise
    
    # Формируем строку списка мероприятий
    events_text = ""
    if events_list:
        for event in events_list:
            events_text += f"• {event['name']} ({event['date']}): {round(event['hours'])} ч.\n"
    
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
                            run.text = text.replace("NAME", name)
                        
                        # Заменяем 00 на количество часов с форматированием
                        if "00" in text:
                            run.text = text.replace("00", str(round(hours)))
                            run.font.name = "Impact"
                            run.font.size = Pt(25)
                        
                        # Заменяем EVENT_NAME на название мероприятия
                        if "EVENT_NAME" in text or "{{EVENT}}" in text:
                            if event_name:
                                run.text = text.replace("EVENT_NAME", event_name).replace("{{EVENT}}", event_name)
                        
                        # Заменяем EVENT_DATE на дату мероприятия
                        if "EVENT_DATE" in text or "{{DATE}}" in text:
                            if event_date:
                                run.text = text.replace("EVENT_DATE", event_date).replace("{{DATE}}", event_date)
                        
                        # Заменяем LEADER_NAME на имя лидера
                        if "LEADER_NAME" in text or "{{LEADER}}" in text:
                            if leader_name:
                                run.text = text.replace("LEADER_NAME", leader_name).replace("{{LEADER}}", leader_name)
                        
                        # Заменяем PERIOD на период
                        if "PERIOD" in text:
                            if period:
                                run.text = text.replace("PERIOD", period)
                        
                        # Заменяем EVENT_LIST на список мероприятий
                        if "EVENT_LIST" in text:
                            run.text = text.replace("EVENT_LIST", events_text)
    
    # Сохраняем во временный файл
    temp_dir = tempfile.mkdtemp()
    temp_pptx = os.path.join(temp_dir, "certificate.pptx")
    
    try:
        prs.save(temp_pptx)
        
        # Возвращаем путь к временному файлу
        return temp_pptx, temp_dir
    except Exception as e:
        print(f"Ошибка при сохранении PPTX: {e}")
        # Очищаем временную директорию при ошибке
        try:
            os.remove(temp_pptx)
            os.rmdir(temp_dir)
        except:
            pass
        raise

@team_leader_required
def generate_certificate(request, participant_id):
    try:
        participant = EventParticipant.objects.get(id=participant_id)
        event = participant.event
        scanner = participant.volunteer
        full_name = f"{scanner.first_name} {scanner.last_name}".upper()
        hours = round(participant.hours_awarded)
        event_name = event.name
        event_date = event.date.strftime("%d.%m.%Y")
        leader_name = f"{event.leader.first_name} {event.leader.last_name}" if event.leader else None
        file_data = create_certificate_pdf(full_name, hours, event_name, event_date, leader_name)
        participant.hours_awarded = 0
        participant.save()
        filename = f"certificate_{scanner.last_name}_{event.name}.pdf"
        user_agent = request.META.get('HTTP_USER_AGENT', '').lower()
        is_mobile = any(x in user_agent for x in ['iphone', 'android', 'ipad', 'mobile'])
        if is_mobile:
            import io, zipfile
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                zipf.writestr(filename, file_data)
            zip_buffer.seek(0)
            response = HttpResponse(zip_buffer.read(), content_type='application/zip')
            response['Content-Disposition'] = f'attachment; filename="{scanner.last_name}_{event.name}.zip"'
            return response
        response = HttpResponse(file_data, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
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
        temp_zip = io.BytesIO()
        
        with zipfile.ZipFile(temp_zip, 'w') as zipf:
            for participant in participants:
                scanner = participant.volunteer
                
                # Полное имя сканера для замены (в верхнем регистре)
                full_name = f"{scanner.first_name} {scanner.last_name}".upper()
                
                # Подготавливаем данные для сертификата
                event_name = event.name
                event_date = event.date.strftime("%d.%m.%Y")
                leader_name = f"{event.leader.first_name} {event.leader.last_name}" if event.leader else None
                hours = participant.hours_awarded
                
                # Создаем сертификат на основе шаблона PPTX
                temp_pptx_path, temp_dir = get_certificate_from_template(
                    name=full_name, 
                    hours=hours,
                    event_name=event_name,
                    event_date=event_date,
                    leader_name=leader_name
                )
                
                try:
                    # Конвертируем в PDF
                    pdf_data = convert_pptx_to_pdf(temp_pptx_path)
                    
                    # Файл для архива и его формат
                    file_data = None
                    extension = ""
                    
                    if pdf_data:
                        file_data = pdf_data
                        extension = 'pdf'
                    else:
                        # Если конвертация не удалась, используем PPTX
                        with open(temp_pptx_path, 'rb') as f:
                            file_data = f.read()
                        extension = 'pptx'
                    
                    # Обнуляем часы сканера при получении благодарственного письма
                    participant.hours_awarded = 0
                    participant.save()
                    
                    # Добавляем в архив
                    filename = f"certificate_{scanner.last_name}_{scanner.first_name}.{extension}"
                    zipf.writestr(filename, file_data)
                finally:
                    # Очищаем временные файлы
                    try:
                        os.remove(temp_pptx_path)
                        if os.path.exists(temp_pptx_path.replace('.pptx', '.pdf')):
                            os.remove(temp_pptx_path.replace('.pptx', '.pdf'))
                        os.rmdir(temp_dir)
                    except Exception as e:
                        print(f"Ошибка при удалении временных файлов: {e}")
        
        temp_zip.seek(0)
        
        # Отдаем архив для скачивания
        response = FileResponse(
            temp_zip,
            as_attachment=True,
            filename=f"certificates_{event.name}.zip"
        )
        return response
        
    except Exception as e:
        import traceback
        print(f"Ошибка при генерации сертификатов: {e}")
        print(traceback.format_exc())
        return JsonResponse({"error": str(e)}, status=400)

@team_leader_required
def generate_scanner_certificate(request, scanner_id):
    try:
        scanner = get_object_or_404(Scanner, id=scanner_id)
        participations = EventParticipant.objects.filter(volunteer=scanner).select_related('event')
        if not participations.exists():
            return JsonResponse({"error": "Сканер еще не участвовал на мероприятиях"}, status=400)
        total_hours = participations.aggregate(total_hours=Sum('hours_awarded', output_field=FloatField()))['total_hours'] or 0.0
        if total_hours == 0:
            return JsonResponse({"error": "Сканер уже получил благодарственное письмо и у него нет часов"}, status=400)
        date_range = participations.aggregate(first_event=Min('event__date'), last_event=Max('event__date'))
        first_date = date_range['first_event']
        last_date = date_range['last_event']
        period_text = f"{first_date.strftime('%d.%m.%Y')} - {last_date.strftime('%d.%m.%Y')}"
        events_list = [{
            'name': p.event.name,
            'date': p.event.date.strftime("%d.%m.%Y"),
            'hours': p.hours_awarded
        } for p in participations]
        full_name = f"{scanner.first_name} {scanner.last_name}".upper()
        hours = round(total_hours)
        file_data = create_certificate_pdf(full_name, hours, period=period_text, events_list=events_list)
        for participant in participations:
            participant.hours_awarded = 0
            participant.save()
        filename = f"certificate_{scanner.last_name}_{scanner.first_name}.pdf"
        user_agent = request.META.get('HTTP_USER_AGENT', '').lower()
        is_mobile = any(x in user_agent for x in ['iphone', 'android', 'ipad', 'mobile'])
        if is_mobile:
            import io, zipfile
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                zipf.writestr(filename, file_data)
            zip_buffer.seek(0)
            response = HttpResponse(zip_buffer.read(), content_type='application/zip')
            response['Content-Disposition'] = f'attachment; filename="{scanner.last_name}_{scanner.first_name}.zip"'
            return response
        response = HttpResponse(file_data, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        return JsonResponse({"error": str(e), "trace": trace}, status=400)

@team_leader_required
def generate_all_scanner_certificates(request):
    """
    Генерирует благодарственные письма для всех сканеров
    """
    try:
        # Получаем всех сканеров, которые участвовали хотя бы в одном мероприятии
        scanners = Scanner.objects.filter(id__in=EventParticipant.objects.values_list('volunteer', flat=True).distinct())
        
        # Создаем zip-архив с благодарностями
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
                
                # Полное имя сканера (в верхнем регистре)
                full_name = f"{scanner.first_name} {scanner.last_name}".upper()
                
                # Период участия
                period_text = f"{first_date.strftime('%d.%m.%Y')} - {last_date.strftime('%d.%m.%Y')}"
                
                # Создаем сертификат на основе шаблона PPTX
                temp_pptx_path, temp_dir = get_certificate_from_template(
                    name=full_name,
                    hours=total_hours,
                    period=period_text,
                    events_list=events_list
                )
                
                try:
                    # Конвертируем в PDF
                    pdf_data = convert_pptx_to_pdf(temp_pptx_path)
                    
                    # Файл для архива и его формат
                    file_data = None
                    extension = ""
                    
                    if pdf_data:
                        file_data = pdf_data
                        extension = 'pdf'
                    else:
                        # Если конвертация не удалась, используем PPTX
                        with open(temp_pptx_path, 'rb') as f:
                            file_data = f.read()
                        extension = 'pptx'
                    
                    # Обнуляем часы сканера при получении благодарственного письма
                    for p in participations:
                        p.hours_awarded = 0
                        p.save()
                    
                    # Добавляем в архив
                    filename = f"certificate_{scanner.last_name}_{scanner.first_name}.{extension}"
                    zipf.writestr(filename, file_data)
                finally:
                    # Очищаем временные файлы
                    try:
                        os.remove(temp_pptx_path)
                        if os.path.exists(temp_pptx_path.replace('.pptx', '.pdf')):
                            os.remove(temp_pptx_path.replace('.pptx', '.pdf'))
                        os.rmdir(temp_dir)
                    except Exception as e:
                        print(f"Ошибка при удалении временных файлов: {e}")
        
        temp_zip.seek(0)
        
        # Отдаем архив для скачивания
        response = FileResponse(
            temp_zip,
            as_attachment=True,
            filename=f"all_scanner_certificates.zip"
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

@team_leader_required
def all_scanners_list(request):
    """
    Отображает список всех сканеров с возможностью фильтрации и поиска
    """
    query = request.GET.get('q', '')
    filter_min_hours = request.GET.get('min_hours', '')
    filter_max_hours = request.GET.get('max_hours', '')
    
    # Базовый запрос со всеми сканерами
    scanners = Scanner.objects.all().order_by('last_name', 'first_name')
    
    # Применяем фильтр по имени/email
    if query:
        scanners = scanners.filter(
            Q(first_name__icontains=query) | 
            Q(last_name__icontains=query) | 
            Q(email__icontains=query)
        )
    
    # Добавляем информацию о часах для каждого сканера
    scanners_with_hours = []
    for scanner in scanners:
        # Получаем суммарные часы для сканера
        total_hours = EventParticipant.objects.filter(
            volunteer=scanner
        ).aggregate(
            total_hours=Coalesce(Sum('hours_awarded'), 0.0)
        )['total_hours']
        
        # Применяем фильтры по часам
        if filter_min_hours and float(filter_min_hours) > total_hours:
            continue
        
        if filter_max_hours and float(filter_max_hours) < total_hours:
            continue
        
        # Получаем список мероприятий, в которых участвовал сканер
        participations = EventParticipant.objects.filter(
            volunteer=scanner
        ).select_related('event').order_by('-event__date')
        
        events = []
        for p in participations:
            events.append({
                'id': p.event.id,
                'name': p.event.name,
                'date': p.event.date.strftime('%d.%m.%Y'),
                'hours': p.hours_awarded
            })
        
        scanners_with_hours.append({
            'id': scanner.id,
            'first_name': scanner.first_name,
            'last_name': scanner.last_name,
            'email': scanner.email,
            'total_hours': total_hours,
            'events': events
        })
    
    # Сортируем по убыванию часов
    scanners_with_hours = sorted(scanners_with_hours, key=lambda x: x['total_hours'], reverse=True)
    
    context = {
        'scanners': scanners_with_hours,
        'query': query,
        'min_hours': filter_min_hours,
        'max_hours': filter_max_hours
    }
    
    return render(request, 'core/all_scanners.html', context)

def create_certificate_pdf(name, hours, event_name=None, event_date=None, leader_name=None, period=None, events_list=None):
    from PIL import Image as PILImage, ImageDraw, ImageFont
    from reportlab.pdfgen import canvas
    import tempfile, os

    base_dir = settings.BASE_DIR
    bg_path = os.path.join(base_dir, 'static', 'templates', 'background.png')
    overlay_path = os.path.join(base_dir, 'static', 'templates', 'back_black.png')
    logo_path = os.path.join(base_dir, 'static', 'templates', 'image.png')
    stamp_path = os.path.join(base_dir, 'static', 'templates', 'stamp.png')
    impact_path = os.path.join(base_dir, 'static', 'fonts', 'IMPACT.TTF')
    sans_italic = os.path.join(base_dir, 'static', 'fonts', 'OpenSans-Italic.ttf')
    sans_bold_italic = os.path.join(base_dir, 'static', 'fonts', 'OpenSans-BoldItalic.ttf')

    width, height = 1123, 794
    bg = PILImage.open(bg_path).convert('RGBA').resize((width, height))
    overlay = PILImage.open(overlay_path).convert('RGBA').resize((width, height))
    img = PILImage.alpha_composite(bg, overlay)
    draw = ImageDraw.Draw(img)

    # Логотип в левый нижний угол (logo_scale = 0.52)
    logo = PILImage.open(logo_path).convert('RGBA')
    logo_w, logo_h = logo.size
    logo_scale = 0.52
    logo = logo.resize((int(logo_w * logo_scale), int(logo_h * logo_scale)), PILImage.LANCZOS)
    img.paste(logo, (40, height - logo.height - 40), logo)

    # Шрифты
    impact = ImageFont.truetype(impact_path, size=98)
    impact_name = ImageFont.truetype(impact_path, size=64)
    sans_bold_italic_f = ImageFont.truetype(sans_bold_italic, size=24)
    sans_italic_f = ImageFont.truetype(sans_italic, size=18)
    sans_italic_f_small = ImageFont.truetype(sans_italic, size=16)
    sans_bold_italic_f_small = ImageFont.truetype(sans_bold_italic, size=18)

    def get_text_size(text, font):
        bbox = font.getbbox(text)
        return bbox[2] - bbox[0], bbox[3] - bbox[1]

    # CERTIFICAT
    cert_text = "CERTIFICAT"
    cert_w, cert_h = get_text_size(cert_text, impact)
    cert_x = width - cert_w - 80
    cert_y = 70
    draw.text((cert_x, cert_y), cert_text, font=impact, fill=(255,255,255,255))

    # OF APPRECIATION
    app_text = "OF APPRECIATION"
    app_w, app_h = get_text_size(app_text, sans_bold_italic_f)
    app_x = width - app_w - 85
    app_y = cert_y + cert_h + 40
    draw.text((app_x, app_y), app_text, font=sans_bold_italic_f, fill=(255,255,255,255))

    # This certificate is presented to:
    pres_text = "This certificate is presented to:"
    pres_w, pres_h = get_text_size(pres_text, sans_italic_f)
    pres_x = width - pres_w - 120
    pres_y = app_y + app_h + 40
    draw.text((pres_x, pres_y), pres_text, font=sans_italic_f, fill=(255,255,255,255))

    # NAME
    name_text = name
    name_w, name_h = get_text_size(name_text, impact_name)
    name_x = width - name_w - 120
    name_y = pres_y + pres_h + 10
    draw.text((name_x, name_y), name_text, font=impact_name, fill=(76,175,80,255))

    # Блок благодарности (как на скриншоте)
    thanks_text = (
        "We, the Ticketon company, would like to sincerely express our gratitude and appreciation "
        "towards your incredible work and support in organizing our events in 2024 and 2025. "
        "You played important role in organization of each event. We hope to see you again in upcoming_events!"
    )
    italic_font = ImageFont.truetype(sans_italic, size=22)
    bold_italic_font = ImageFont.truetype(sans_bold_italic, size=22)
    underline_font = italic_font
    start_x = name_x
    max_width = 600
    start_y = name_y + name_h + 40
    # Разбиваем на слова, чтобы upcoming_events! не переносился
    words = thanks_text.split(' ')
    x, y = start_x, start_y
    line_h = italic_font.getbbox('Ag')[3] - italic_font.getbbox('Ag')[1] + 6
    def draw_word(word, font, color, underline=False):
        nonlocal x, y
        display_word = word.replace('upcoming_events!', 'upcoming events!')
        w, h = get_text_size(display_word, font)
        if x + w > start_x + max_width:
            x = start_x
            y += line_h
        draw.text((x, y), display_word, font=font, fill=color)
        if underline:
            uy = y + h + 2
            draw.line([(x, uy), (x + w, uy)], fill=color, width=2)
        x += w + italic_font.getbbox(' ')[2]
    for word in words:
        if 'Ticketon' in word:
            draw_word(word, bold_italic_font, (255,255,255,255))
        else:
            draw_word(word, italic_font, (255,255,255,255))
    
    # Рисуем стрелку с остриём только слева, справа — ровно
    arrow_w, arrow_height = 540, 110
    arrow = PILImage.new('RGBA', (arrow_w, arrow_height), (0,0,0,0))
    adraw = ImageDraw.Draw(arrow)
    points = [
        (0, arrow_height//2), (40, 0), (arrow_w, 0), (arrow_w, arrow_height), (40, arrow_height)
    ]
    adraw.polygon(points, fill=(76,175,80,255))

    margin = 18
    gap = 14
    gap_hours = 6
    section_w = (arrow_w - 2 * margin - 2 * gap) // 3
    center_y = arrow_height // 2

    # Шрифты для стрелки
    impact_25 = ImageFont.truetype(impact_path, size=25)
    impact_19_9 = ImageFont.truetype(impact_path, size=20)
    impact_18 = ImageFont.truetype(impact_path, size=18)

    # Левая секция: часы
    hours_text = f"{int(round(hours)):02d}"
    hw, hh = get_text_size(hours_text, impact_25)
    hlabel = "hours"
    hlw, hlh = get_text_size(hlabel, impact_18)
    line_w = section_w * 0.8
    total_h = hh + gap + hlh
    base_y = center_y - total_h // 2 - 10
    hours_x = margin + (section_w - hw) // 2
    hours_y = base_y
    adraw.text((hours_x, hours_y), hours_text, font=impact_25, fill=(255,255,255,255))
    line_y = hours_y + hh + gap
    line_x1 = margin + (section_w - line_w) // 2
    line_x2 = line_x1 + line_w
    adraw.line([(line_x1, line_y), (line_x2, line_y)], fill=(255,255,255,255), width=4)
    hlabel_x = margin + (section_w - hlw) // 2
    hlabel_y = line_y + gap_hours
    adraw.text((hlabel_x, hlabel_y), hlabel, font=impact_18, fill=(255,255,255,255))

    # Центральная секция: печать (крупнее, строго по центру секции)
    if os.path.exists(stamp_path):
        stamp = PILImage.open(stamp_path).convert('RGBA')
        stamp_size = int(arrow_height * 0.82)
        stamp = stamp.resize((stamp_size, stamp_size), PILImage.LANCZOS)
        stamp_x = margin + section_w + gap + (section_w - stamp_size)//2
        stamp_y = center_y - stamp_size//2
        arrow.paste(stamp, (stamp_x, stamp_y), stamp)

    # Правая секция: подпись и director (по центру секции, между ними больше расстояния)
    sign_text = "Torgunakova V. K."
    sign_font = impact_19_9
    sign_w, sign_h = get_text_size(sign_text, sign_font)
    dir_text = "director"
    dir_font = impact_18
    dir_w, dir_h = get_text_size(dir_text, dir_font)
    # Центрируем всю группу по вертикали секции
    group_h = sign_h + 12 + dir_h
    group_y = center_y - group_h // 2
    sign_x = margin + 2*section_w + 2*gap + (section_w - sign_w)//2
    sign_y = group_y
    adraw.text((sign_x, sign_y), sign_text, font=sign_font, fill=(0,0,0,255))
    dir_x = margin + 2*section_w + 2*gap + (section_w - dir_w)//2
    dir_y = sign_y + sign_h + 12
    adraw.text((dir_x, dir_y), dir_text, font=dir_font, fill=(0,0,0,255))

    arrow_x = width - arrow_w
    arrow_y = height - arrow_height - 60
    img.paste(arrow, (arrow_x, arrow_y), arrow)

    temp_dir = tempfile.mkdtemp()
    temp_img_path = os.path.join(temp_dir, 'cert.png')
    img.convert('RGB').save(temp_img_path, 'PNG')

    temp_pdf_path = os.path.join(temp_dir, 'certificate.pdf')
    c = canvas.Canvas(temp_pdf_path, pagesize=(width, height))
    c.drawImage(temp_img_path, 0, 0, width=width, height=height)
    c.save()

    with open(temp_pdf_path, 'rb') as f:
        pdf_data = f.read()
    os.remove(temp_img_path)
    os.remove(temp_pdf_path)
    os.rmdir(temp_dir)
    return pdf_data

def convert_pptx_to_pdf(pptx_path):
    """
    Конвертирует PPTX в PDF используя PowerPoint через COM-интерфейс (Windows)
    """
    pdf_path = pptx_path.replace('.pptx', '.pdf')
    
    # Определяем константы для PowerPoint
    try:
        import comtypes.client
        import pythoncom
        
        # Инициализация COM
        pythoncom.CoInitialize()
        
        # Форматы для экспорта (константы для PowerPoint)
        ppSaveAsPDF = 32  # Формат PDF
        
        # Создаем экземпляр PowerPoint
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = True  # PowerPoint должен быть видимым из-за ограничений безопасности
        
        # Полный абсолютный путь к файлам
        abs_pptx_path = os.path.abspath(pptx_path)
        abs_pdf_path = os.path.abspath(pdf_path)
        
        print(f"Открываем презентацию: {abs_pptx_path}")
        
        # Открываем презентацию
        presentation = powerpoint.Presentations.Open(abs_pptx_path)
        
        # Небольшая пауза для загрузки презентации
        time.sleep(2)
        
        try:
            print(f"Сохраняем как PDF: {abs_pdf_path}")
            # Сохраняем как PDF
            presentation.SaveAs(abs_pdf_path, ppSaveAsPDF)
            
            # Пауза для сохранения PDF
            time.sleep(2)
            
            # Закрываем презентацию
            presentation.Close()
        except Exception as e:
            print(f"Ошибка при сохранении PDF: {e}")
        finally:
            # Закрываем PowerPoint
            powerpoint.Quit()
        
        # Освобождаем COM-объекты
        pythoncom.CoUninitialize()
        
        # Проверяем, создан ли PDF-файл
        if os.path.exists(pdf_path):
            print(f"PDF файл успешно создан: {pdf_path}")
            with open(pdf_path, 'rb') as f:
                pdf_data = f.read()
            return pdf_data
        else:
            print(f"PDF-файл не создан по пути {pdf_path}")
            return None
    except Exception as e:
        import traceback
        print(f"Ошибка при конвертации PPTX в PDF: {e}")
        print(traceback.format_exc())
        return None

@team_leader_required
def event_delete(request, event_id):
    event = get_object_or_404(Event, id=event_id)
    if not request.user.is_staff:
        messages.error(request, 'Удаление доступно только администраторам.')
        return redirect('event_edit', event_id=event.id)
    if request.method == 'POST':
        event.delete()
        messages.success(request, 'Мероприятие успешно удалено.')
        return redirect('events')
    return render(request, 'core/event_confirm_delete.html', {'event': event})
