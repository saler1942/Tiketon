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
import zipfile
from django.core.mail import send_mail
from django.contrib.auth.decorators import login_required, user_passes_test
from django.utils import timezone
from dateutil import parser as dateparser
from django.core.paginator import Paginator

from .models import Scanner, Event, EventParticipant

# Проверка доступа (только тимлидеры и админы)
def is_team_leader_or_admin(user):
    return user.is_staff or user.is_team_leader

def team_leader_required(view_func):
    def wrapper(request, *args, **kwargs):
        if not request.user.is_authenticated or not is_team_leader_or_admin(request.user):
            return redirect('login')
        return view_func(request, *args, **kwargs)
    return wrapper

def generate_certificate_pdf(name, hours, event_name=None, event_date=None, leader_name=None, period=None, events_list=None):
    """
    Генерирует PDF благодарственного письма
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.5*inch, rightMargin=0.5*inch)
    
    elements = []
    
    # Стили
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name='Title', 
        parent=styles['Title'], 
        fontSize=24, 
        alignment=1, 
        spaceAfter=20
    )
    subtitle_style = ParagraphStyle(
        name='Subtitle', 
        parent=styles['Normal'], 
        fontSize=16, 
        alignment=1, 
        spaceAfter=20
    )
    normal_style = ParagraphStyle(
        name='Normal', 
        parent=styles['Normal'], 
        fontSize=12, 
        alignment=1, 
        spaceAfter=12
    )
    hours_style = ParagraphStyle(
        name='Hours', 
        parent=styles['Normal'], 
        fontSize=20, 
        fontName='Helvetica-Bold', 
        alignment=1, 
        spaceAfter=20
    )
    
    # Добавление логотипа, если есть
    logo_path = os.path.join(settings.BASE_DIR, 'static', 'img', 'freedom_logo.jpg')
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=2*inch, height=1*inch)
        elements.append(logo)
        elements.append(Spacer(1, 0.25*inch))
    
    # Заголовок и благодарность
    elements.append(Paragraph('БЛАГОДАРНОСТЬ', title_style))
    elements.append(Paragraph(f'Выражаем благодарность <b>{name}</b>', subtitle_style))
    
    # Информация о мероприятии или общая информация
    if event_name:
        elements.append(Paragraph(f'За волонтерство на мероприятии <b>"{event_name}"</b>', normal_style))
        if event_date:
            elements.append(Paragraph(f'Дата проведения: {event_date}', normal_style))
    else:
        elements.append(Paragraph('За активное участие в волонтерской деятельности', normal_style))
        if period:
            elements.append(Paragraph(f'В период: {period}', normal_style))
    
    # Часы
    elements.append(Paragraph(f'Волонтерские часы: <b>{hours}</b>', hours_style))
    
    # Список мероприятий если есть
    if events_list:
        events_table_data = [['Мероприятие', 'Дата', 'Часы']]
        for event in events_list:
            events_table_data.append([event['name'], event['date'], str(round(event['hours']))])
        
        events_table = Table(events_table_data, colWidths=[4*inch, 2*inch, 1*inch])
        events_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(Spacer(1, 0.25*inch))
        elements.append(events_table)
    
    # Подпись
    if leader_name:
        elements.append(Spacer(1, 0.5*inch))
        elements.append(Paragraph(f'Тимлидер: {leader_name}', normal_style))
    
    # Дата
    elements.append(Spacer(1, 0.25*inch))
    elements.append(Paragraph(f'Дата выдачи: {datetime.now().strftime("%d.%m.%Y")}', normal_style))
    
    # Дополнительная информация
    elements.append(Spacer(1, 0.5*inch))
    elements.append(Paragraph('FREEDOM Волонтерский центр', styles["Heading3"]))
    
    # Построение PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

def generate_certificate_png(name, hours, event_name=None, event_date=None, leader_name=None, period=None, events_list=None):
    """
    Генерирует PNG благодарственного письма
    """
    # Создаем изображение
    width, height = 1200, 842  # A4 в ландшафтной ориентации
    image = PILImage.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(image)
    
    # Загружаем шрифты
    try:
        font_path = os.path.join(settings.BASE_DIR, 'static', 'fonts')
        if not os.path.exists(font_path):
            os.makedirs(font_path)
        
        # Проверяем наличие шрифтов, если их нет - используем стандартные
        title_font_path = os.path.join(font_path, 'arial_bold.ttf')
        normal_font_path = os.path.join(font_path, 'arial.ttf')
        
        if not os.path.exists(title_font_path):
            # Используем системный шрифт
            import platform
            if platform.system() == "Windows":
                title_font_path = "C:\\Windows\\Fonts\\arialbd.ttf"
                normal_font_path = "C:\\Windows\\Fonts\\arial.ttf"
            else:
                # Для Linux/Mac используем стандартный шрифт
                title_font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
                normal_font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
        
        title_font = ImageFont.truetype(title_font_path, size=40)
        subtitle_font = ImageFont.truetype(title_font_path, size=30)
        normal_font = ImageFont.truetype(normal_font_path, size=20)
        hours_font = ImageFont.truetype(title_font_path, size=36)
    except Exception:
        # Если не удалось загрузить шрифты, используем стандартный шрифт
        title_font = ImageFont.load_default()
        subtitle_font = ImageFont.load_default()
        normal_font = ImageFont.load_default()
        hours_font = ImageFont.load_default()
    
    # Добавляем логотип
    logo_path = os.path.join(settings.BASE_DIR, 'static', 'img', 'freedom_logo.jpg')
    if os.path.exists(logo_path):
        try:
            logo = PILImage.open(logo_path)
            logo = logo.resize((200, 100))
            image.paste(logo, (width // 2 - 100, 50))
        except Exception:
            pass
    
    # Заголовок
    draw.text((width // 2, 180), "БЛАГОДАРНОСТЬ", fill="black", font=title_font, anchor="mm")
    
    # Имя
    draw.text((width // 2, 250), f"Выражаем благодарность {name}", fill="black", font=subtitle_font, anchor="mm")
    
    # Информация о мероприятии
    y_pos = 320
    if event_name:
        draw.text((width // 2, y_pos), f'За волонтерство на мероприятии "{event_name}"', fill="black", font=normal_font, anchor="mm")
        y_pos += 40
        if event_date:
            draw.text((width // 2, y_pos), f"Дата проведения: {event_date}", fill="black", font=normal_font, anchor="mm")
            y_pos += 40
    else:
        draw.text((width // 2, y_pos), "За активное участие в волонтерской деятельности", fill="black", font=normal_font, anchor="mm")
        y_pos += 40
        if period:
            draw.text((width // 2, y_pos), f"В период: {period}", fill="black", font=normal_font, anchor="mm")
            y_pos += 40
    
    # Часы
    draw.text((width // 2, y_pos + 30), f"Волонтерские часы: {hours}", fill="black", font=hours_font, anchor="mm")
    y_pos += 100
    
    # Список мероприятий
    if events_list:
        draw.text((width // 2, y_pos), "Список мероприятий:", fill="black", font=normal_font, anchor="mm")
        y_pos += 40
        
        # Добавляем таблицу с мероприятиями
        line_height = 25
        start_x = 200
        col_widths = [500, 200, 100]
        
        # Заголовок таблицы
        draw.text((start_x + col_widths[0]//2, y_pos), "Мероприятие", fill="black", font=normal_font, anchor="mm")
        draw.text((start_x + col_widths[0] + col_widths[1]//2, y_pos), "Дата", fill="black", font=normal_font, anchor="mm")
        draw.text((start_x + col_widths[0] + col_widths[1] + col_widths[2]//2, y_pos), "Часы", fill="black", font=normal_font, anchor="mm")
        y_pos += line_height
        
        # Горизонтальная линия после заголовка
        draw.line((start_x, y_pos, start_x + sum(col_widths), y_pos), fill="black", width=1)
        y_pos += 5
        
        # Данные таблицы
        for event in events_list[:8]:  # Ограничиваем до 8 мероприятий
            draw.text((start_x + col_widths[0]//2, y_pos), event['name'], fill="black", font=normal_font, anchor="mm")
            draw.text((start_x + col_widths[0] + col_widths[1]//2, y_pos), event['date'], fill="black", font=normal_font, anchor="mm")
            draw.text((start_x + col_widths[0] + col_widths[1] + col_widths[2]//2, y_pos), str(round(event['hours'])), fill="black", font=normal_font, anchor="mm")
            y_pos += line_height
    
    # Подпись
    y_pos = height - 180
    if leader_name:
        draw.text((width // 2, y_pos), f"Тимлидер: {leader_name}", fill="black", font=normal_font, anchor="mm")
        y_pos += 40
    
    # Дата
    draw.text((width // 2, y_pos), f"Дата выдачи: {datetime.now().strftime('%d.%m.%Y')}", fill="black", font=normal_font, anchor="mm")
    y_pos += 40
    
    # Дополнительная информация
    draw.text((width // 2, y_pos), "FREEDOM Волонтерский центр", fill="black", font=normal_font, anchor="mm")
    
    # Сохраняем в буфер
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer.getvalue()

@team_leader_required
def generate_certificate(request, participant_id):
    """
    Генерирует благодарственное письмо
    """
    try:
        format_type = request.GET.get('format', 'pdf')  # По умолчанию pdf, может быть png
        
        if format_type not in ['png', 'pdf']:
            format_type = 'pdf'
        
        participant = EventParticipant.objects.get(id=participant_id)
        event = participant.event
        scanner = participant.volunteer
        
        # Полное имя сканера для замены (в верхнем регистре)
        full_name = f"{scanner.first_name} {scanner.last_name}".upper()
        
        # Форматируем часы как округленное целое число
        hours = round(participant.hours_awarded)
        
        # Данные для сертификата
        event_name = event.name
        event_date = event.date.strftime("%d.%m.%Y")
        leader_name = f"{event.leader.first_name} {event.leader.last_name}" if event.leader else None
        
        # Генерируем сертификат
        if format_type == 'pdf':
            file_data = generate_certificate_pdf(
                full_name, hours, event_name, event_date, leader_name
            )
            content_type = 'application/pdf'
        else:  # png
            file_data = generate_certificate_png(
                full_name, hours, event_name, event_date, leader_name
            )
            content_type = 'image/png'
        
        # Обнуляем часы сканера при получении благодарственного письма
        participant.hours_awarded = 0
        participant.save()
        
        # Отдаем файл для скачивания
        extension = format_type
        filename = f"certificate_{scanner.last_name}_{event.name}.{extension}"
        response = HttpResponse(file_data, content_type=content_type)
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
        format_type = request.GET.get('format', 'pdf')  # По умолчанию pdf, может быть png
        
        if format_type not in ['png', 'pdf']:
            format_type = 'pdf'
        
        event = Event.objects.get(id=event_id)
        participants = EventParticipant.objects.filter(event=event)
        
        # Создаем zip-архив с благодарностями
        temp_zip = io.BytesIO()
        
        with zipfile.ZipFile(temp_zip, 'w') as zipf:
            for participant in participants:
                scanner = participant.volunteer
                
                # Полное имя сканера (в верхнем регистре)
                full_name = f"{scanner.first_name} {scanner.last_name}".upper()
                
                # Форматируем часы как округленное целое число
                hours = round(participant.hours_awarded)
                
                # Данные для сертификата
                event_name = event.name
                event_date = event.date.strftime("%d.%m.%Y")
                leader_name = f"{event.leader.first_name} {event.leader.last_name}" if event.leader else None
                
                # Генерируем сертификат
                if format_type == 'pdf':
                    file_data = generate_certificate_pdf(
                        full_name, hours, event_name, event_date, leader_name
                    )
                else:  # png
                    file_data = generate_certificate_png(
                        full_name, hours, event_name, event_date, leader_name
                    )
                
                # Обнуляем часы сканера при получении благодарственного письма
                participant.hours_awarded = 0
                participant.save()
                
                # Подготавливаем имя файла
                filename = f"certificate_{scanner.last_name}_{scanner.first_name}.{format_type}"
                zipf.writestr(filename, file_data)
        
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
def generate_scanner_certificate(request, scanner_id):
    """
    Генерирует благодарственное письмо сканеру за все мероприятия, в которых он участвовал
    """
    try:
        format_type = request.GET.get('format', 'pdf')  # По умолчанию pdf, может быть png
        
        if format_type not in ['png', 'pdf']:
            format_type = 'pdf'
        
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
        
        # Период участия
        period_text = f"{first_date.strftime('%d.%m.%Y')} - {last_date.strftime('%d.%m.%Y')}"
        
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
        
        # Форматируем суммарные часы как округленное целое число
        hours = round(total_hours)
        
        # Генерируем сертификат
        if format_type == 'pdf':
            file_data = generate_certificate_pdf(
                full_name, hours, period=period_text, events_list=events_list
            )
            content_type = 'application/pdf'
        else:  # png
            file_data = generate_certificate_png(
                full_name, hours, period=period_text, events_list=events_list
            )
            content_type = 'image/png'
        
        # Обнуляем часы сканера при получении благодарственного письма
        for participant in participations:
            participant.hours_awarded = 0
            participant.save()
        
        # Отдаем файл для скачивания
        extension = format_type
        filename = f"certificate_{scanner.last_name}_{scanner.first_name}.{extension}"
        response = HttpResponse(file_data, content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        return JsonResponse({"error": str(e), "trace": trace}, status=400) 