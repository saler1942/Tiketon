from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import random
import string
import re
from django.core.exceptions import ValidationError

def validate_english_only(value):
    """Validates that the input contains only English characters, digits, spaces and hyphens."""
    if not re.match(r'^[a-zA-Z0-9\s\-]+$', value):
        raise ValidationError('Only English characters, digits, spaces and hyphens are allowed.')

# Create your models here.

class Scanner(models.Model):
    first_name = models.CharField(max_length=100, validators=[validate_english_only])
    last_name = models.CharField(max_length=100, validators=[validate_english_only])
    email = models.EmailField(blank=True, null=True)
    telegram_id = models.CharField(max_length=50, blank=True, null=True, help_text="ID пользователя в Telegram")
    created_at = models.DateTimeField(default=timezone.now)
    total_certificate_hours = models.FloatField(default=0.0, help_text="Общее количество часов, полученных в сертификатах")
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"
    
    def clean(self):
        # Additional validation to ensure names are in English
        if any(ord(char) > 127 for char in self.first_name if char not in [' ', '-']):
            raise ValidationError({'first_name': 'First name must contain only English characters.'})
        if any(ord(char) > 127 for char in self.last_name if char not in [' ', '-']):
            raise ValidationError({'last_name': 'Last name must contain only English characters.'})

class TeamLeader(models.Model):
    first_name = models.CharField(max_length=100)
    last_name = models.CharField(max_length=100)
    email = models.EmailField(blank=True, null=True)
    telegram_id = models.CharField(max_length=50, blank=True, null=True, help_text="ID пользователя в Telegram")
    scanner = models.OneToOneField(Scanner, on_delete=models.CASCADE, related_name='teamleader', null=True, blank=True)
    created_at = models.DateTimeField(default=timezone.now)
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"
    
    def save(self, *args, **kwargs):
        # При сохранении тимлидера, проверяем наличие связанного сканера
        if not self.scanner:
            # Ищем сканера с таким же именем и фамилией
            scanner = Scanner.objects.filter(
                first_name=self.first_name,
                last_name=self.last_name
            ).first()
            
            # Если сканер не найден, создаем его
            if not scanner:
                scanner = Scanner.objects.create(
                    first_name=self.first_name,
                    last_name=self.last_name,
                    email=self.email,
                    telegram_id=self.telegram_id
                )
            
            self.scanner = scanner
        
        super().save(*args, **kwargs)

class TeamLeaderProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    
    def __str__(self):
        return self.user.username

def generate_random_code():
    """Генерирует случайный код для события"""
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))

class Event(models.Model):
    name = models.CharField(max_length=200)
    date = models.DateField()
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    location = models.CharField(max_length=200, default='Астана')
    description = models.TextField(blank=True, default='')
    created_by = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(default=timezone.now)
    code = models.CharField(max_length=10, default='DEFAULT000')
    duration_hours = models.FloatField(default=1.0)
    max_scanners = models.IntegerField(default=10)
    
    def __str__(self):
        return self.name

class EventParticipant(models.Model):
    event = models.ForeignKey(Event, on_delete=models.CASCADE)
    volunteer = models.ForeignKey(Scanner, on_delete=models.CASCADE)
    registered_at = models.DateTimeField(default=timezone.now)
    hours_awarded = models.FloatField(default=0.0)
    
    class Meta:
        unique_together = ('event', 'volunteer')
    
    def __str__(self):
        return f"{self.volunteer} - {self.event}"

# Добавляем метод проверки тимлидера для User
def is_team_leader(user):
    return user.groups.filter(name='Тимлидеры').exists()

User.add_to_class('is_team_leader', property(is_team_leader))

# Новая модель для настроек автоматического удаления
class PurgeSettings(models.Model):
    purge_date = models.DateField(
        verbose_name="Дата ежегодного обновления", 
        help_text="Дата, в которую будет происходить ежегодное удаление мероприятий"
    )
    notification_days_before = models.IntegerField(
        default=7,
        verbose_name="Дней до уведомления", 
        help_text="За сколько дней до удаления отправлять уведомления"
    )
    active = models.BooleanField(
        default=True, 
        verbose_name="Активно",
        help_text="Включить/выключить автоматическое удаление"
    )
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Последнее обновление")
    updated_by = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True,
        verbose_name="Кто обновил"
    )
    
    class Meta:
        verbose_name = "Настройка обновления базы"
        verbose_name_plural = "Настройки обновления базы"
    
    def __str__(self):
        return f"Удаление мероприятий {self.purge_date.strftime('%d.%m')}"
    
    def save(self, *args, **kwargs):
        # Если модель создается впервые, установим 13 июля текущего года
        if not self.pk and not self.purge_date:
            current_year = timezone.now().year
            self.purge_date = timezone.datetime(current_year, 7, 13).date()
        
        # Сохраняем только одну запись
        if not self.pk:
            PurgeSettings.objects.all().delete()
            
        super().save(*args, **kwargs)

# Модель для хранения истории уведомлений
class NotificationLog(models.Model):
    NOTIFICATION_TYPE_CHOICES = (
        ('email', 'Email'),
        ('telegram', 'Telegram'),
    )
    
    sent_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата отправки")
    sent_by = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True,
        verbose_name="Отправитель"
    )
    recipient_email = models.EmailField(verbose_name="Email получателя", blank=True, null=True)
    recipient_telegram_id = models.CharField(max_length=50, blank=True, null=True, verbose_name="Telegram ID получателя")
    notification_type = models.CharField(
        max_length=10, 
        choices=NOTIFICATION_TYPE_CHOICES, 
        default='email',
        verbose_name="Тип уведомления"
    )
    message = models.TextField(verbose_name="Сообщение")
    is_test = models.BooleanField(default=False, verbose_name="Тестовое уведомление")
    
    class Meta:
        verbose_name = "Лог уведомлений"
        verbose_name_plural = "Логи уведомлений"
        ordering = ['-sent_at']
    
    def __str__(self):
        recipient = self.recipient_email or self.recipient_telegram_id or "Неизвестный получатель"
        return f"Уведомление для {recipient} от {self.sent_at.strftime('%d.%m.%Y %H:%M')}"
