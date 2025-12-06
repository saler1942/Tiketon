from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import random
import string

# Create your models here.

class Scanner(models.Model):
    first_name = models.CharField(max_length=100, verbose_name='Имя')
    last_name = models.CharField(max_length=100, verbose_name='Фамилия')
    email = models.EmailField(unique=True, verbose_name='Email')
    total_certificate_hours = models.FloatField(default=0.0, verbose_name='Часы сертификата')
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Дата создания')
    
    class Meta:
        verbose_name = 'Сканер'
        verbose_name_plural = 'Сканеры'
        ordering = ['last_name', 'first_name']
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"

class TeamLeader(models.Model):
    first_name = models.CharField(max_length=100, verbose_name='Имя')
    last_name = models.CharField(max_length=100, verbose_name='Фамилия')
    email = models.EmailField(verbose_name='Email')
    scanner = models.OneToOneField(Scanner, on_delete=models.SET_NULL, null=True, blank=True, verbose_name='Сканер')
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Дата создания')
    
    class Meta:
        verbose_name = 'Тимлидер'
        verbose_name_plural = 'Тимлидеры'
        ordering = ['last_name', 'first_name']
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"
    
    def save(self, *args, **kwargs):
        # При сохранении тимлидера, проверяем наличие связанного сканера
        if not self.scanner:
            # Ищем сканер с таким же именем и фамилией
            scanner = Scanner.objects.filter(
                first_name=self.first_name,
                last_name=self.last_name
            ).first()
            
            # Если сканер не найден, создаем его
            if not scanner:
                scanner = Scanner.objects.create(
                    first_name=self.first_name,
                    last_name=self.last_name,
                    email=self.email
                )
            
            self.scanner = scanner
        
        super().save(*args, **kwargs)

class TeamLeaderProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    
    class Meta:
        verbose_name = 'Профиль тимлидера'
        verbose_name_plural = 'Профили тимлидеров'
    
    def __str__(self):
        return self.user.username

def generate_random_code():
    """Генерирует случайный код для события"""
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))

class Event(models.Model):
    name = models.CharField(max_length=200, verbose_name='Название')
    date = models.DateField(verbose_name='Дата')
    location = models.CharField(max_length=200, blank=True, verbose_name='Место проведения')
    description = models.TextField(blank=True, verbose_name='Описание')
    max_scanners = models.IntegerField(default=0, verbose_name='Максимум сканеров')
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, verbose_name='Создал')
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Дата создания')
    start_date = models.DateTimeField(null=True, blank=True, verbose_name='Дата начала')
    end_date = models.DateTimeField(null=True, blank=True, verbose_name='Дата окончания')
    code = models.CharField(max_length=10, default='DEFAULT000', verbose_name='Код')
    duration_hours = models.FloatField(default=1.0, verbose_name='Длительность')
    
    class Meta:
        verbose_name = 'Мероприятие'
        verbose_name_plural = 'Мероприятия'
        ordering = ['-date']
    
    def __str__(self):
        return self.name

class EventParticipant(models.Model):
    event = models.ForeignKey(Event, on_delete=models.CASCADE, verbose_name='Мероприятие')
    volunteer = models.ForeignKey(Scanner, on_delete=models.CASCADE, verbose_name='Сканер')
    registered_at = models.DateTimeField(default=timezone.now, verbose_name='Дата регистрации')
    hours_awarded = models.FloatField(default=0.0, verbose_name='Начисленные часы')
    hours_awarded_backup = models.FloatField(default=0.0, verbose_name='Резервные часы')
    
    class Meta:
        verbose_name = 'Ивент'
        verbose_name_plural = 'Ивенты'
        unique_together = ('event', 'volunteer')
    
    def __str__(self):
        return f"{self.volunteer} - {self.event}"

# Добавляем метод проверки ответственного для User
def is_team_leader(user):
    return user.groups.filter(name='Ответственные').exists()

User.add_to_class('is_team_leader', property(is_team_leader))
