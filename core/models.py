from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import random
import string

# Create your models here.

class Scanner(models.Model):
    first_name = models.CharField(max_length=100)
    last_name = models.CharField(max_length=100)
    email = models.EmailField(blank=True, null=True)
    created_at = models.DateTimeField(default=timezone.now)
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"

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
