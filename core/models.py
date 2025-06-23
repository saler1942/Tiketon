from django.db import models
from django.contrib.auth.models import User

# Create your models here.

class Scanner(models.Model):
    first_name = models.CharField(max_length=64)
    last_name = models.CharField(max_length=64)
    email = models.EmailField(unique=True)

    def __str__(self):
        return f'{self.first_name} {self.last_name}'

class Event(models.Model):
    name = models.CharField(max_length=128)
    date = models.DateField()
    volunteers_required = models.PositiveIntegerField()
    leader = models.ForeignKey(User, on_delete=models.CASCADE)
    duration_hours = models.FloatField(null=True, blank=True, help_text="Длительность мероприятия в часах")

    def __str__(self):
        return self.name

class EventParticipant(models.Model):
    event = models.ForeignKey(Event, on_delete=models.CASCADE, related_name='participants')
    volunteer = models.ForeignKey(Scanner, on_delete=models.CASCADE)
    is_late = models.BooleanField(default=False)
    late_minutes = models.PositiveIntegerField(null=True, blank=True)
    hours_awarded = models.FloatField(default=0)

    class Meta:
        unique_together = ('event', 'volunteer')

    def __str__(self):
        return f'{self.volunteer} @ {self.event}'

# Добавляем метод проверки тимлидера для User
def is_team_leader(user):
    return user.groups.filter(name='Тимлидеры').exists()

User.add_to_class('is_team_leader', property(is_team_leader))
