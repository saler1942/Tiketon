from django.contrib import admin
from .models import Volunteer, Event, EventParticipant
from django.contrib.auth.models import User, Group
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.utils.translation import gettext_lazy as _

# Отменяем стандартную регистрацию User
admin.site.unregister(User)

# Создаем кастомный класс для админки User
class UserAdmin(BaseUserAdmin):
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_team_leader', 'is_staff')
    list_filter = ('is_staff', 'is_superuser', 'groups')
    search_fields = ('username', 'email', 'first_name', 'last_name')
    ordering = ('email',)
    fieldsets = (
        (None, {'fields': ('username', 'password')}),
        (_('Personal info'), {'fields': ('first_name', 'last_name', 'email')}),
        (_('Permissions'), {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups')}),
        (_('Important dates'), {'fields': ('last_login', 'date_joined')}),
    )
    actions = ['make_team_leader']

    def is_team_leader(self, obj):
        return obj.groups.filter(name='Тимлидеры').exists()
    is_team_leader.boolean = True
    is_team_leader.short_description = 'Тимлидер'

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        group, created = Group.objects.get_or_create(name='Тимлидеры')
        obj.groups.add(group)

    def make_team_leader(self, request, queryset):
        group, created = Group.objects.get_or_create(name='Тимлидеры')
        for user in queryset:
            user.groups.add(group)
        self.message_user(request, 'Выбранные пользователи добавлены в тимлидеры.')
    make_team_leader.short_description = 'Сделать тимлидером'

# Регистрируем кастомную админку для User
admin.site.register(User, UserAdmin)

# Регистрируем остальные модели
admin.site.register(Volunteer)
admin.site.register(Event)
admin.site.register(EventParticipant)
