from django.contrib import admin
from .models import Scanner, Event, EventParticipant, TeamLeader
from django.contrib.auth.models import User, Group
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.utils.translation import gettext_lazy as _

# Proxy-модель для User с кастомным verbose_name
class TeamLeaderUser(User):
    class Meta:
        proxy = True
        verbose_name = 'Тим лидер (пользователь)'
        verbose_name_plural = 'Тим лидеры (пользователи)'

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
        (_('Персональная информация'), {'fields': ('first_name', 'last_name', 'email')}),
        (_('Разрешения'), {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups')}),
        (_('Важные даты'), {'fields': ('last_login', 'date_joined')}),
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

# Регистрируем proxy-модель TeamLeaderUser вместо User
admin.site.register(TeamLeaderUser, UserAdmin)

# Класс для админки TeamLeader
class TeamLeaderAdmin(admin.ModelAdmin):
    list_display = ('first_name', 'last_name', 'email', 'scanner')
    list_filter = ('created_at',)
    search_fields = ('first_name', 'last_name', 'email')
    ordering = ('last_name', 'first_name')
    raw_id_fields = ('scanner',)

# Класс для админки сканера с фильтрацией и поиском
class ScannerAdmin(admin.ModelAdmin):
    list_display = ('first_name', 'last_name', 'email', 'total_certificate_hours')
    list_filter = ('first_name', 'last_name')
    search_fields = ('first_name', 'last_name', 'email')
    ordering = ('last_name', 'first_name')

# Класс для админки мероприятий с фильтрацией и поиском
class EventAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'name', 'start_date', 'end_date', 'location'
    ]
    list_filter = ['start_date', 'end_date', 'location']
    search_fields = ('name', 'created_by__username', 'created_by__first_name', 'created_by__last_name')
    ordering = ('-date',)

# Класс для админки участников с фильтрацией и поиском
class EventParticipantAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'event', 'volunteer'
    ]
    list_filter = ['event']
    search_fields = ('event__name', 'volunteer__first_name', 'volunteer__last_name', 'volunteer__email')
    ordering = ('event', 'volunteer')

# Регистрируем остальные модели с кастомной админкой
admin.site.register(Scanner, ScannerAdmin)
admin.site.register(TeamLeader, TeamLeaderAdmin)
admin.site.register(Event, EventAdmin)
admin.site.register(EventParticipant, EventParticipantAdmin)

# Меняем название админки
admin.site.site_header = 'Freedom Ticketon | TEAM SYRYM'
admin.site.site_title = 'Freedom Ticketon'
admin.site.index_title = 'Администрирование системы'
