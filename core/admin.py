from django.contrib import admin
from django import forms
from django.db import models
from dal import autocomplete
from .models import Scanner, Event, EventParticipant, TeamLeader
from django.contrib.auth.models import User, Group, Permission
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin, GroupAdmin
from django.utils.translation import gettext_lazy as _

# Proxy-модель для User с кастомным verbose_name
class TeamLeaderUser(User):
    class Meta:
        proxy = True
        verbose_name = 'Ответственный'
        verbose_name_plural = 'Ответственные'

# Отменяем стандартную регистрацию User и Group
admin.site.unregister(User)
admin.site.unregister(Group)

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
        return obj.groups.filter(name='Ответственные').exists()
    is_team_leader.boolean = True
    is_team_leader.short_description = 'Тимлидер'

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change)
        group, created = Group.objects.get_or_create(name='Ответственные')
        obj.groups.add(group)

    def make_team_leader(self, request, queryset):
        group, created = Group.objects.get_or_create(name='Ответственные')
        for user in queryset:
            user.groups.add(group)
        self.message_user(request, 'Выбранные пользователи добавлены в Ответственные.')
    make_team_leader.short_description = 'Сделать Ответственным'

# Кастомная форма для TeamLeaderAdmin
class TeamLeaderAdminForm(forms.ModelForm):
    scanner_admin = forms.ModelChoiceField(
        queryset=Scanner.objects.all().order_by('last_name', 'first_name'),
        required=False,
        label='Сканер',
        widget=autocomplete.ModelSelect2(
            url='/admin/scanner-autocomplete/',
            attrs={
                'data-placeholder': 'Поиск по имени или фамилии...',
                'data-minimum-input-length': 1,
                'onchange': 'updateUserInfo(this);'
            }
        )
    )
    
    class Meta:
        model = User
        fields = '__all__'
    
    class Media:
        js = ('admin/js/scanner_autofill.js',)
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Если редактируем существующего пользователя, загружаем его сканера
        if self.instance and self.instance.pk:
            try:
                teamleader = TeamLeader.objects.get(
                    first_name=self.instance.first_name,
                    last_name=self.instance.last_name
                )
                if teamleader.scanner:
                    self.fields['scanner_admin'].initial = teamleader.scanner
            except TeamLeader.DoesNotExist:
                pass
        
        # Переименовываем поля разрешений с новыми описаниями
        if 'is_active' in self.fields:
            self.fields['is_active'].label = 'Junior'
            self.fields['is_active'].help_text = 'Данный пользователь имеет только доступ к сайту, и не больше'
        
        if 'is_staff' in self.fields:
            self.fields['is_staff'].label = 'Middle'
            self.fields['is_staff'].help_text = 'Данный пользователь имеет доступ к админ панели сайта для редактирования данных или же добавления новых ответственных'
        
        if 'is_superuser' in self.fields:
            self.fields['is_superuser'].label = 'Senior'
            self.fields['is_superuser'].help_text = 'Данный пользователь имеет полный доступ ко всей информационной части сайта и админ панели, senior может быть добавлен только пользователем с такой же ролью'

# Объединенный класс для админки TeamLeader и TeamLeaderUser
class TeamLeaderAdmin(BaseUserAdmin):
    form = TeamLeaderAdminForm
    list_display = ('username', 'first_name', 'last_name', 'email', 'get_scanner', 'get_status')
    list_filter = ('is_staff', 'is_superuser', 'date_joined')
    search_fields = ('username', 'email', 'first_name', 'last_name')
    ordering = ('last_name', 'first_name')
    
    def get_fieldsets(self, request, obj=None):
        fieldsets = (
            (None, {'fields': ('username', 'password')}),
            (_('Персональная информация'), {'fields': ('first_name', 'last_name', 'email')}),
            (_('Назначение сканера'), {'fields': ('scanner_admin',)}),
            (_('Разрешения'), {'fields': ('is_active', 'is_staff')}),
            (_('Важные даты'), {'fields': ('last_login', 'date_joined')}),
        )
        
        # Добавляем поле is_superuser только для суперпользователей
        if request.user.is_superuser:
            fieldsets = (
                (None, {'fields': ('username', 'password')}),
                (_('Персональная информация'), {'fields': ('first_name', 'last_name', 'email')}),
                (_('Назначение сканера'), {'fields': ('scanner_admin',)}),
                (_('Разрешения'), {'fields': ('is_active', 'is_staff', 'is_superuser')}),
                (_('Важные даты'), {'fields': ('last_login', 'date_joined')}),
            )
        
        return fieldsets
    
    def get_scanner(self, obj):
        """Получить сканера для тимлидера"""
        try:
            teamleader = TeamLeader.objects.get(
                first_name=obj.first_name,
                last_name=obj.last_name
            )
            return teamleader.scanner
        except TeamLeader.DoesNotExist:
            return None
    get_scanner.short_description = 'Сканер'
    
    def get_status(self, obj):
        """Получить статус пользователя"""
        if obj.is_superuser:
            return 'Senior'
        elif obj.is_staff:
            return 'Middle'
        else:
            return 'Junior'
    get_status.short_description = 'Статус'
    
    def save_model(self, request, obj, form, change):
        """Сохраняем пользователя и связываем со сканером"""
        super().save_model(request, obj, form, change)
        
        # Если пользователь отмечен как тимлидер, добавляем его в группу "Ответственные"
        if obj.is_staff or obj.is_superuser:
            group, created = Group.objects.get_or_create(name='Ответственные')
            obj.groups.add(group)
            
            # Назначаем базовые права для доступа к админке для Middle пользователей
            if obj.is_staff and not obj.is_superuser:
                try:
                    permissions = [
                        Permission.objects.get(codename='view_scanner'),
                        Permission.objects.get(codename='change_scanner'),
                        Permission.objects.get(codename='view_event'),
                        Permission.objects.get(codename='change_event'),
                        Permission.objects.get(codename='view_eventparticipant'),
                        Permission.objects.get(codename='change_eventparticipant'),
                    ]
                    for perm in permissions:
                        obj.user_permissions.add(perm)
                except Permission.DoesNotExist:
                    # Если какие-то права не найдены, пропускаем их
                    pass
            
            # Если есть сканер в форме, связываем его с тимлидером
            scanner_admin = form.cleaned_data.get('scanner_admin')
            if scanner_admin:
                # Создаем или обновляем тимлидера
                teamleader, created = TeamLeader.objects.update_or_create(
                    first_name=obj.first_name,
                    last_name=obj.last_name,
                    defaults={
                        'email': obj.email,
                        'scanner': scanner_admin
                    }
                )
        else:
            # Если пользователь не тимлидер, убираем из группы
            group = Group.objects.filter(name='Ответственные').first()
            if group:
                obj.groups.remove(group)
    
    def make_team_leader(self, request, queryset):
        """Экшн для назначения ответственного"""
        group, created = Group.objects.get_or_create(name='Ответственные')
        for user in queryset:
            user.groups.add(group)
        self.message_user(request, 'Выбранные пользователи добавлены в Ответственные.')
    make_team_leader.short_description = 'Сделать Ответственным'

# Регистрируем объединенную админку для TeamLeaderUser
admin.site.register(TeamLeaderUser, TeamLeaderAdmin)

# Inline класс для EventParticipant в ScannerAdmin
class EventParticipantInline(admin.TabularInline):
    model = EventParticipant
    extra = 0
    fields = ('event', 'hours_awarded', 'hours_awarded_backup')
    readonly_fields = ('event',)
    can_delete = False

# Класс для админки сканера с фильтрацией и поиском
class ScannerAdmin(admin.ModelAdmin):
    list_display = ('first_name', 'last_name', 'email', 'total_certificate_hours', 'get_events_count')
    list_filter = ('first_name', 'last_name')
    search_fields = ('first_name', 'last_name', 'email')
    ordering = ('last_name', 'first_name')
    fields = ('first_name', 'last_name', 'email', 'total_certificate_hours')
    inlines = [EventParticipantInline]
    
    def get_events_count(self, obj):
        """Получить количество мероприятий, в которых участвовал сканер"""
        return EventParticipant.objects.filter(volunteer=obj).count()
    get_events_count.short_description = 'Мероприятий'
    
    def get_readonly_fields(self, request, obj=None):
        # Делаем поле total_certificate_hours редактируемым для админов
        if request.user.is_superuser:
            return ()
        return ('total_certificate_hours',)

# Inline класс для EventParticipant в EventAdmin
class EventParticipantEventInline(admin.TabularInline):
    model = EventParticipant
    extra = 0
    fields = ('volunteer', 'hours_awarded', 'hours_awarded_backup')
    can_delete = False

# Класс для админки мероприятий с фильтрацией и поиском
class EventAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'name', 'start_date', 'end_date', 'location', 'get_participants_count'
    ]
    list_filter = ['start_date', 'end_date', 'location']
    search_fields = ('name', 'created_by__username', 'created_by__first_name', 'created_by__last_name')
    ordering = ('-date',)
    inlines = [EventParticipantEventInline]
    
    def get_participants_count(self, obj):
        """Получить количество сканеров на мероприятии"""
        return EventParticipant.objects.filter(event=obj).count()
    get_participants_count.short_description = 'Сканеров'

# Класс для админки участников с фильтрацией и поиском
class EventParticipantAdmin(admin.ModelAdmin):
    list_display = [
        'id', 'get_event_name', 'volunteer', 'hours_awarded', 'get_volunteer_email'
    ]
    list_filter = ['event', 'volunteer']
    search_fields = ('event__name', 'volunteer__first_name', 'volunteer__last_name', 'volunteer__email')
    ordering = ('event__name', 'volunteer__last_name', 'volunteer__first_name')
    fields = ('event', 'volunteer', 'hours_awarded', 'hours_awarded_backup')
    
    def get_event_name(self, obj):
        """Получить название мероприятия"""
        return obj.event.name
    get_event_name.short_description = 'Мероприятие'
    
    def get_volunteer_email(self, obj):
        """Получить email сканера"""
        return obj.volunteer.email
    get_volunteer_email.short_description = 'Email сканера'
    
    def get_readonly_fields(self, request, obj=None):
        # Делаем поля часов редактируемыми для админов
        if request.user.is_superuser:
            return ()
        return ('hours_awarded', 'hours_awarded_backup')

# Autocomplete view для сканеров
class ScannerAutocomplete(autocomplete.Select2QuerySetView):
    def get_queryset(self):
        qs = Scanner.objects.all().order_by('last_name', 'first_name')
        
        if self.q:
            qs = qs.filter(
                models.Q(first_name__icontains=self.q) |
                models.Q(last_name__icontains=self.q) |
                models.Q(email__icontains=self.q)
            )
        
        return qs
    
    def get_result_label(self, result):
        return f"{result.first_name} {result.last_name} ({result.email})"

# Регистрируем остальные модели с кастомной админкой
admin.site.register(Scanner, ScannerAdmin)
admin.site.register(Event, EventAdmin)
# EventParticipant не регистрируем отдельно, только как inline

# Меняем название админки
admin.site.site_header = 'Freedom Ticketon | CHEESES TEAM'
admin.site.site_title = 'Freedom Ticketon'
admin.site.index_title = 'Администрирование системы'
