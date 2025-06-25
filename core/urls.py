from django.urls import path
from .views import login_request, code_verify, events_list, event_create, event_edit, volunteer_search as scanner_search, home, export_event_participants, export_all_events, generate_certificate, generate_all_certificates, generate_scanner_certificate, scanner_certificates, scanner_events_api, debug_template, all_scanners_list, event_delete
from django.http import HttpResponse

urlpatterns = [
    path('', home, name='home'),
    path('login/', login_request, name='login'),
    path('code-verify/', code_verify, name='code_verify'),
    # stub для событий
    path('events/', events_list, name='events'),
    path('events/create/', event_create, name='event_create'),
    path('events/<int:event_id>/edit/', event_edit, name='event_edit'),
    path('events/<int:event_id>/export/', export_event_participants, name='event_export'),
    path('events/export/all/', export_all_events, name='export_all_events'),
    path('api/scanner-search/', scanner_search, name='volunteer_search'),
    # URL для благодарственных писем
    path('certificates/participant/<int:participant_id>/', generate_certificate, name='generate_certificate'),
    path('certificates/event/<int:event_id>/', generate_all_certificates, name='generate_all_certificates'),
    path('certificates/scanner/<int:scanner_id>/', generate_scanner_certificate, name='generate_scanner_certificate'),
    # Новая страница для поиска сканеров и генерации сертификатов
    path('certificates/', scanner_certificates, name='scanner_certificates'),
    path('api/scanner-events/<int:scanner_id>/', scanner_events_api, name='scanner_events_api'),
    # Список всех сканеров с фильтрацией
    path('scanners/', all_scanners_list, name='all_scanners'),
    # Диагностические URL
    path('debug/template/', debug_template, name='debug_template'),
    path('events/<int:event_id>/delete/', event_delete, name='event_delete'),
] 