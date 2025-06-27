from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('login/', views.login_request, name='login'),
    path('events/', views.events_list, name='events'),
    path('events/create/', views.event_create, name='event_create'),
    path('events/<int:event_id>/edit/', views.event_edit, name='event_edit'),
    path('events/<int:event_id>/delete/', views.event_delete, name='event_delete'),
    path('events/<int:event_id>/export/', views.export_event_participants, name='export_event_participants'),
    path('events/export-all/', views.export_all_events, name='export_all_events'),
    path('verify/', views.code_verify, name='code_verify'),
    path('api/scanner-search/', views.volunteer_search, name='volunteer_search'),
    
    # Сертификаты
    path('certificates/<int:participant_id>/', views.generate_certificate, name='generate_certificate'),
    path('certificates/event/<int:event_id>/', views.generate_all_certificates, name='generate_all_certificates'),
    path('certificates/scanner/<int:scanner_id>/', views.generate_scanner_certificate, name='generate_scanner_certificate'),
    path('certificates/all-scanners/', views.generate_all_scanner_certificates, name='generate_all_scanner_certificates'),
    path('certificates/', views.scanner_certificates, name='scanner_certificates'),
    
    # API для получения данных о мероприятиях сканера
    path('api/scanner-events/<int:scanner_id>/', views.scanner_events_api, name='scanner_events_api'),
    
    # Список сканеров
    path('scanners/', views.all_scanners_list, name='all_scanners_list'),
    
    # Диагностика шаблона
    path('debug/template/', views.debug_template, name='debug_template'),
] 