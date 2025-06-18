from django.urls import path
from .views import login_request, code_verify, events_list, event_create, event_edit, volunteer_search, home, export_event_participants, import_excel, export_all_events
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
    path('api/volunteer-search/', volunteer_search, name='volunteer_search'),
    path('import/', import_excel, name='import_excel'),
] 