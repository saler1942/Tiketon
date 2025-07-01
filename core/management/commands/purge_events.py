from django.core.management.base import BaseCommand
from django.utils import timezone
from django.core.mail import send_mail
from django.conf import settings
from django.db.models import Count
from core.models import Event, PurgeSettings, NotificationLog, TeamLeader
from django.contrib.auth.models import User, Group
from datetime import timedelta
from core.utils import send_telegram_message
import logging

logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'Purges events older than one year and sends notification via Telegram one week before deletion'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Show what would be deleted without actually deleting',
        )
        parser.add_argument(
            '--notify-only',
            action='store_true',
            help='Only send notification messages without deleting anything',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        notify_only = options['notify_only']
        
        # Получаем настройки из базы данных или используем значения по умолчанию
        purge_settings = PurgeSettings.objects.first()
        
        if not purge_settings:
            self.stdout.write(self.style.WARNING("No purge settings found. Using default settings."))
            notification_days_before = 7
            is_active = True
        else:
            notification_days_before = purge_settings.notification_days_before
            is_active = purge_settings.active
        
        # Если функция отключена в настройках и не запущена принудительно, выходим
        if not is_active and not (dry_run or notify_only):
            self.stdout.write(self.style.WARNING("Automatic purge is disabled in settings. Use --dry-run to preview or update settings to enable."))
            return
        
        # Current date
        now = timezone.now().date()
        
        # Date one year ago
        one_year_ago = now - timedelta(days=365)
        
        # Date for notification (defined by settings)
        notification_date = one_year_ago + timedelta(days=notification_days_before)
        
        # Get events to be deleted (older than one year)
        events_to_delete = Event.objects.filter(date__lt=one_year_ago)
        delete_count = events_to_delete.count()
        
        # Get events to be notified about (will be deleted in the notification period)
        events_to_notify = Event.objects.filter(
            date__gte=one_year_ago,
            date__lte=notification_date
        )
        notify_count = events_to_notify.count()
        
        # Notification logic
        if notify_count > 0:
            # Group events by creator for targeted notification
            team_leaders = {}
            
            for event in events_to_notify:
                if event.created_by_id not in team_leaders:
                    team_leaders[event.created_by_id] = {
                        'email': event.created_by.email,
                        'name': f"{event.created_by.first_name} {event.created_by.last_name}",
                        'events': []
                    }
                    
                    # Пытаемся найти Telegram ID для тимлидера
                    try:
                        team_leader = TeamLeader.objects.filter(
                            first_name=event.created_by.first_name,
                            last_name=event.created_by.last_name
                        ).first()
                        if team_leader and team_leader.telegram_id:
                            team_leaders[event.created_by_id]['telegram_id'] = team_leader.telegram_id
                    except Exception as e:
                        self.stdout.write(self.style.WARNING(f"Failed to find Telegram ID for {team_leaders[event.created_by_id]['name']}: {str(e)}"))
                
                team_leaders[event.created_by_id]['events'].append({
                    'id': event.id,
                    'name': event.name,
                    'date': event.date.strftime('%d.%m.%Y')
                })
            
            # Send notification to each team leader
            for leader_id, leader_data in team_leaders.items():
                events_list = "\n".join([f"- {e['name']} ({e['date']})" for e in leader_data['events']])
                
                message = f"""
                Здравствуйте, {leader_data['name']}!

                Уведомляем вас, что следующие созданные вами мероприятия будут автоматически удалены через {notification_days_before} дней:

                {events_list}

                Это происходит в соответствии с политикой хранения данных - мероприятия хранятся один год.
                Если вам нужно сохранить данные об этих мероприятиях, пожалуйста, экспортируйте их в течение следующих {notification_days_before} дней.

                С уважением,
                Команда Ticketon
                """
                
                if not dry_run:
                    # Отправляем через Telegram, если есть ID
                    if 'telegram_id' in leader_data and leader_data['telegram_id']:
                        try:
                            success = send_telegram_message(message, [leader_data['telegram_id']])
                            
                            if success:
                                # Логируем отправку
                                NotificationLog.objects.create(
                                    recipient_telegram_id=leader_data['telegram_id'],
                                    message=message,
                                    is_test=False,
                                    notification_type='telegram'
                                )
                                
                                self.stdout.write(self.style.SUCCESS(f"Telegram notification sent to {leader_data['name']} ({leader_data['telegram_id']})"))
                            else:
                                self.stdout.write(self.style.ERROR(f"Failed to send Telegram notification to {leader_data['name']}"))
                                
                                # Если не удалось отправить через Telegram, пробуем через email
                                if leader_data['email']:
                                    try:
                                        send_mail(
                                            subject='Важно: предстоящее удаление мероприятий',
                                            message=message,
                                            from_email=settings.DEFAULT_FROM_EMAIL,
                                            recipient_list=[leader_data['email']],
                                            fail_silently=False,
                                        )
                                        
                                        # Логируем отправку
                                        NotificationLog.objects.create(
                                            recipient_email=leader_data['email'],
                                            message=message,
                                            is_test=False,
                                            notification_type='email'
                                        )
                                        
                                        self.stdout.write(self.style.SUCCESS(f"Email notification sent to {leader_data['email']} as fallback"))
                                    except Exception as e:
                                        self.stdout.write(self.style.ERROR(f"Failed to send email notification to {leader_data['email']}: {str(e)}"))
                        except Exception as e:
                            self.stdout.write(self.style.ERROR(f"Failed to send Telegram notification to {leader_data['name']}: {str(e)}"))
                    
                    # Если нет Telegram ID, отправляем через email
                    elif leader_data['email']:
                        try:
                            send_mail(
                                subject='Важно: предстоящее удаление мероприятий',
                                message=message,
                                from_email=settings.DEFAULT_FROM_EMAIL,
                                recipient_list=[leader_data['email']],
                                fail_silently=False,
                            )
                            
                            # Логируем отправку
                            NotificationLog.objects.create(
                                recipient_email=leader_data['email'],
                                message=message,
                                is_test=False,
                                notification_type='email'
                            )
                            
                            self.stdout.write(self.style.SUCCESS(f"Email notification sent to {leader_data['email']}"))
                        except Exception as e:
                            self.stdout.write(self.style.ERROR(f"Failed to send email notification to {leader_data['email']}: {str(e)}"))
                else:
                    notification_method = "Telegram" if 'telegram_id' in leader_data else "Email"
                    recipient = leader_data.get('telegram_id', leader_data['email'])
                    self.stdout.write(f"Would send {notification_method} notification to {recipient} about {len(leader_data['events'])} events")
        
        # Deletion logic (skip if notify-only mode)
        if not notify_only and delete_count > 0:
            if dry_run:
                self.stdout.write(f"Would delete {delete_count} events older than {one_year_ago}")
                for event in events_to_delete:
                    self.stdout.write(f"  - {event.name} ({event.date})")
            else:
                # Also send a notification about the actual deletion
                admin_chat_ids = settings.TELEGRAM_CHAT_IDS
                if admin_chat_ids and admin_chat_ids != ['']:
                    try:
                        message = f'Удалено {delete_count} мероприятий, созданных до {one_year_ago}'
                        
                        success = send_telegram_message(message, admin_chat_ids)
                        
                        if success:
                            # Логируем отправку для каждого получателя
                            for chat_id in admin_chat_ids:
                                if chat_id:
                                    NotificationLog.objects.create(
                                        recipient_telegram_id=chat_id,
                                        message=message,
                                        is_test=False,
                                        notification_type='telegram'
                                    )
                            
                            self.stdout.write(self.style.SUCCESS(f"Admin notification sent to Telegram"))
                        else:
                            self.stdout.write(self.style.ERROR(f"Failed to send admin notification to Telegram"))
                            
                            # Если не удалось отправить через Telegram, пробуем через email
                            admin_emails = [admin[1] for admin in settings.ADMINS] if hasattr(settings, 'ADMINS') else []
                            if admin_emails:
                                try:
                                    send_mail(
                                        subject='Уведомление: удаление устаревших мероприятий',
                                        message=message,
                                        from_email=settings.DEFAULT_FROM_EMAIL,
                                        recipient_list=admin_emails,
                                        fail_silently=False,
                                    )
                                    
                                    # Логируем отправку
                                    for email in admin_emails:
                                        NotificationLog.objects.create(
                                            recipient_email=email,
                                            message=message,
                                            is_test=False,
                                            notification_type='email'
                                        )
                                    
                                    self.stdout.write(self.style.SUCCESS(f"Admin notification sent to email as fallback"))
                                except Exception as e:
                                    self.stdout.write(self.style.ERROR(f"Failed to send admin email notification: {str(e)}"))
                    except Exception as e:
                        self.stdout.write(self.style.ERROR(f"Failed to send admin notification: {str(e)}"))
                
                # Perform the actual deletion
                events_to_delete.delete()
                self.stdout.write(self.style.SUCCESS(f"Successfully deleted {delete_count} events older than {one_year_ago}"))
        
        # Summary
        if notify_count == 0 and delete_count == 0:
            self.stdout.write("No events to notify about or delete") 