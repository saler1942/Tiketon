from django.db import migrations

class Migration(migrations.Migration):

    dependencies = [
        ('core', '0002_remove_event_duration_minutes_event_duration_hours'),
    ]

    operations = [
        migrations.RenameModel(
            old_name='Volunteer',
            new_name='Scanner',
        ),
    ] 