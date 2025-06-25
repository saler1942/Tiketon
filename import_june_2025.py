#!/usr/bin/env python
import os
import sys
import django
import openpyxl
from datetime import datetime
import re

# Setup Django environment
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'tiketon.settings')
django.setup()

from django.contrib.auth.models import User, Group
from core.models import Scanner, Event, EventParticipant

def has_phone_number(text):
    """Check if text contains a phone number pattern"""
    if not text:
        return False
    
    # Check for phone number patterns like +7, 8777, etc.
    return bool(re.search(r'(\+[78]|[78]\d{3}|\d{10})', str(text)))

def extract_name(text):
    """Extract name from text, removing phone numbers"""
    if not text:
        return None
    
    # Convert to string if not already
    text = str(text)
    
    # Remove phone number patterns
    clean_text = re.sub(r'(\+?[78]\d{10}|\d{10}|\+[78])', '', text)
    
    return clean_text.strip()

def find_team_leader(name_text, team_leaders):
    """Find team leader by partial name match"""
    if not name_text:
        return None
    
    name_parts = name_text.lower().split()
    
    # Try to find a match among team leaders
    for tl in team_leaders:
        tl_full_name = f"{tl.first_name} {tl.last_name}".lower()
        tl_parts = tl_full_name.split()
        
        # Check if any part of the name matches
        for part in name_parts:
            if len(part) > 2:  # Avoid matching very short parts
                for tl_part in tl_parts:
                    if part in tl_part or tl_part in part:
                        print(f"Matched '{part}' with team leader: {tl.first_name} {tl.last_name}")
                        return tl
    
    print(f"Team leader not found: {name_text}")
    return None

def main():
    print("Starting import of June 2025 events and participants...")
    
    # Path to the Excel file
    excel_path = 'Копия сканеры тикетон астана.xlsx'
    
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(excel_path)
        print(f"Available sheets: {wb.sheetnames}")
        
        # Select the sheet with June 2025 data
        sheet = wb['февраль']  # Use the specific sheet for June 2025
        print(f"Selected sheet: {sheet.title}")
        
        # Get all team leaders from the database
        team_leaders = list(User.objects.filter(groups__name='Тимлидеры'))
        print(f"Found {len(team_leaders)} team leaders in database")
        
        # Print team leaders for debugging
        for tl in team_leaders:
            print(f"DB Team Leader: {tl.first_name} {tl.last_name}")
        
        # Variables to track current event and team leader
        current_event = None
        current_team_leader = None
        
        # Process rows
        for row_idx in range(1, sheet.max_row + 1):
            # Column A: Team leader name and phone
            cell_a = sheet.cell(row=row_idx, column=1).value
            # Column B: Event name
            event_name = sheet.cell(row=row_idx, column=2).value
            # Column C: Date
            event_date_cell = sheet.cell(row=row_idx, column=3).value
            # Column D: Number of scanners
            scanners_count = sheet.cell(row=row_idx, column=4).value
            # Column E: Hours
            event_hours = sheet.cell(row=row_idx, column=5).value
            
            # Check if this row contains a team leader (has phone number)
            if cell_a and has_phone_number(cell_a):
                team_leader_name = extract_name(cell_a)
                if team_leader_name:
                    # Try to find team leader in the database by partial match
                    current_team_leader = find_team_leader(team_leader_name, team_leaders)
            
            # Process event information
            if event_name and event_date_cell and current_team_leader:
                # Convert date to proper format
                if isinstance(event_date_cell, datetime):
                    event_date = event_date_cell.date()
                else:
                    try:
                        # Try to parse date string
                        event_date = datetime.strptime(str(event_date_cell), "%Y-%m-%d").date()
                    except ValueError:
                        try:
                            # Try with day.month.year format
                            event_date = datetime.strptime(str(event_date_cell), "%d.%m.%Y").date()
                        except ValueError:
                            print(f"Invalid date format: {event_date_cell}")
                            continue
                
                # Create or get the event
                event, created = Event.objects.get_or_create(
                    name=event_name,
                    date=event_date,
                    leader=current_team_leader,
                    defaults={
                        'volunteers_required': int(scanners_count) if isinstance(scanners_count, (int, float)) else 0,
                        'duration_hours': float(event_hours) if isinstance(event_hours, (int, float)) else 0
                    }
                )
                
                if created:
                    print(f"Created event: {event.name} on {event.date}")
                else:
                    print(f"Event already exists: {event.name} on {event.date}")
                    # Update fields if needed
                    if isinstance(scanners_count, (int, float)):
                        event.volunteers_required = int(scanners_count)
                    if isinstance(event_hours, (int, float)):
                        event.duration_hours = float(event_hours)
                    event.save()
                
                current_event = event
            
            # If we have a valid event and the row contains a scanner name (not a team leader row)
            if current_event and cell_a and not has_phone_number(cell_a) and not event_name:
                scanner_name = str(cell_a).strip()
                if scanner_name:
                    # Try to find or create scanner
                    names = scanner_name.split()
                    if len(names) >= 1:
                        if len(names) == 1:
                            first_name = names[0]
                            last_name = ""
                        else:
                            first_name = names[0]
                            last_name = ' '.join(names[1:])
                        
                        scanner, created = Scanner.objects.get_or_create(
                            first_name=first_name,
                            last_name=last_name
                        )
                        
                        if created:
                            print(f"Created scanner: {scanner.first_name} {scanner.last_name}")
                        
                        # Add scanner to event if not already added
                        participant, p_created = EventParticipant.objects.get_or_create(
                            event=current_event,
                            volunteer=scanner
                        )
                        
                        if p_created:
                            print(f"Added {scanner.first_name} {scanner.last_name} to {current_event.name}")
                        
                        # Установка часов для участника
                        if current_event.duration_hours:
                            # Учитываем опоздание, если оно есть
                            late_hours = (participant.late_minutes or 0) / 60
                            awarded_hours = max(current_event.duration_hours - late_hours, 0)
                            
                            # Обновляем часы
                            participant.hours_awarded = awarded_hours
                            participant.save()
                            print(f"  - Установлено {awarded_hours} часов для {scanner.first_name} {scanner.last_name}")
        
        # Обновление часов для всех участников ивентов июня 2025
        print("\nОбновление часов для всех участников...")
        events = Event.objects.filter(date__year=2025, date__month=6)
        for event in events:
            if event.duration_hours:
                participants = EventParticipant.objects.filter(event=event)
                print(f"Обновление часов для ивента '{event.name}' ({event.date})")
                print(f"Длительность ивента: {event.duration_hours} часов")
                print(f"Количество участников: {participants.count()}")
                
                for participant in participants:
                    late_hours = (participant.late_minutes or 0) / 60
                    awarded_hours = max(event.duration_hours - late_hours, 0)
                    participant.hours_awarded = awarded_hours
                    participant.save()
                    print(f"  - {participant.volunteer.first_name} {participant.volunteer.last_name}: {awarded_hours} часов")
    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 