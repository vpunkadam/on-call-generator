#!/usr/bin/env python3
"""
SRE On-Call Schedule Generator for macOS
Generates on-call schedules for tier2, tier3, and upgrade shifts
"""

import datetime
import random
from collections import defaultdict
from typing import List, Dict, Tuple, Set
import os
import sys

class OnCallScheduler:
    def __init__(self):
        self.tier2_users = []
        self.tier3_users = []
        self.upgrade_users = []
        self.pto_dates = defaultdict(set)  # user -> set of dates
        
        # Define shifts
        self.shifts = {
            'tier2': {
                'morning': {'start': '11:00', 'end': '17:00', 'timezone': 'EST'},
                'evening': {'start': '17:00', 'end': '23:00', 'timezone': 'EST'}
            },
            'tier3': {
                'morning': {'start': '11:00', 'end': '17:00', 'timezone': 'EST'},
                'evening': {'start': '17:00', 'end': '23:00', 'timezone': 'EST'}
            },
            'upgrade': {
                'full': {'start': '12:00', 'end': '20:30', 'timezone': 'EST'}
            }
        }
        
    def load_users_from_file(self, filename: str) -> List[str]:
        """Load users from a text file, one per line"""
        users = []
        try:
            with open(filename, 'r') as f:
                users = [line.strip() for line in f if line.strip()]
        except FileNotFoundError:
            print(f"Error: File '{filename}' not found")
            sys.exit(1)
        return users
    
    def parse_date_range(self, date_range: str) -> Set[datetime.date]:
        """Parse date range in DD/MM/YYYY-DD/MM/YYYY format"""
        dates = set()
        try:
            start_str, end_str = date_range.strip().split('-')
            start = datetime.datetime.strptime(start_str.strip(), '%d/%m/%Y').date()
            end = datetime.datetime.strptime(end_str.strip(), '%d/%m/%Y').date()
            
            current = start
            while current <= end:
                dates.add(current)
                current += datetime.timedelta(days=1)
        except:
            print(f"Invalid date format: {date_range}")
            print("Please use DD/MM/YYYY-DD/MM/YYYY format")
        
        return dates
    
    def collect_pto(self):
        """Collect PTO information for all users"""
        print("\n=== PTO/Time Off Collection ===")
        all_users = set(self.tier2_users + self.tier3_users + self.upgrade_users)
        
        for user in sorted(all_users):
            print(f"\nEnter PTO for {user} (press Enter if none):")
            print("Format: DD/MM/YYYY-DD/MM/YYYY (can enter multiple ranges, one per line)")
            print("Press Enter with empty line when done")
            
            while True:
                pto_range = input("> ").strip()
                if not pto_range:
                    break
                
                pto_dates = self.parse_date_range(pto_range)
                if pto_dates:
                    self.pto_dates[user].update(pto_dates)
                    print(f"  Added {len(pto_dates)} PTO days for {user}")
    
    def is_user_available(self, user: str, date: datetime.date) -> bool:
        """Check if user is available on given date"""
        return date not in self.pto_dates.get(user, set())
    
    def get_month_weeks(self, year: int, month: int) -> List[Tuple[datetime.date, datetime.date]]:
        """Get all weeks (Mon-Sun) that contain days from the specified month"""
        import calendar
        
        # Get first and last day of month
        first_day = datetime.date(year, month, 1)
        last_day = datetime.date(year, month, calendar.monthrange(year, month)[1])
        
        # Find the Monday of the first week
        days_to_monday = first_day.weekday()
        first_monday = first_day - datetime.timedelta(days=days_to_monday)
        
        # Find the Sunday of the last week
        days_to_sunday = 6 - last_day.weekday()
        last_sunday = last_day + datetime.timedelta(days=days_to_sunday)
        
        # Generate all weeks
        weeks = []
        current_monday = first_monday
        while current_monday <= last_sunday:
            week_end = current_monday + datetime.timedelta(days=6)
            weeks.append((current_monday, week_end))
            current_monday += datetime.timedelta(days=7)
        
        return weeks
    
    def generate_schedule(self, year: int, month: int):
        """Generate the on-call schedule for a given month"""
        schedule = defaultdict(lambda: defaultdict(dict))
        
        # Track daily assignments to prevent overlaps
        daily_assignments = defaultdict(set)  # date -> set of users
        
        # Get all weeks for the month
        weeks = self.get_month_weeks(year, month)
        
        # Generate for each week
        for week_start, week_end in weeks:
            # Assign upgrade shift (one person for entire week)
            upgrade_user = self.assign_upgrade_shift(week_start, daily_assignments)
            if upgrade_user:
                current = week_start
                while current <= week_end:
                    schedule[current]['upgrade']['full'] = upgrade_user
                    daily_assignments[current].add(upgrade_user)
                    current += datetime.timedelta(days=1)
            
            # Assign tier2 and tier3 shifts for each day
            current = week_start
            while current <= week_end:
                # Tier2 shifts
                self.assign_daily_shifts('tier2', current, self.tier2_users, 
                                       schedule, daily_assignments)
                
                # Tier3 shifts
                self.assign_daily_shifts('tier3', current, self.tier3_users, 
                                       schedule, daily_assignments)
                
                current += datetime.timedelta(days=1)
        
        return schedule
    
    def assign_upgrade_shift(self, week_start: datetime.date, 
                           daily_assignments: Dict) -> str:
        """Assign upgrade shift for entire week"""
        available_users = []
        
        for user in self.upgrade_users:
            # Check if user is available for entire week
            available_all_week = True
            for day in range(7):
                date = week_start + datetime.timedelta(days=day)
                if not self.is_user_available(user, date) or user in daily_assignments[date]:
                    available_all_week = False
                    break
            
            if available_all_week:
                available_users.append(user)
        
        if available_users:
            return random.choice(available_users)
        else:
            print(f"Warning: No available users for upgrade shift week of {week_start}")
            return None
    
    def assign_daily_shifts(self, tier: str, date: datetime.date, 
                          users: List[str], schedule: Dict, 
                          daily_assignments: Dict):
        """Assign morning and evening shifts for a tier"""
        available_users = [u for u in users 
                         if self.is_user_available(u, date) 
                         and u not in daily_assignments[date]]
        
        if len(available_users) >= 2:
            # Randomly assign two different users
            random.shuffle(available_users)
            schedule[date][tier]['morning'] = available_users[0]
            schedule[date][tier]['evening'] = available_users[1]
            daily_assignments[date].add(available_users[0])
            daily_assignments[date].add(available_users[1])
        elif len(available_users) == 1:
            # Only one user available - assign to morning
            schedule[date][tier]['morning'] = available_users[0]
            daily_assignments[date].add(available_users[0])
            print(f"Warning: Only one user available for {tier} on {date}")
        else:
            print(f"Warning: No users available for {tier} on {date}")
    
    def print_schedule(self, schedule: Dict, year: int, month: int):
        """Print the schedule in a readable format"""
        import calendar
        
        month_name = calendar.month_name[month]
        print(f"\n=== ON-CALL SCHEDULE FOR {month_name.upper()} {year} ===\n")
        
        # Get all weeks for the month
        weeks = self.get_month_weeks(year, month)
        
        for week_num, (week_start, week_end) in enumerate(weeks, 1):
            print(f"\n--- Week {week_num}: {week_start.strftime('%B %d')} - {week_end.strftime('%B %d, %Y')} ---")
            
            current = week_start
            while current <= week_end:
                if current in schedule:
                    print(f"\n{current.strftime('%A, %B %d')}:")
                    
                    # Print tier2
                    if 'tier2' in schedule[current]:
                        print("  Tier 2:")
                        if 'morning' in schedule[current]['tier2']:
                            print(f"    11am-5pm EST: {schedule[current]['tier2']['morning']}")
                        if 'evening' in schedule[current]['tier2']:
                            print(f"    5pm-11pm EST: {schedule[current]['tier2']['evening']}")
                    
                    # Print tier3
                    if 'tier3' in schedule[current]:
                        print("  Tier 3:")
                        if 'morning' in schedule[current]['tier3']:
                            print(f"    11am-5pm EST: {schedule[current]['tier3']['morning']}")
                        if 'evening' in schedule[current]['tier3']:
                            print(f"    5pm-11pm EST: {schedule[current]['tier3']['evening']}")
                    
                    # Print upgrade
                    if 'upgrade' in schedule[current] and 'full' in schedule[current]['upgrade']:
                        print(f"  Upgrade: 12pm-8:30pm EST: {schedule[current]['upgrade']['full']}")
                
                current += datetime.timedelta(days=1)
    
    def export_to_csv(self, schedule: Dict, filename: str = "oncall_schedule.csv"):
        """Export schedule to CSV file"""
        import csv
        
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Date', 'Day', 'Schedule', 'Shift', 'Time', 'User'])
            
            dates = sorted(schedule.keys())
            for date in dates:
                day_name = date.strftime('%A')
                date_str = date.strftime('%Y-%m-%d')
                
                for tier in ['tier2', 'tier3', 'upgrade']:
                    if tier in schedule[date]:
                        if tier == 'upgrade':
                            if 'full' in schedule[date][tier]:
                                writer.writerow([date_str, day_name, tier, 'full', 
                                               '12:00pm-8:30pm EST', 
                                               schedule[date][tier]['full']])
                        else:
                            if 'morning' in schedule[date][tier]:
                                writer.writerow([date_str, day_name, tier, 'morning', 
                                               '11:00am-5:00pm EST', 
                                               schedule[date][tier]['morning']])
                            if 'evening' in schedule[date][tier]:
                                writer.writerow([date_str, day_name, tier, 'evening', 
                                               '5:00pm-11:00pm EST', 
                                               schedule[date][tier]['evening']])
        
        print(f"\nSchedule exported to {filename}")

def main():
    print("=== SRE On-Call Schedule Generator ===\n")
    
    scheduler = OnCallScheduler()
    
    # Load users from files
    print("Enter filename for Tier 2 users:")
    tier2_file = input("> ").strip()
    scheduler.tier2_users = scheduler.load_users_from_file(tier2_file)
    print(f"Loaded {len(scheduler.tier2_users)} Tier 2 users")
    
    print("\nEnter filename for Tier 3 users:")
    tier3_file = input("> ").strip()
    scheduler.tier3_users = scheduler.load_users_from_file(tier3_file)
    print(f"Loaded {len(scheduler.tier3_users)} Tier 3 users")
    
    print("\nEnter filename for Upgrade shift users:")
    upgrade_file = input("> ").strip()
    scheduler.upgrade_users = scheduler.load_users_from_file(upgrade_file)
    print(f"Loaded {len(scheduler.upgrade_users)} Upgrade users")
    
    # Collect PTO information
    scheduler.collect_pto()
    
    # Get schedule parameters
    print("\n=== Schedule Parameters ===")
    print("Enter month to schedule (MM/YYYY):")
    month_year_str = input("> ").strip()
    try:
        month, year = month_year_str.split('/')
        month = int(month)
        year = int(year)
        if month < 1 or month > 12:
            raise ValueError("Invalid month")
    except:
        print("Invalid format. Using current month.")
        now = datetime.datetime.now()
        month = now.month
        year = now.year
    
    # Show what weeks will be included
    import calendar
    weeks = scheduler.get_month_weeks(year, month)
    print(f"\nScheduling for {calendar.month_name[month]} {year}")
    print(f"This includes {len(weeks)} weeks:")
    for i, (start, end) in enumerate(weeks, 1):
        print(f"  Week {i}: {start.strftime('%B %d')} - {end.strftime('%B %d')}")
    
    # Generate schedule
    print("\nGenerating schedule...")
    schedule = scheduler.generate_schedule(year, month)
    
    # Display schedule
    scheduler.print_schedule(schedule, year, month)
    
    # Export option
    print("\nExport to CSV? (y/n):")
    if input("> ").strip().lower() == 'y':
        print("Enter filename (default: oncall_schedule.csv):")
        filename = input("> ").strip() or "oncall_schedule.csv"
        scheduler.export_to_csv(schedule, filename)

if __name__ == "__main__":
    main()