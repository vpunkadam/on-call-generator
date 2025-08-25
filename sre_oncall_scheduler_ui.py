#!/usr/bin/env python3
"""
SRE On-Call Schedule Generator with Web UI - Enhanced with Prior Month Import
Generates on-call schedules for tier2, tier3, and upgrade shifts
"""

import datetime
import random
from collections import defaultdict
from typing import List, Dict, Tuple, Set, Optional
import json
from flask import Flask, render_template_string, request, jsonify
import webbrowser
import threading
import time
import os
import sys

app = Flask(__name__)

class OnCallScheduler:
    def __init__(self):
        self.tier2_users = []
        self.tier3_users = []
        self.upgrade_users = []
        self.pto_dates = defaultdict(set)  # user -> set of dates
        self.upgrade_rotation_queue = []  # Track upgrade rotation
        self.last_upgrade_user = None  # Track last scheduled upgrade user
        self.tier3_morning_rotation_queue = []  # Track tier3 morning rotation
        self.last_tier3_morning_user = None  # Track last scheduled tier3 morning user
        self.tier3_evening_rotation_queue = []  # Track tier3 evening rotation
        self.last_tier3_evening_user = None  # Track last scheduled tier3 evening user
        
        # Load cumulative shift history from file if it exists
        self.cumulative_shift_counts = defaultdict(int)  # Historical total across all months
        self.load_cumulative_shift_history()
        
        self.shift_counts = defaultdict(int)  # Track shifts for current month only
        
        # Store prior month's last week assignments for continuity
        self.prior_month_last_week = {
            'upgrade': None,
            'tier3_morning': None,
            'tier3_evening': None,
            'date': None  # To track when this was set
        }
        
        # Coverage warnings and fallback tracking
        self.coverage_warnings = []  # Track any coverage issues
        self.fallback_assignments = []  # Track emergency/fallback assignments
        
        # Track weekly assignments per user per month (for 2-week limit)
        self.monthly_weekly_assignments = defaultdict(int)  # user -> count of weekly shifts this month
        
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
    
    def import_prior_month_schedule(self, schedule_data: Dict) -> Dict:
        """Import prior month's schedule and extract last week's assignments + shift counts"""
        result = {
            'success': False,
            'message': '',
            'last_week_assignments': {},
            'shift_counts': {}
        }
        
        try:
            # Find the last week's assignments
            if not schedule_data:
                result['message'] = 'No schedule data provided'
                return result
            
            # Convert schedule data and find the latest dates
            dates = []
            for date_str in schedule_data.keys():
                try:
                    date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                    dates.append(date)
                except:
                    continue
            
            if not dates:
                result['message'] = 'No valid dates found in schedule'
                return result
            
            # Find the last Monday in the schedule
            max_date = max(dates)
            last_monday = max_date - datetime.timedelta(days=max_date.weekday())
            
            # Extract assignments from the last week
            last_week_assignments = {
                'upgrade': None,
                'tier3_morning': None,
                'tier3_evening': None,
                'week_start': last_monday.strftime('%Y-%m-%d')
            }
            
            # Count all shifts from the imported data
            imported_shift_counts = defaultdict(int)
            
            # Process all days in schedule
            for date_str in schedule_data:
                day_schedule = schedule_data[date_str]
                
                # Count all shifts
                for tier in ['tier2', 'tier3', 'upgrade']:
                    if tier in day_schedule:
                        if tier == 'upgrade':
                            if 'full' in day_schedule[tier]:
                                user = day_schedule[tier]['full']
                                if user:
                                    imported_shift_counts[user] += 1
                        else:
                            for shift in ['morning', 'evening']:
                                if shift in day_schedule[tier]:
                                    user = day_schedule[tier][shift]
                                    if user:
                                        imported_shift_counts[user] += 1
                
                # Check if this is in the last week
                try:
                    check_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                    if last_monday <= check_date <= last_monday + datetime.timedelta(days=6):
                        # Check upgrade
                        if 'upgrade' in day_schedule and 'full' in day_schedule['upgrade']:
                            last_week_assignments['upgrade'] = day_schedule['upgrade']['full']
                        
                        # Check tier3 morning
                        if 'tier3' in day_schedule and 'morning' in day_schedule['tier3']:
                            last_week_assignments['tier3_morning'] = day_schedule['tier3']['morning']
                        
                        # Check tier3 evening
                        if 'tier3' in day_schedule and 'evening' in day_schedule['tier3']:
                            last_week_assignments['tier3_evening'] = day_schedule['tier3']['evening']
                except:
                    continue
            
            # Update internal state
            self.prior_month_last_week = {
                'upgrade': last_week_assignments['upgrade'],
                'tier3_morning': last_week_assignments['tier3_morning'],
                'tier3_evening': last_week_assignments['tier3_evening'],
                'date': datetime.datetime.now()
            }
            
            # Also set the last user properties for the scheduler
            self.last_upgrade_user = last_week_assignments['upgrade']
            self.last_tier3_morning_user = last_week_assignments['tier3_morning']
            self.last_tier3_evening_user = last_week_assignments['tier3_evening']
            
            # Update shift counts from imported data
            self.shift_counts = imported_shift_counts
            
            result['success'] = True
            result['message'] = f"Successfully imported prior month's schedule. Last week started {last_monday.strftime('%B %d, %Y')}. Loaded shift counts for {len(imported_shift_counts)} users."
            result['last_week_assignments'] = last_week_assignments
            result['shift_counts'] = dict(imported_shift_counts)
            
            # Save history for persistence
            self.save_shift_history_from_import(imported_shift_counts)
            
            return result
            
        except Exception as e:
            result['message'] = f"Error importing schedule: {str(e)}"
            return result
    
    def validate_schedule(self, schedule: Dict, year: int, month: int) -> Dict:
        """Validate the generated schedule for requirement violations"""
        validation_errors = {
            'critical': [],  # Must fix - like PTO violations
            'warnings': [],  # Should review - like excessive shifts
            'info': []       # Informational - like coverage gaps already handled
        }
        
        # Track users and their assignments
        user_assignments = defaultdict(list)
        daily_user_counts = defaultdict(lambda: defaultdict(int))
        
        # Collect all assignments
        for date, day_schedule in schedule.items():
            for tier in ['tier2', 'tier3', 'upgrade']:
                if tier in day_schedule:
                    if tier == 'upgrade':
                        if 'full' in day_schedule[tier]:
                            user = day_schedule[tier]['full']
                            # Strip any tags like (DOUBLE) or (EMERGENCY)
                            base_user = user.split(' (')[0] if ' (' in user else user
                            user_assignments[base_user].append((date, tier, 'full'))
                            daily_user_counts[date][base_user] += 1
                    else:
                        for shift in ['morning', 'evening']:
                            if shift in day_schedule[tier]:
                                user = day_schedule[tier][shift]
                                # Strip any tags
                                base_user = user.split(' (')[0] if ' (' in user else user
                                user_assignments[base_user].append((date, tier, shift))
                                daily_user_counts[date][base_user] += 1
        
        # Validation 1: Check PTO violations
        for user, assignments in user_assignments.items():
            if user in self.pto_dates:
                for date, tier, shift in assignments:
                    if date in self.pto_dates[user]:
                        validation_errors['critical'].append(
                            f"{user} is scheduled for {tier} {shift} on {date.strftime('%Y-%m-%d')} but is on PTO"
                        )
        
        # Validation 2: Check for users with more than 2 shifts on same day
        for date, user_counts in daily_user_counts.items():
            for user, count in user_counts.items():
                if count > 2:
                    validation_errors['critical'].append(
                        f"{user} is assigned {count} shifts on {date.strftime('%Y-%m-%d')} (max should be 2)"
                    )
                elif count == 2:
                    # This is a double shift - should already be marked but verify
                    validation_errors['warnings'].append(
                        f"{user} has a double shift on {date.strftime('%Y-%m-%d')}"
                    )
        
        # Validation 3: Check for upgrade tier violations
        for user, assignments in user_assignments.items():
            for date, tier, shift in assignments:
                if tier == 'upgrade' and user not in self.upgrade_users:
                    # Check if it's not an emergency assignment
                    actual_assignment = schedule[date][tier][shift]
                    if 'EMERGENCY' not in actual_assignment:
                        validation_errors['critical'].append(
                            f"{user} is assigned to upgrade on {date.strftime('%Y-%m-%d')} but is not an upgrade user"
                        )
        
        # Validation 4: Check for excessive shift imbalance (accounting for PTO)
        if user_assignments:
            # Calculate expected working days for each user (total days minus PTO days)
            import calendar
            num_days = calendar.monthrange(year, month)[1]
            
            # Count PTO days for each user in this month
            user_pto_days = {}
            all_users = set(self.tier2_users + self.tier3_users + self.upgrade_users)
            
            for user in all_users:
                # self.pto_dates[user] already contains the actual date objects for this user's PTO
                if user in self.pto_dates:
                    # Count how many of their PTO dates fall within this month
                    pto_count = sum(1 for date in self.pto_dates[user] 
                                  if date.year == year and date.month == month)
                    user_pto_days[user] = pto_count
                else:
                    user_pto_days[user] = 0
            
            # Calculate shift ratios (shifts per available day)
            # But exclude users with more than 2 PTO days from fairness comparisons
            shift_ratios = {}
            excluded_users = []
            
            for user in all_users:
                pto_days = user_pto_days.get(user, 0)
                
                # Exclude users with more than 2 PTO days from fairness metrics
                if pto_days > 2:
                    excluded_users.append(user)
                    continue
                    
                available_days = num_days - pto_days
                if available_days > 0:
                    shifts = len(user_assignments.get(user, []))
                    shift_ratios[user] = shifts / available_days
                else:
                    shift_ratios[user] = 0  # User on PTO entire month
            
            # Check imbalance based on ratios, not absolute counts
            # Only among users not excluded for excessive PTO
            if shift_ratios:
                max_ratio = max(shift_ratios.values())
                min_ratio = min([r for r in shift_ratios.values() if r > 0] or [0])  # Exclude users with 0 ratio
                
                # Only warn if the ratio difference is significant (>30% difference)
                if max_ratio > 0 and min_ratio > 0 and (max_ratio - min_ratio) / min_ratio > 0.3:
                    max_user = max(shift_ratios, key=shift_ratios.get)
                    min_user = min([u for u, r in shift_ratios.items() if r > 0], key=shift_ratios.get)
                    max_shifts = len(user_assignments.get(max_user, []))
                    min_shifts = len(user_assignments.get(min_user, []))
                    validation_errors['warnings'].append(
                        f"Shift imbalance detected: {max_user} has {max_shifts} shifts with {user_pto_days.get(max_user, 0)} PTO days, "
                        f"{min_user} has {min_shifts} shifts with {user_pto_days.get(min_user, 0)} PTO days"
                    )
            
            # Add info about excluded users if any
            if excluded_users:
                validation_errors['info'].append(
                    f"Users excluded from fairness metrics due to >2 PTO days: {', '.join(excluded_users)}"
                )
        
        # Validation 5: Check for unfilled critical shifts
        import calendar
        num_days = calendar.monthrange(year, month)[1]
        
        for day in range(1, num_days + 1):
            date = datetime.date(year, month, day)
            if date in schedule:
                # Check upgrade coverage
                if 'upgrade' not in schedule[date] or 'full' not in schedule[date]['upgrade']:
                    validation_errors['warnings'].append(
                        f"No upgrade coverage on {date.strftime('%Y-%m-%d')}"
                    )
                
                # Check tier3 coverage
                if 'tier3' in schedule[date]:
                    if 'morning' not in schedule[date]['tier3']:
                        validation_errors['info'].append(
                            f"No tier3 morning coverage on {date.strftime('%Y-%m-%d')}"
                        )
                    if 'evening' not in schedule[date]['tier3']:
                        validation_errors['info'].append(
                            f"No tier3 evening coverage on {date.strftime('%Y-%m-%d')}"
                        )
        
        # Validation 6: Check for back-to-back weekly assignments and weekly limit
        weekly_assignments = defaultdict(list)
        user_weekly_count = defaultdict(int)  # Track total weekly assignments per user
        
        for date, day_schedule in sorted(schedule.items()):
            week_start = date - datetime.timedelta(days=date.weekday())
            
            if 'upgrade' in day_schedule and 'full' in day_schedule['upgrade']:
                user = day_schedule['upgrade']['full'].split(' (')[0]
                if week_start not in [w for w, _ in weekly_assignments[('upgrade', user)]]:
                    weekly_assignments[('upgrade', user)].append((week_start, date))
                    user_weekly_count[user] += 1
            
            if 'tier3' in day_schedule:
                if 'morning' in day_schedule['tier3']:
                    user = day_schedule['tier3']['morning'].split(' (')[0]
                    if week_start not in [w for w, _ in weekly_assignments[('tier3_morning', user)]]:
                        weekly_assignments[('tier3_morning', user)].append((week_start, date))
                        user_weekly_count[user] += 1
                if 'evening' in day_schedule['tier3']:
                    user = day_schedule['tier3']['evening'].split(' (')[0]
                    if week_start not in [w for w, _ in weekly_assignments[('tier3_evening', user)]]:
                        weekly_assignments[('tier3_evening', user)].append((week_start, date))
                        user_weekly_count[user] += 1
        
        # Check for users exceeding 2 weekly shifts per month
        for user, count in user_weekly_count.items():
            if count > 2:
                validation_errors['critical'].append(
                    f"{user} has {count} weekly shifts this month (maximum allowed is 2)"
                )
        
        # Check if someone got 2 shifts while others have 0 (fairness check)
        if user_weekly_count:
            max_weekly = max(user_weekly_count.values())
            # Get all users who should have weekly shifts
            all_weekly_users = set(self.tier3_users + self.upgrade_users)
            users_with_zero_weekly = []
            
            for user in all_weekly_users:
                if user_weekly_count.get(user, 0) == 0:
                    # Check if they were available for any week
                    if user in self.pto_dates:
                        # Count PTO days in month
                        pto_days_in_month = sum(1 for date in self.pto_dates[user] 
                                               if date.year == year and date.month == month)
                        # If they have less than 7 consecutive PTO days, they could have taken a shift
                        if pto_days_in_month < 7:
                            users_with_zero_weekly.append(user)
                    else:
                        users_with_zero_weekly.append(user)
            
            if max_weekly >= 2 and users_with_zero_weekly:
                validation_errors['warnings'].append(
                    f"Fairness issue: Some users have 2 weekly shifts while {', '.join(users_with_zero_weekly)} have 0"
                )
        
        # Check for consecutive weeks
        for (tier, user), weeks in weekly_assignments.items():
            weeks.sort()
            for i in range(len(weeks) - 1):
                week1_start, _ = weeks[i]
                week2_start, _ = weeks[i + 1]
                if (week2_start - week1_start).days == 7:
                    validation_errors['warnings'].append(
                        f"{user} has back-to-back {tier} assignments (weeks of {week1_start.strftime('%Y-%m-%d')} and {week2_start.strftime('%Y-%m-%d')})"
                    )
        
        # Print validation summary
        if validation_errors['critical'] or validation_errors['warnings'] or validation_errors['info']:
            print("\n=== Schedule Validation Report ===")
            if validation_errors['critical']:
                print("\n❌ CRITICAL ERRORS (Must Fix):")
                for error in validation_errors['critical']:
                    print(f"  - {error}")
            if validation_errors['warnings']:
                print("\n⚠️  WARNINGS (Should Review):")
                for warning in validation_errors['warnings']:
                    print(f"  - {warning}")
            if validation_errors['info']:
                print("\nℹ️  INFO (Already Handled):")
                # Only show first 5 info messages to avoid clutter
                for info in validation_errors['info'][:5]:
                    print(f"  - {info}")
                if len(validation_errors['info']) > 5:
                    print(f"  ... and {len(validation_errors['info']) - 5} more")
        else:
            print("\n✅ Schedule validation passed - all requirements met")
        
        return validation_errors
    
    def save_shift_history_from_import(self, shift_counts: Dict):
        """Save imported shift counts to JSON file for persistence"""
        import json
        
        filename = 'cumulative_shift_history.json'
        
        try:
            # Load existing history if it exists
            existing_data = {}
            try:
                with open(filename, 'r') as f:
                    existing_data = json.load(f)
            except FileNotFoundError:
                pass
            
            # Update with new counts
            existing_data['last_import'] = datetime.datetime.now().isoformat()
            existing_data['cumulative_counts'] = dict(shift_counts)
            
            with open(filename, 'w') as f:
                json.dump(existing_data, f, indent=2)
            
            return True
        except Exception as e:
            print(f"Error saving shift history: {e}")
            return False
    
    def load_cumulative_shift_history(self):
        """Load cumulative shift counts from saved file"""
        import json
        
        filename = 'cumulative_shift_history.json'
        
        try:
            with open(filename, 'r') as f:
                data = json.load(f)
                counts = data.get('cumulative_counts', {})
                self.cumulative_shift_counts = defaultdict(int, counts)
                print(f"Loaded cumulative shift history for {len(counts)} users")
                return counts
        except FileNotFoundError:
            print("No existing shift history found, starting fresh")
            return {}
        except Exception as e:
            print(f"Error loading shift history: {e}")
            return {}
    
    def update_and_save_cumulative_history(self):
        """Update cumulative counts with current month's shifts and save to file"""
        import json
        
        # Update cumulative counts with this month's shifts
        for user, count in self.shift_counts.items():
            self.cumulative_shift_counts[user] += count
        
        filename = 'cumulative_shift_history.json'
        
        try:
            data = {
                'last_update': datetime.datetime.now().isoformat(),
                'cumulative_counts': dict(self.cumulative_shift_counts)
            }
            
            with open(filename, 'w') as f:
                json.dump(data, f, indent=2)
            
            print(f"Saved cumulative shift history for {len(self.cumulative_shift_counts)} users")
            return True
        except Exception as e:
            print(f"Error saving cumulative history: {e}")
            return False
    
    def load_users_from_file(self, filename: str) -> List[str]:
        """Load users from a text file, one per line"""
        users = []
        try:
            with open(filename, 'r') as f:
                users = [line.strip() for line in f if line.strip()]
        except FileNotFoundError:
            print(f"Error: File '{filename}' not found")
        return users
    
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
    
    def is_user_available(self, user: str, date: datetime.date) -> bool:
        """Check if user is available on given date"""
        return date not in self.pto_dates.get(user, set())
    
    def is_user_available_for_week(self, user: str, week_start: datetime.date, daily_assignments: Dict) -> bool:
        """Check if user is available for an entire week"""
        for day in range(7):
            date = week_start + datetime.timedelta(days=day)
            if not self.is_user_available(user, date) or user in daily_assignments.get(date, set()):
                return False
        return True
    
    def generate_schedule(self, year: int, month: int):
        """Generate the on-call schedule for a given month"""
        schedule = defaultdict(lambda: defaultdict(dict))
        
        # Track daily assignments to prevent overlaps
        daily_assignments = defaultdict(set)  # date -> set of users
        
        # Reset shift counts for THIS MONTH only (cumulative counts remain)
        self.shift_counts = defaultdict(int)
        
        # Reset monthly weekly assignment counter
        self.monthly_weekly_assignments = defaultdict(int)
        
        # Initialize rotation queues
        self.upgrade_rotation_queue = self.upgrade_users.copy()
        random.shuffle(self.upgrade_rotation_queue)
        
        self.tier3_morning_rotation_queue = self.tier3_users.copy()
        random.shuffle(self.tier3_morning_rotation_queue)
        
        self.tier3_evening_rotation_queue = self.tier3_users.copy()
        random.shuffle(self.tier3_evening_rotation_queue)
        
        # Get all weeks for the month
        weeks = self.get_month_weeks(year, month)
        
        # Check if we have prior month data and if it's relevant
        if self.prior_month_last_week['date']:
            print(f"Using prior month's last week assignments:")
            print(f"  Upgrade: {self.last_upgrade_user}")
            print(f"  Tier3 Morning: {self.last_tier3_morning_user}")
            print(f"  Tier3 Evening: {self.last_tier3_evening_user}")
        
        # Generate for each week
        for week_start, week_end in weeks:
            # Assign upgrade shift (one person for entire week)
            upgrade_user = self.assign_weekly_shift_with_rotation(
                'upgrade', week_start, daily_assignments, 
                self.upgrade_rotation_queue, self.last_upgrade_user,
                self.upgrade_users
            )
            if upgrade_user:
                current = week_start
                while current <= week_end:
                    schedule[current]['upgrade']['full'] = upgrade_user
                    daily_assignments[current].add(upgrade_user)
                    current += datetime.timedelta(days=1)
                self.last_upgrade_user = upgrade_user
                # Count this as 7 shifts (one per day of the week)
                self.shift_counts[upgrade_user] += 7
                # Increment weekly assignment counter
                self.monthly_weekly_assignments[upgrade_user] += 1
            
            # Assign tier3 weekly shifts
            # Morning shift
            tier3_morning_user = self.assign_weekly_shift_with_rotation(
                'tier3_morning', week_start, daily_assignments,
                self.tier3_morning_rotation_queue, self.last_tier3_morning_user,
                self.tier3_users
            )
            if tier3_morning_user:
                current = week_start
                while current <= week_end:
                    schedule[current]['tier3']['morning'] = tier3_morning_user
                    daily_assignments[current].add(tier3_morning_user)
                    current += datetime.timedelta(days=1)
                self.last_tier3_morning_user = tier3_morning_user
                # Count this as 7 shifts
                self.shift_counts[tier3_morning_user] += 7
                # Increment weekly assignment counter
                self.monthly_weekly_assignments[tier3_morning_user] += 1
            
            # Evening shift
            tier3_evening_user = self.assign_weekly_shift_with_rotation(
                'tier3_evening', week_start, daily_assignments,
                self.tier3_evening_rotation_queue, self.last_tier3_evening_user,
                self.tier3_users
            )
            if tier3_evening_user:
                current = week_start
                while current <= week_end:
                    schedule[current]['tier3']['evening'] = tier3_evening_user
                    daily_assignments[current].add(tier3_evening_user)
                    current += datetime.timedelta(days=1)
                self.last_tier3_evening_user = tier3_evening_user
                # Count this as 7 shifts
                self.shift_counts[tier3_evening_user] += 7
                # Increment weekly assignment counter
                self.monthly_weekly_assignments[tier3_evening_user] += 1
            
            # Assign tier2 daily shifts
            current = week_start
            while current <= week_end:
                self.assign_daily_shifts_with_fairness('tier2', current, self.tier2_users, 
                                       schedule, daily_assignments)
                current += datetime.timedelta(days=1)
        
        # Validate the generated schedule
        validation_errors = self.validate_schedule(schedule, year, month)
        
        # Print fairness report
        self.print_fairness_report(year, month)
        
        # If there are critical validation errors, add them to warnings
        if validation_errors['critical']:
            for error in validation_errors['critical']:
                self.coverage_warnings.append(f"VALIDATION ERROR: {error}")
        
        # Update cumulative counts with this month's shifts and save to file
        self.update_and_save_cumulative_history()
        
        return schedule
    
    def assign_weekly_shift_with_rotation(self, shift_type: str, week_start: datetime.date, 
                                         daily_assignments: Dict, rotation_queue: List[str],
                                         last_user: str, all_users: List[str]) -> str:
        """Assign weekly shift with fair rotation - everyone gets 1 before anyone gets 2"""
        # Rules:
        # 1. Is available for the entire week
        # 2. Wasn't scheduled last week (no back-to-back)
        # 3. Prioritize users with 0 weekly shifts, then 1, up to max 2
        # 4. Within same priority level, follow rotation queue
        
        # Check if everyone has at least 1 weekly shift
        users_with_zero_shifts = [u for u in all_users if self.monthly_weekly_assignments.get(u, 0) == 0]
        users_with_one_shift = [u for u in all_users if self.monthly_weekly_assignments.get(u, 0) == 1]
        
        # Categorize available users by their weekly shift count
        available_by_priority = {0: [], 1: []}
        
        for user in rotation_queue:
            # Skip if this user was scheduled last week (unless last_user is None)
            if last_user and user == last_user:
                continue
            
            # Skip if user already has 2 weekly shifts this month
            if self.monthly_weekly_assignments[user] >= 2:
                continue
                
            # Check if user is available for entire week
            available_all_week = True
            for day in range(7):
                date = week_start + datetime.timedelta(days=day)
                if not self.is_user_available(user, date) or user in daily_assignments[date]:
                    available_all_week = False
                    break
            
            if available_all_week:
                shift_count = self.monthly_weekly_assignments.get(user, 0)
                if shift_count in available_by_priority:
                    available_by_priority[shift_count].append(user)
        
        # Select user with priority: 0 shifts > 1 shift > no selection if everyone has 1+
        selected_user = None
        
        # First priority: users with 0 weekly shifts
        if available_by_priority[0]:
            selected_user = available_by_priority[0][0]
        # Second priority: users with 1 weekly shift (only if no one has 0)
        elif available_by_priority[1]:
            # Only assign a 2nd shift if everyone has at least 1
            if not users_with_zero_shifts or all(u in rotation_queue and not self.is_user_available_for_week(u, week_start, daily_assignments) for u in users_with_zero_shifts):
                selected_user = available_by_priority[1][0]
        
        if selected_user:
            rotation_queue.remove(selected_user)
            
            # If queue is empty, refill it (new rotation cycle)
            if not rotation_queue:
                rotation_queue.extend(all_users)
                # Remove the just-scheduled user to avoid immediate repeat
                if selected_user in rotation_queue:
                    rotation_queue.remove(selected_user)
                random.shuffle(rotation_queue)
            
            return selected_user
        
        # If no one from queue is available, try anyone except last week's person
        # Still respect the "everyone gets 1 before anyone gets 2" rule
        available_by_priority = {0: [], 1: []}
        
        for user in all_users:
            # Skip if this user was scheduled last week (unless last_user is None)
            if last_user and user == last_user:
                continue
            
            # Skip if user already has 2 weekly shifts this month
            if self.monthly_weekly_assignments[user] >= 2:
                continue
                
            if self.is_user_available_for_week(user, week_start, daily_assignments):
                shift_count = self.monthly_weekly_assignments.get(user, 0)
                if shift_count in available_by_priority:
                    available_by_priority[shift_count].append(user)
        
        # Prioritize users with 0 shifts, then 1 shift
        if available_by_priority[0]:
            selected_user = random.choice(available_by_priority[0])
        elif available_by_priority[1]:
            # Check if everyone has at least 1 shift before giving someone a 2nd
            users_with_zero = [u for u in all_users if self.monthly_weekly_assignments.get(u, 0) == 0]
            if not users_with_zero or all(not self.is_user_available_for_week(u, week_start, daily_assignments) for u in users_with_zero):
                selected_user = random.choice(available_by_priority[1])
            else:
                # Log that we're skipping 2nd assignments because some users still have 0
                print(f"Note: Skipping 2nd weekly assignment for week of {week_start} - some users still have 0 weekly shifts")
                return None
        
        if selected_user:
            # Update queue to reflect this choice
            if selected_user in rotation_queue:
                rotation_queue.remove(selected_user)
            return selected_user
        
        # Last resort: if no one else is available, allow last week's person
        if last_user and last_user in all_users:
            available_all_week = True
            for day in range(7):
                date = week_start + datetime.timedelta(days=day)
                if not self.is_user_available(last_user, date) or \
                   last_user in daily_assignments[date]:
                    available_all_week = False
                    break
            
            if available_all_week:
                print(f"Warning: {last_user} scheduled for back-to-back {shift_type} weeks due to availability constraints")
                return last_user
        
        print(f"Warning: No available users for {shift_type} shift week of {week_start}")
        return None
    
    def assign_daily_shifts_with_fairness(self, tier: str, date: datetime.date, 
                                          users: List[str], schedule: Dict, 
                                          daily_assignments: Dict):
        """Assign morning and evening shifts for a tier with fairness consideration"""
        # Get available users for this date
        available_users = [u for u in users 
                         if self.is_user_available(u, date) 
                         and u not in daily_assignments[date]]
        
        if len(available_users) >= 2:
            # Sort by CUMULATIVE shift count (ascending) to prioritize users with fewer historical shifts
            # This ensures long-term fairness across months
            available_users.sort(key=lambda u: self.cumulative_shift_counts[u] + self.shift_counts[u])
            
            # Assign the two users with the least shifts
            schedule[date][tier]['morning'] = available_users[0]
            schedule[date][tier]['evening'] = available_users[1]
            daily_assignments[date].add(available_users[0])
            daily_assignments[date].add(available_users[1])
            
            # Update shift counts
            self.shift_counts[available_users[0]] += 1
            self.shift_counts[available_users[1]] += 1
            
        elif len(available_users) == 1:
            # Only one user available - assign to morning
            schedule[date][tier]['morning'] = available_users[0]
            daily_assignments[date].add(available_users[0])
            self.shift_counts[available_users[0]] += 1
            
            # Add warning about insufficient coverage
            warning = f"WARNING: Only 1 user available for {tier} on {date.strftime('%Y-%m-%d')} - evening shift unfilled"
            self.coverage_warnings.append(warning)
            
            # Try fallback: find user with least shifts who can do double shift
            self.attempt_fallback_coverage(tier, date, 'evening', users, schedule, daily_assignments, available_users[0])
        else:
            # No users available - try emergency fallback
            warning = f"CRITICAL: No users available for {tier} on {date.strftime('%Y-%m-%d')}"
            self.coverage_warnings.append(warning)
            
            # Attempt emergency coverage
            self.attempt_emergency_coverage(tier, date, users, schedule, daily_assignments)
    
    def attempt_fallback_coverage(self, tier: str, date: datetime.date, shift: str, 
                                 users: List[str], schedule: Dict, daily_assignments: Dict,
                                 already_assigned: str):
        """Try to find fallback coverage when normal assignment fails"""
        # First try: users from other tiers who can cover
        # Second try: allow same user to do double shift if absolutely necessary
        
        # Special rule: Only upgrade users can cover upgrade shifts
        if tier == 'upgrade':
            # For upgrade, only try double shift - no cross-tier coverage allowed
            pass
        else:
            # For tier2/tier3, try users from other non-upgrade tiers
            all_users = set(self.tier2_users + self.tier3_users)
            other_tier_users = [u for u in all_users if u not in users]
            
            for user in sorted(other_tier_users, key=lambda u: self.shift_counts[u]):
                if self.is_user_available(user, date) and user not in daily_assignments[date]:
                    schedule[date][tier][shift] = user
                    daily_assignments[date].add(user)
                    self.shift_counts[user] += 1
                    self.fallback_assignments.append({
                        'date': date.strftime('%Y-%m-%d'),
                        'tier': tier,
                        'shift': shift,
                        'user': user,
                        'reason': 'cross-tier coverage'
                    })
                    return
        
        # Last resort: allow double shift for already assigned user
        if already_assigned:
            schedule[date][tier][shift] = f"{already_assigned} (DOUBLE)"
            self.shift_counts[already_assigned] += 1
            self.fallback_assignments.append({
                'date': date.strftime('%Y-%m-%d'),
                'tier': tier,
                'shift': shift,
                'user': already_assigned,
                'reason': 'double shift - no other coverage available'
            })
    
    def attempt_emergency_coverage(self, tier: str, date: datetime.date, 
                                  users: List[str], schedule: Dict, daily_assignments: Dict):
        """Attempt emergency coverage when no regular users are available"""
        # Special rule: Only upgrade users can cover upgrade shifts
        if tier == 'upgrade':
            all_users = set(self.upgrade_users)
        else:
            # For tier2/tier3, can use any tier2/tier3 users but not upgrade
            all_users = set(self.tier2_users + self.tier3_users)
        
        # Sort by shift count to maintain fairness even in emergency
        sorted_users = sorted(all_users, key=lambda u: self.shift_counts[u])
        
        shifts_to_fill = ['morning', 'evening'] if tier != 'upgrade' else ['full']
        
        for shift in shifts_to_fill:
            assigned = False
            for user in sorted_users:
                # Skip if already assigned today (unless no other option)
                if user not in daily_assignments[date]:
                    schedule[date][tier][shift] = f"{user} (EMERGENCY)"
                    daily_assignments[date].add(user)
                    self.shift_counts[user] += 1
                    self.fallback_assignments.append({
                        'date': date.strftime('%Y-%m-%d'),
                        'tier': tier,
                        'shift': shift,
                        'user': user,
                        'reason': 'emergency coverage - normal assignment failed'
                    })
                    assigned = True
                    break
            
            if not assigned and sorted_users:
                # Absolute last resort: assign to user with least shifts even if already assigned
                user = sorted_users[0]
                schedule[date][tier][shift] = f"{user} (EMERGENCY-DOUBLE)"
                self.shift_counts[user] += 1
                self.fallback_assignments.append({
                    'date': date.strftime('%Y-%m-%d'),
                    'tier': tier,
                    'shift': shift,
                    'user': user,
                    'reason': 'emergency double coverage - critical shortage'
                })
    
    def print_fairness_report(self, year=None, month=None):
        """Print a report showing shift distribution and coverage issues"""
        
        # Print PTO Summary first
        if year and month and self.pto_dates:
            print("\n=== PTO Summary ===")
            all_users = set(self.tier2_users + self.tier3_users + self.upgrade_users)
            users_with_pto = []
            excluded_from_fairness = []
            
            for user in sorted(all_users):
                if user in self.pto_dates:
                    pto_count = sum(1 for date in self.pto_dates[user] 
                                  if date.year == year and date.month == month)
                    if pto_count > 0:
                        users_with_pto.append((user, pto_count))
                        # Track users excluded from fairness metrics (>2 PTO days)
                        if pto_count > 2:
                            excluded_from_fairness.append(user)
            
            if users_with_pto:
                print(f"Users with PTO in {month}/{year}:")
                for user, days in sorted(users_with_pto, key=lambda x: x[1], reverse=True):
                    fairness_note = " [Excluded from fairness metrics]" if user in excluded_from_fairness else ""
                    print(f"  • {user}: {days} days{fairness_note}")
            else:
                print("No users have PTO this month")
            
            if excluded_from_fairness:
                print(f"\n📊 Users with >2 PTO days are excluded from fairness comparisons")
            print("-" * 40)
        
        # Print coverage warnings if any
        if self.coverage_warnings:
            print("\n⚠️  COVERAGE WARNINGS:")
            for warning in self.coverage_warnings:
                print(f"  - {warning}")
        
        if self.fallback_assignments:
            print("\n🔄 FALLBACK ASSIGNMENTS:")
            for fb in self.fallback_assignments:
                print(f"  - {fb['date']}: {fb['tier']} {fb['shift']} -> {fb['user']} ({fb['reason']})")
        
        print("\n=== Shift Distribution Report ===")
        print(f"{'User':<20} {'Month':<10} {'Cumulative':<12} {'Weekly':<10} {'Type'}")
        print("-" * 60)
        
        # Separate by user type
        tier2_counts = {u: self.shift_counts[u] for u in self.tier2_users if u in self.shift_counts}
        tier3_counts = {u: self.shift_counts[u] for u in self.tier3_users if u in self.shift_counts}
        upgrade_counts = {u: self.shift_counts[u] for u in self.upgrade_users if u in self.shift_counts}
        
        # Print Tier 2
        if self.tier2_users:
            print("\nTier 2 Users (daily shifts only):")
            # Include ALL tier2 users, even those with 0 shifts
            all_tier2 = {u: self.shift_counts.get(u, 0) for u in self.tier2_users}
            for user, count in sorted(all_tier2.items(), key=lambda x: x[1], reverse=True):
                # Count PTO days for this specific month if year/month provided
                if year and month and user in self.pto_dates:
                    pto_days = sum(1 for date in self.pto_dates[user] 
                                 if date.year == year and date.month == month)
                    if pto_days > 2:
                        pto_note = f" (PTO: {pto_days} days) *"
                    elif pto_days > 0:
                        pto_note = f" (PTO: {pto_days} days)"
                    else:
                        pto_note = ""
                else:
                    pto_note = ""
                weekly_count = self.monthly_weekly_assignments.get(user, 0)
                cumulative = self.cumulative_shift_counts.get(user, 0) + count
                print(f"{user:<20} {count:<10} {cumulative:<12} {weekly_count:<10}{pto_note}")
            if all_tier2:
                avg = sum(all_tier2.values()) / len(all_tier2)
                # Calculate cumulative average
                cumulative_avg = sum(self.cumulative_shift_counts.get(u, 0) + self.shift_counts.get(u, 0) 
                                    for u in self.tier2_users) / len(self.tier2_users) if self.tier2_users else 0
                print(f"Average: Month={avg:.1f}, Cumulative={cumulative_avg:.1f}")
        
        # Print Tier 3
        if self.tier3_users:
            print("\nTier 3 Users (weekly shifts):")
            # Include ALL tier3 users, even those with 0 shifts
            all_tier3 = {u: self.shift_counts.get(u, 0) for u in self.tier3_users}
            for user, count in sorted(all_tier3.items(), key=lambda x: x[1], reverse=True):
                # Count PTO days for this specific month if year/month provided
                if year and month and user in self.pto_dates:
                    pto_days = sum(1 for date in self.pto_dates[user] 
                                 if date.year == year and date.month == month)
                    if pto_days > 2:
                        pto_note = f" (PTO: {pto_days} days) *"
                    elif pto_days > 0:
                        pto_note = f" (PTO: {pto_days} days)"
                    else:
                        pto_note = ""
                else:
                    pto_note = ""
                weekly_count = self.monthly_weekly_assignments.get(user, 0)
                cumulative = self.cumulative_shift_counts.get(user, 0) + count
                print(f"{user:<20} {count:<10} {cumulative:<12} {weekly_count:<10}{pto_note}")
            if all_tier3:
                avg = sum(all_tier3.values()) / len(all_tier3)
                # Calculate cumulative average
                cumulative_avg = sum(self.cumulative_shift_counts.get(u, 0) + self.shift_counts.get(u, 0) 
                                    for u in self.tier3_users) / len(self.tier3_users) if self.tier3_users else 0
                print(f"Average: Month={avg:.1f}, Cumulative={cumulative_avg:.1f}")
        
        # Print Upgrade
        if self.upgrade_users:
            print("\nUpgrade Users (weekly shifts):")
            # Include ALL upgrade users, even those with 0 shifts
            all_upgrade = {u: self.shift_counts.get(u, 0) for u in self.upgrade_users}
            for user, count in sorted(all_upgrade.items(), key=lambda x: x[1], reverse=True):
                # Count PTO days for this specific month if year/month provided
                if year and month and user in self.pto_dates:
                    pto_days = sum(1 for date in self.pto_dates[user] 
                                 if date.year == year and date.month == month)
                    if pto_days > 2:
                        pto_note = f" (PTO: {pto_days} days) *"
                    elif pto_days > 0:
                        pto_note = f" (PTO: {pto_days} days)"
                    else:
                        pto_note = ""
                else:
                    pto_note = ""
                weekly_count = self.monthly_weekly_assignments.get(user, 0)
                cumulative = self.cumulative_shift_counts.get(user, 0) + count
                print(f"{user:<20} {count:<10} {cumulative:<12} {weekly_count:<10}{pto_note}")
            if all_upgrade:
                avg = sum(all_upgrade.values()) / len(all_upgrade)
                # Calculate cumulative average
                cumulative_avg = sum(self.cumulative_shift_counts.get(u, 0) + self.shift_counts.get(u, 0) 
                                    for u in self.upgrade_users) / len(self.upgrade_users) if self.upgrade_users else 0
                print(f"Average: Month={avg:.1f}, Cumulative={cumulative_avg:.1f}")
        
        print("=" * 40)
        
        # Add legend if there are users with >2 PTO days
        if year and month:
            all_users = set(self.tier2_users + self.tier3_users + self.upgrade_users)
            has_excluded = any(
                sum(1 for date in self.pto_dates.get(user, []) 
                    if date.year == year and date.month == month) > 2
                for user in all_users
            )
            if has_excluded:
                print("\n* = User has >2 PTO days and is excluded from fairness comparisons")

# Global scheduler instance
scheduler = OnCallScheduler()

# HTML template with embedded CSS and JavaScript
HTML_TEMPLATE = '''<!DOCTYPE html>
<html>
<head>
    <title>SRE On-Call Schedule Generator</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 30px;
        }
        .section {
            margin-bottom: 30px;
            padding: 20px;
            background-color: #f9f9f9;
            border-radius: 8px;
        }
        .user-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        .user-item {
            display: flex;
            align-items: center;
            padding: 10px;
            background-color: white;
            border: 1px solid #e0e0e0;
            border-radius: 5px;
        }
        .user-item input[type="checkbox"] {
            margin-right: 10px;
        }
        .file-input {
            margin-bottom: 15px;
        }
        .file-input input[type="file"] {
            padding: 8px;
            margin-left: 10px;
        }
        .file-input span {
            margin-left: 15px;
            color: #28a745;
            font-weight: bold;
        }
        .prior-month-section {
            background-color: #fff3cd;
            border: 1px solid #ffc107;
            margin-bottom: 20px;
        }
        .prior-month-info {
            padding: 10px;
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 4px;
            margin-top: 10px;
            color: #155724;
        }
        .controls {
            text-align: center;
            margin: 30px 0;
        }
        .controls input {
            padding: 10px;
            font-size: 16px;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-right: 10px;
        }
        .generate-btn {
            padding: 12px 30px;
            font-size: 16px;
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .generate-btn:hover {
            background-color: #218838;
        }
        .import-btn {
            padding: 10px 20px;
            font-size: 14px;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 10px;
        }
        .import-btn:hover {
            background-color: #0056b3;
        }
        .calendar {
            margin-top: 30px;
        }
        .calendar-header {
            text-align: center;
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #333;
        }
        .week {
            margin-bottom: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
        }
        .week-header {
            background-color: #f0f0f0;
            padding: 10px;
            font-weight: bold;
        }
        .days-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 1px;
            background-color: #ddd;
        }
        .day {
            background-color: white;
            padding: 10px;
            min-height: 120px;
        }
        .day-header {
            font-weight: bold;
            margin-bottom: 8px;
            padding-bottom: 5px;
            border-bottom: 1px solid #eee;
        }
        .shift {
            font-size: 12px;
            margin-bottom: 4px;
            padding: 3px 6px;
            border-radius: 3px;
        }
        .tier2 {
            background-color: #e3f2fd;
            color: #1565c0;
        }
        .tier3 {
            background-color: #f3e5f5;
            color: #6a1b9a;
        }
        .upgrade {
            background-color: #e8f5e9;
            color: #2e7d32;
        }
        .export-btn {
            position: fixed;
            bottom: 20px;
            right: 20px;
            padding: 12px 24px;
            background-color: #17a2b8;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }
        .export-btn:hover {
            background-color: #138496;
        }
        .loading {
            text-align: center;
            padding: 40px;
            font-size: 18px;
            color: #666;
        }
        .error {
            color: #dc3545;
            padding: 10px;
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 4px;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>SRE On-Call Schedule Generator</h1>
        
        <div class="section prior-month-section">
            <h2>Import Prior Month's Schedule (Optional)</h2>
            <p>To maintain rotation continuity and prevent back-to-back weekly assignments across months, 
               import the prior month's Excel export file:</p>
            <div class="file-input">
                <label for="prior-month-file">Prior Month Excel:</label>
                <input type="file" id="prior-month-file" accept=".xlsx" onchange="importPriorMonth(this)">
                <button class="import-btn" onclick="clearPriorMonth()">Clear Prior Month</button>
            </div>
            <div id="prior-month-info"></div>
        </div>
        
        <div class="section">
            <h2>Load Users</h2>
            <div class="file-input">
                <label for="tier2-file" style="display: inline-block; width: 150px;">Tier 2 Users:</label>
                <input type="file" id="tier2-file" accept=".txt" onchange="loadUsersFromFile('tier2', this)">
                <span id="tier2-count"></span>
            </div>
            <div class="file-input">
                <label for="tier3-file" style="display: inline-block; width: 150px;">Tier 3 Users:</label>
                <input type="file" id="tier3-file" accept=".txt" onchange="loadUsersFromFile('tier3', this)">
                <span id="tier3-count"></span>
            </div>
            <div class="file-input">
                <label for="upgrade-file" style="display: inline-block; width: 150px;">Upgrade Users:</label>
                <input type="file" id="upgrade-file" accept=".txt" onchange="loadUsersFromFile('upgrade', this)">
                <span id="upgrade-count"></span>
            </div>
        </div>
        
        <div class="section" id="pto-section" style="display: none;">
            <h2>Select PTO Dates</h2>
            <p>Check the boxes next to users who will be on PTO, then select their dates:</p>
            <div id="pto-users" class="user-grid"></div>
            <div id="pto-dates" style="margin-top: 20px;"></div>
        </div>
        
        <div class="controls">
            <input type="text" id="month-year" placeholder="MM/YYYY" value="">
            <button class="generate-btn" onclick="generateSchedule()">Generate Schedule</button>
        </div>
        
        <div id="calendar" class="calendar"></div>
        
        <button class="export-btn" onclick="exportSchedule()" style="display: none;">Export to Excel</button>
    </div>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script>
        let allUsers = new Set();
        let ptoUsers = new Set();
        let currentSchedule = null;
        let priorMonthData = null;
        
        // Set default month to current month
        document.getElementById('month-year').value = 
            new Date().toLocaleDateString('en-US', {month: '2-digit', year: 'numeric'}).replace('/', '/');
        
        function importPriorMonth(fileInput) {
            const file = fileInput.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    
                    // Get the main schedule sheet
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    // Convert to our schedule format
                    const schedule = {};
                    jsonData.forEach(row => {
                        const date = row['Date'];
                        const tier = row['Schedule'];
                        const shift = row['Shift'];
                        const user = row['User'];
                        
                        if (!schedule[date]) {
                            schedule[date] = {};
                        }
                        
                        if (tier === 'upgrade') {
                            if (!schedule[date].upgrade) schedule[date].upgrade = {};
                            schedule[date].upgrade.full = user;
                        } else if (tier === 'tier2' || tier === 'tier3') {
                            if (!schedule[date][tier]) schedule[date][tier] = {};
                            if (shift === 'morning') {
                                schedule[date][tier].morning = user;
                            } else if (shift === 'evening') {
                                schedule[date][tier].evening = user;
                            }
                        }
                    });
                    
                    // Send to server
                    fetch('/import_prior_month', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({schedule: schedule})
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.success) {
                            priorMonthData = data.last_week_assignments;
                            displayPriorMonthInfo(data);
                            
                            // Display shift counts if available
                            if (data.shift_counts && Object.keys(data.shift_counts).length > 0) {
                                let countsHtml = '<div style="margin-top: 10px; padding: 10px; background: #f0f0f0; border-radius: 5px;">';
                                countsHtml += '<strong>Cumulative Shift Counts Loaded:</strong><br>';
                                const sortedUsers = Object.entries(data.shift_counts).sort((a, b) => b[1] - a[1]);
                                sortedUsers.slice(0, 10).forEach(([user, count]) => {
                                    countsHtml += `${user}: ${count} shifts<br>`;
                                });
                                if (sortedUsers.length > 10) {
                                    countsHtml += `<em>...and ${sortedUsers.length - 10} more users</em><br>`;
                                }
                                countsHtml += '</div>';
                                
                                // Add to prior month display area
                                const priorMonthSection = document.getElementById('prior-month-section');
                                const existingCountsDiv = document.getElementById('shift-counts-display');
                                if (existingCountsDiv) {
                                    existingCountsDiv.innerHTML = countsHtml;
                                } else {
                                    const countsDiv = document.createElement('div');
                                    countsDiv.id = 'shift-counts-display';
                                    countsDiv.innerHTML = countsHtml;
                                    priorMonthSection.appendChild(countsDiv);
                                }
                            }
                        } else {
                            alert('Error importing prior month: ' + data.message);
                        }
                    });
                    
                } catch (error) {
                    console.error('Error parsing Excel file:', error);
                    alert('Error reading Excel file. Please ensure it is a valid schedule export.');
                }
            };
            
            reader.readAsArrayBuffer(file);
        }
        
        function clearPriorMonth() {
            fetch('/clear_prior_month', {
                method: 'POST'
            })
            .then(response => response.json())
            .then(data => {
                priorMonthData = null;
                document.getElementById('prior-month-info').innerHTML = '';
                document.getElementById('prior-month-file').value = '';
            });
        }
        
        function displayPriorMonthInfo(data) {
            const infoDiv = document.getElementById('prior-month-info');
            const assignments = data.last_week_assignments;
            
            let html = '<div class="prior-month-info">';
            html += '<strong>' + data.message + '</strong><br>';
            html += '<br><strong>Last Week Assignments:</strong><br>';
            html += '• Upgrade: ' + (assignments.upgrade || 'None') + '<br>';
            html += '• Tier3 Morning: ' + (assignments.tier3_morning || 'None') + '<br>';
            html += '• Tier3 Evening: ' + (assignments.tier3_evening || 'None') + '<br>';
            html += '<br><em>These users will not be scheduled for the first week of the new month.</em>';
            html += '</div>';
            
            infoDiv.innerHTML = html;
        }
        
        function loadUsersFromFile(tier, fileInput) {
            const file = fileInput.files[0];
            if (!file) {
                console.error('No file selected');
                return;
            }
            
            console.log(`Loading ${file.name} for ${tier}`);
            
            const reader = new FileReader();
            reader.onload = function(e) {
                const content = e.target.result;
                const users = content.split('\\n')
                    .map(line => line.trim())
                    .filter(line => line.length > 0);
                
                console.log(`Found ${users.length} users:`, users);
                
                // Send users to server
                fetch('/load_users_direct', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({tier: tier, users: users})
                })
                .then(response => response.json())
                .then(data => {
                    if (data.error) {
                        alert(data.error);
                        console.error('Error:', data.error);
                    } else {
                        document.getElementById(`${tier}-count`).textContent = 
                            `✔ ${data.count} users loaded`;
                        console.log(`Successfully loaded ${data.count} users for ${tier}`);
                        updatePTOSection();
                    }
                })
                .catch(error => {
                    console.error('Fetch error:', error);
                    alert('Error loading file: ' + error.message);
                });
            };
            
            reader.onerror = function(error) {
                console.error('FileReader error:', error);
                alert('Error reading file');
            };
            
            reader.readAsText(file);
        }
        
        function updatePTOSection() {
            fetch('/get_all_users')
            .then(response => response.json())
            .then(data => {
                allUsers = new Set(data.users);
                if (allUsers.size > 0) {
                    document.getElementById('pto-section').style.display = 'block';
                    renderPTOUsers();
                }
            });
        }
        
        function renderPTOUsers() {
            const container = document.getElementById('pto-users');
            container.innerHTML = '';
            
            Array.from(allUsers).sort().forEach(user => {
                const div = document.createElement('div');
                div.className = 'user-item';
                div.innerHTML = `
                    <input type="checkbox" id="pto-${user}" onchange="togglePTO('${user}')">
                    <label for="pto-${user}">${user}</label>
                `;
                container.appendChild(div);
            });
        }
        
        function togglePTO(user) {
            if (document.getElementById(`pto-${user}`).checked) {
                ptoUsers.add(user);
            } else {
                ptoUsers.delete(user);
                const dateInput = document.getElementById(`dates-${user}`);
                if (dateInput) dateInput.remove();
            }
            updatePTODates();
        }
        
        function updatePTODates() {
            const container = document.getElementById('pto-dates');
            container.innerHTML = '';
            
            if (ptoUsers.size === 0) return;
            
            container.innerHTML = '<h3>Enter PTO dates:</h3>';
            container.innerHTML += '<p style="font-size: 14px; color: #666;">Format: MM/DD/YYYY or DD/MM/YYYY (auto-detected)<br>';
            container.innerHTML += 'Examples: 03/15/2024-03/20/2024 or 15/03/2024-20/03/2024<br>';
            container.innerHTML += 'Multiple ranges: Use commas to separate (e.g., 03/01/2024-03/05/2024, 03/15/2024-03/20/2024)</p>';
            
            Array.from(ptoUsers).sort().forEach(user => {
                const div = document.createElement('div');
                div.id = `dates-${user}`;
                div.style.marginBottom = '15px';
                div.innerHTML = `
                    <label style="display: inline-block; width: 100px; vertical-align: top;">${user}:</label>
                    <textarea id="pto-range-${user}" 
                             placeholder="01/03/2024-05/03/2024, 15/03/2024-20/03/2024" 
                             style="width: 400px; height: 60px; vertical-align: top;"></textarea>
                `;
                container.appendChild(div);
            });
        }
        
        function generateSchedule() {
            const monthYear = document.getElementById('month-year').value;
            if (!monthYear || !monthYear.match(/^\\d{2}\\/\\d{4}$/)) {
                alert('Please enter month in MM/YYYY format');
                return;
            }
            
            // Collect PTO data
            const ptoData = {};
            ptoUsers.forEach(user => {
                const ranges = document.getElementById(`pto-range-${user}`).value;
                if (ranges) {
                    ptoData[user] = ranges;
                }
            });
            
            document.getElementById('calendar').innerHTML = '<div class="loading">Generating schedule...</div>';
            
            fetch('/generate', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    month_year: monthYear,
                    pto: ptoData
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    document.getElementById('calendar').innerHTML = 
                        `<div class="error">${data.error}</div>`;
                } else {
                    currentSchedule = data;
                    renderCalendar(data);
                    document.querySelector('.export-btn').style.display = 'block';
                }
            });
        }
        
        function renderCalendar(data) {
            const container = document.getElementById('calendar');
            container.innerHTML = `<div class="calendar-header">${data.month_name} ${data.year}</div>`;
            
            // Display validation errors first if any critical ones exist
            if (data.validation_errors && data.validation_errors.critical && data.validation_errors.critical.length > 0) {
                const criticalDiv = document.createElement('div');
                criticalDiv.style.cssText = 'background: #ffcccc; border: 3px solid #cc0000; padding: 15px; margin: 10px 0; border-radius: 5px;';
                criticalDiv.innerHTML = '<h3 style="color: #cc0000; margin-top: 0;">❌ CRITICAL VALIDATION ERRORS - Requirements Violated!</h3>';
                criticalDiv.innerHTML += '<p style="color: #cc0000; font-weight: bold;">The following requirements were NOT met:</p>';
                data.validation_errors.critical.forEach(error => {
                    criticalDiv.innerHTML += `<div style="margin: 5px 0; color: #cc0000;">• ${error}</div>`;
                });
                container.appendChild(criticalDiv);
            }
            
            // Display validation warnings
            if (data.validation_errors && data.validation_errors.warnings && data.validation_errors.warnings.length > 0) {
                const valWarningsDiv = document.createElement('div');
                valWarningsDiv.style.cssText = 'background: #fff8dc; border: 2px solid #ff9800; padding: 15px; margin: 10px 0; border-radius: 5px;';
                valWarningsDiv.innerHTML = '<h3 style="color: #e65100; margin-top: 0;">⚠️ Validation Warnings</h3>';
                data.validation_errors.warnings.forEach(warning => {
                    valWarningsDiv.innerHTML += `<div style="margin: 5px 0;">• ${warning}</div>`;
                });
                container.appendChild(valWarningsDiv);
            }
            
            // Display coverage warnings if any
            if (data.coverage_warnings && data.coverage_warnings.length > 0) {
                const warningsDiv = document.createElement('div');
                warningsDiv.style.cssText = 'background: #ffebee; border: 2px solid #f44336; padding: 15px; margin: 10px 0; border-radius: 5px;';
                warningsDiv.innerHTML = '<h3 style="color: #d32f2f; margin-top: 0;">⚠️ Coverage Issues Detected</h3>';
                data.coverage_warnings.forEach(warning => {
                    warningsDiv.innerHTML += `<div style="margin: 5px 0;">• ${warning}</div>`;
                });
                container.appendChild(warningsDiv);
            }
            
            // Display fallback assignments if any
            if (data.fallback_assignments && data.fallback_assignments.length > 0) {
                const fallbackDiv = document.createElement('div');
                fallbackDiv.style.cssText = 'background: #fff3cd; border: 2px solid #ffc107; padding: 15px; margin: 10px 0; border-radius: 5px;';
                fallbackDiv.innerHTML = '<h3 style="color: #856404; margin-top: 0;">🔄 Fallback Assignments (Double Shifts & Emergency Coverage)</h3>';
                
                // Group by date for better readability
                const byDate = {};
                data.fallback_assignments.forEach(fb => {
                    if (!byDate[fb.date]) byDate[fb.date] = [];
                    byDate[fb.date].push(fb);
                });
                
                Object.keys(byDate).sort().forEach(date => {
                    fallbackDiv.innerHTML += `<div style="margin: 10px 0;"><strong>${date}:</strong>`;
                    byDate[date].forEach(fb => {
                        const isDouble = fb.reason.includes('double');
                        const color = isDouble ? '#d32f2f' : '#856404';
                        fallbackDiv.innerHTML += `<div style="margin-left: 20px; color: ${color};">• ${fb.tier} ${fb.shift}: ${fb.user} - <em>${fb.reason}</em></div>`;
                    });
                    fallbackDiv.innerHTML += '</div>';
                });
                container.appendChild(fallbackDiv);
            }
            
            data.weeks.forEach((week, idx) => {
                const weekDiv = document.createElement('div');
                weekDiv.className = 'week';
                
                const weekHeader = document.createElement('div');
                weekHeader.className = 'week-header';
                weekHeader.textContent = `Week ${idx + 1}: ${week.start} - ${week.end}`;
                weekDiv.appendChild(weekHeader);
                
                const daysGrid = document.createElement('div');
                daysGrid.className = 'days-grid';
                
                week.days.forEach(day => {
                    const dayDiv = document.createElement('div');
                    dayDiv.className = 'day';
                    
                    const dayHeader = document.createElement('div');
                    dayHeader.className = 'day-header';
                    dayHeader.textContent = day.date;
                    dayDiv.appendChild(dayHeader);
                    
                    // Add shifts with special styling for double/emergency shifts
                    if (day.shifts.tier2_morning) {
                        const shift = document.createElement('div');
                        const isSpecial = day.shifts.tier2_morning.includes('DOUBLE') || day.shifts.tier2_morning.includes('EMERGENCY');
                        shift.className = 'shift tier2';
                        if (isSpecial) {
                            shift.style.cssText = 'background: #ffcccc !important; border: 2px solid #ff0000 !important; font-weight: bold;';
                        }
                        shift.textContent = `T2 AM: ${day.shifts.tier2_morning}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier2_evening) {
                        const shift = document.createElement('div');
                        const isSpecial = day.shifts.tier2_evening.includes('DOUBLE') || day.shifts.tier2_evening.includes('EMERGENCY');
                        shift.className = 'shift tier2';
                        if (isSpecial) {
                            shift.style.cssText = 'background: #ffcccc !important; border: 2px solid #ff0000 !important; font-weight: bold;';
                        }
                        shift.textContent = `T2 PM: ${day.shifts.tier2_evening}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier3_morning) {
                        const shift = document.createElement('div');
                        const isSpecial = day.shifts.tier3_morning.includes('DOUBLE') || day.shifts.tier3_morning.includes('EMERGENCY');
                        shift.className = 'shift tier3';
                        if (isSpecial) {
                            shift.style.cssText = 'background: #ffcccc !important; border: 2px solid #ff0000 !important; font-weight: bold;';
                        }
                        shift.textContent = `T3 AM: ${day.shifts.tier3_morning}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier3_evening) {
                        const shift = document.createElement('div');
                        const isSpecial = day.shifts.tier3_evening.includes('DOUBLE') || day.shifts.tier3_evening.includes('EMERGENCY');
                        shift.className = 'shift tier3';
                        if (isSpecial) {
                            shift.style.cssText = 'background: #ffcccc !important; border: 2px solid #ff0000 !important; font-weight: bold;';
                        }
                        shift.textContent = `T3 PM: ${day.shifts.tier3_evening}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.upgrade) {
                        const shift = document.createElement('div');
                        const isSpecial = day.shifts.upgrade.includes('DOUBLE') || day.shifts.upgrade.includes('EMERGENCY');
                        shift.className = 'shift upgrade';
                        if (isSpecial) {
                            shift.style.cssText = 'background: #ffcccc !important; border: 2px solid #ff0000 !important; font-weight: bold;';
                        }
                        shift.textContent = `Upgrade: ${day.shifts.upgrade}`;
                        dayDiv.appendChild(shift);
                    }
                    
                    daysGrid.appendChild(dayDiv);
                });
                
                weekDiv.appendChild(daysGrid);
                container.appendChild(weekDiv);
            });
        }
        
        function exportSchedule() {
            if (!currentSchedule) return;
            
            fetch('/export', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(currentSchedule)
            })
            .then(response => response.blob())
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `oncall_schedule_${currentSchedule.month}_${currentSchedule.year}.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            });
        }
    </script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/import_prior_month', methods=['POST'])
def import_prior_month():
    data = request.json
    schedule_data = data.get('schedule', {})
    
    result = scheduler.import_prior_month_schedule(schedule_data)
    return jsonify(result)

@app.route('/clear_prior_month', methods=['POST'])
def clear_prior_month():
    scheduler.prior_month_last_week = {
        'upgrade': None,
        'tier3_morning': None,
        'tier3_evening': None,
        'date': None
    }
    scheduler.last_upgrade_user = None
    scheduler.last_tier3_morning_user = None
    scheduler.last_tier3_evening_user = None
    
    return jsonify({'success': True})

@app.route('/load_users_direct', methods=['POST'])
def load_users_direct():
    data = request.json
    tier = data['tier']
    users = data['users']
    
    if not users:
        return jsonify({'error': 'No users provided'})
    
    if tier == 'tier2':
        scheduler.tier2_users = users
    elif tier == 'tier3':
        scheduler.tier3_users = users
    elif tier == 'upgrade':
        scheduler.upgrade_users = users
    
    return jsonify({'count': len(users)})

@app.route('/get_all_users')
def get_all_users():
    all_users = list(set(scheduler.tier2_users + scheduler.tier3_users + scheduler.upgrade_users))
    return jsonify({'users': all_users})

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    month_year = data['month_year']
    pto_data = data.get('pto', {})
    
    try:
        month, year = month_year.split('/')
        month = int(month)
        year = int(year)
    except:
        return jsonify({'error': 'Invalid month/year format'})
    
    # Clear previous PTO data
    scheduler.pto_dates.clear()
    
    # Parse PTO dates - now handles multiple ranges per user
    parsing_errors = []
    for user, date_ranges_str in pto_data.items():
        if date_ranges_str:
            dates = set()
            # Split by comma to handle multiple ranges
            ranges = [r.strip() for r in date_ranges_str.split(',')]
            
            for date_range in ranges:
                parsed = False
                if '-' in date_range:
                    start_str, end_str = date_range.strip().split('-')
                    start_str, end_str = start_str.strip(), end_str.strip()
                    
                    # Try MM/DD/YYYY format first
                    try:
                        start = datetime.datetime.strptime(start_str, '%m/%d/%Y').date()
                        end = datetime.datetime.strptime(end_str, '%m/%d/%Y').date()
                        parsed = True
                    except ValueError:
                        # Try DD/MM/YYYY format as fallback
                        try:
                            start = datetime.datetime.strptime(start_str, '%d/%m/%Y').date()
                            end = datetime.datetime.strptime(end_str, '%d/%m/%Y').date()
                            parsed = True
                            print(f"Note: {user}'s date {date_range} was parsed as DD/MM/YYYY format")
                        except ValueError as e:
                            parsing_errors.append(f"{user}: Could not parse date range '{date_range}'")
                            continue
                    
                    if parsed:
                        current = start
                        while current <= end:
                            dates.add(current)
                            current += datetime.timedelta(days=1)
                else:
                    # Single date (no range)
                    single_date_str = date_range.strip()
                    
                    # Try MM/DD/YYYY format first
                    try:
                        single_date = datetime.datetime.strptime(single_date_str, '%m/%d/%Y').date()
                        dates.add(single_date)
                    except ValueError:
                        # Try DD/MM/YYYY format as fallback
                        try:
                            single_date = datetime.datetime.strptime(single_date_str, '%d/%m/%Y').date()
                            dates.add(single_date)
                            print(f"Note: {user}'s date {single_date_str} was parsed as DD/MM/YYYY format")
                        except ValueError:
                            parsing_errors.append(f"{user}: Could not parse date '{single_date_str}'")
            
            if dates:
                scheduler.pto_dates[user] = dates
                print(f"Successfully parsed {len(dates)} PTO dates for {user}")
    
    # Report any parsing errors
    if parsing_errors:
        print("\n⚠️  PTO Parsing Errors:")
        for error in parsing_errors:
            print(f"  - {error}")
    
    # Clear previous warnings and fallback assignments
    scheduler.coverage_warnings = []
    scheduler.fallback_assignments = []
    
    # Generate schedule
    schedule = scheduler.generate_schedule(year, month)
    
    # Format for display
    import calendar
    weeks_data = []
    weeks = scheduler.get_month_weeks(year, month)
    
    for week_start, week_end in weeks:
        week_data = {
            'start': week_start.strftime('%b %d'),
            'end': week_end.strftime('%b %d'),
            'days': []
        }
        
        current = week_start
        while current <= week_end:
            day_shifts = {}
            
            if current in schedule:
                if 'tier2' in schedule[current]:
                    if 'morning' in schedule[current]['tier2']:
                        day_shifts['tier2_morning'] = schedule[current]['tier2']['morning']
                    if 'evening' in schedule[current]['tier2']:
                        day_shifts['tier2_evening'] = schedule[current]['tier2']['evening']
                
                if 'tier3' in schedule[current]:
                    if 'morning' in schedule[current]['tier3']:
                        day_shifts['tier3_morning'] = schedule[current]['tier3']['morning']
                    if 'evening' in schedule[current]['tier3']:
                        day_shifts['tier3_evening'] = schedule[current]['tier3']['evening']
                
                if 'upgrade' in schedule[current] and 'full' in schedule[current]['upgrade']:
                    day_shifts['upgrade'] = schedule[current]['upgrade']['full']
            
            week_data['days'].append({
                'date': current.strftime('%a %d'),
                'full_date': current.strftime('%Y-%m-%d'),
                'shifts': day_shifts
            })
            
            current += datetime.timedelta(days=1)
        
        weeks_data.append(week_data)
    
    # Run validation
    validation_errors = scheduler.validate_schedule(schedule, year, month)
    
    return jsonify({
        'month': month,
        'year': year,
        'month_name': calendar.month_name[month],
        'weeks': weeks_data,
        'schedule': {k.strftime('%Y-%m-%d'): v for k, v in schedule.items()},
        'coverage_warnings': scheduler.coverage_warnings,
        'fallback_assignments': scheduler.fallback_assignments,
        'validation_errors': validation_errors
    })

@app.route('/export', methods=['POST'])
def export():
    data = request.json
    schedule_data = data['schedule']
    
    # Create an Excel file with multiple sheets
    import io
    import xlsxwriter
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Create formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D3D3D3',
        'border': 1
    })
    
    tier2_format = workbook.add_format({'bg_color': '#E3F2FD'})
    tier3_format = workbook.add_format({'bg_color': '#F3E5F5'})
    upgrade_format = workbook.add_format({'bg_color': '#E8F5E9'})
    
    # Main schedule sheet
    main_sheet = workbook.add_worksheet('Full Schedule')
    headers = ['Date', 'Day', 'Schedule', 'Shift', 'Time', 'User']
    for col, header in enumerate(headers):
        main_sheet.write(0, col, header, header_format)
    
    # Collect user assignments for individual sheets
    user_assignments = defaultdict(list)
    
    row = 1
    for date_str, day_schedule in sorted(schedule_data.items()):
        date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
        day_name = date.strftime('%A')
        
        for tier in ['tier2', 'tier3', 'upgrade']:
            if tier in day_schedule:
                if tier == 'upgrade':
                    if 'full' in day_schedule[tier]:
                        user = day_schedule[tier]['full']
                        time_range = '12:00pm-8:30pm EST'
                        shift_type = 'full'
                        
                        # Write to main sheet
                        main_sheet.write(row, 0, date_str)
                        main_sheet.write(row, 1, day_name)
                        main_sheet.write(row, 2, tier, upgrade_format)
                        main_sheet.write(row, 3, shift_type)
                        main_sheet.write(row, 4, time_range)
                        main_sheet.write(row, 5, user)
                        row += 1
                        
                        # Collect for user sheet
                        user_assignments[user].append({
                            'date': date_str,
                            'day': day_name,
                            'schedule': tier,
                            'shift': shift_type,
                            'time': time_range
                        })
                else:
                    if 'morning' in day_schedule[tier]:
                        user = day_schedule[tier]['morning']
                        time_range = '11:00am-5:00pm EST'
                        shift_type = 'morning'
                        
                        # Write to main sheet
                        format_to_use = tier2_format if tier == 'tier2' else tier3_format
                        main_sheet.write(row, 0, date_str)
                        main_sheet.write(row, 1, day_name)
                        main_sheet.write(row, 2, tier, format_to_use)
                        main_sheet.write(row, 3, shift_type)
                        main_sheet.write(row, 4, time_range)
                        main_sheet.write(row, 5, user)
                        row += 1
                        
                        # Collect for user sheet
                        user_assignments[user].append({
                            'date': date_str,
                            'day': day_name,
                            'schedule': tier,
                            'shift': shift_type,
                            'time': time_range
                        })
                        
                    if 'evening' in day_schedule[tier]:
                        user = day_schedule[tier]['evening']
                        time_range = '5:00pm-11:00pm EST'
                        shift_type = 'evening'
                        
                        # Write to main sheet
                        format_to_use = tier2_format if tier == 'tier2' else tier3_format
                        main_sheet.write(row, 0, date_str)
                        main_sheet.write(row, 1, day_name)
                        main_sheet.write(row, 2, tier, format_to_use)
                        main_sheet.write(row, 3, shift_type)
                        main_sheet.write(row, 4, time_range)
                        main_sheet.write(row, 5, user)
                        row += 1
                        
                        # Collect for user sheet
                        user_assignments[user].append({
                            'date': date_str,
                            'day': day_name,
                            'schedule': tier,
                            'shift': shift_type,
                            'time': time_range
                        })
    
    # Autofit columns on main sheet
    main_sheet.set_column(0, 0, 12)  # Date
    main_sheet.set_column(1, 1, 10)  # Day
    main_sheet.set_column(2, 2, 10)  # Schedule
    main_sheet.set_column(3, 3, 10)  # Shift
    main_sheet.set_column(4, 4, 20)  # Time
    main_sheet.set_column(5, 5, 15)  # User
    
    # Create individual user sheets
    for user in sorted(user_assignments.keys()):
        # Clean sheet name (Excel has restrictions)
        sheet_name = user[:31]  # Excel max sheet name is 31 chars
        user_sheet = workbook.add_worksheet(sheet_name)
        
        # Headers
        user_headers = ['Date', 'Day', 'Schedule', 'Shift', 'Time']
        for col, header in enumerate(user_headers):
            user_sheet.write(0, col, header, header_format)
        
        # User's assignments
        assignments = sorted(user_assignments[user], key=lambda x: x['date'])
        for row_num, assignment in enumerate(assignments, 1):
            # Determine format based on schedule type
            if assignment['schedule'] == 'tier2':
                row_format = tier2_format
            elif assignment['schedule'] == 'tier3':
                row_format = tier3_format
            else:
                row_format = upgrade_format
            
            user_sheet.write(row_num, 0, assignment['date'])
            user_sheet.write(row_num, 1, assignment['day'])
            user_sheet.write(row_num, 2, assignment['schedule'], row_format)
            user_sheet.write(row_num, 3, assignment['shift'])
            user_sheet.write(row_num, 4, assignment['time'])
        
        # Add summary at the top
        user_sheet.write(row_num + 2, 0, 'Total Shifts:', header_format)
        user_sheet.write(row_num + 2, 1, len(assignments))
        
        # Autofit columns
        user_sheet.set_column(0, 0, 12)  # Date
        user_sheet.set_column(1, 1, 10)  # Day
        user_sheet.set_column(2, 2, 10)  # Schedule
        user_sheet.set_column(3, 3, 10)  # Shift
        user_sheet.set_column(4, 4, 20)  # Time
    
    workbook.close()
    output.seek(0)
    
    return output.getvalue(), 200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': f'attachment; filename=oncall_schedule_{data["month"]}_{data["year"]}.xlsx'
    }

def open_browser():
    time.sleep(1)
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == "__main__":
    print("=== SRE On-Call Schedule Generator ===")
    print("\nStarting web interface...")
    print("Opening browser to http://127.0.0.1:5000")
    print("\nPress Ctrl+C to stop the server\n")
    
    # Open browser automatically
    threading.Thread(target=open_browser).start()
    
    # Run Flask app
    app.run(debug=False, port=5000)
