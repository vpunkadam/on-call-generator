#!/usr/bin/env python3
"""
SRE On-Call Schedule Generator with Web UI
Generates on-call schedules for tier2, tier3, and upgrade shifts
"""

import datetime
import random
from collections import defaultdict
from typing import List, Dict, Tuple, Set
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
        self.shift_counts = defaultdict(int)  # Track total shifts per user
        
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
    
    def generate_schedule(self, year: int, month: int):
        """Generate the on-call schedule for a given month"""
        schedule = defaultdict(lambda: defaultdict(dict))
        
        # Track daily assignments to prevent overlaps
        daily_assignments = defaultdict(set)  # date -> set of users
        
        # Reset shift counts for this generation
        self.shift_counts = defaultdict(int)
        
        # Initialize rotation queues
        self.upgrade_rotation_queue = self.upgrade_users.copy()
        random.shuffle(self.upgrade_rotation_queue)
        
        self.tier3_morning_rotation_queue = self.tier3_users.copy()
        random.shuffle(self.tier3_morning_rotation_queue)
        
        self.tier3_evening_rotation_queue = self.tier3_users.copy()
        random.shuffle(self.tier3_evening_rotation_queue)
        
        # Get all weeks for the month
        weeks = self.get_month_weeks(year, month)
        
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
            
            # Assign tier2 daily shifts
            current = week_start
            while current <= week_end:
                self.assign_daily_shifts_with_fairness('tier2', current, self.tier2_users, 
                                       schedule, daily_assignments)
                current += datetime.timedelta(days=1)
        
        # Print fairness report
        self.print_fairness_report()
        
        return schedule
    
    def assign_weekly_shift_with_rotation(self, shift_type: str, week_start: datetime.date, 
                                         daily_assignments: Dict, rotation_queue: List[str],
                                         last_user: str, all_users: List[str]) -> str:
        """Assign weekly shift with fair rotation - no back-to-back weeks"""
        # First, try to find someone from the rotation queue who:
        # 1. Is available for the entire week
        # 2. Wasn't scheduled last week (no back-to-back)
        # 3. Hasn't been scheduled yet in this rotation
        
        available_from_queue = []
        for user in rotation_queue:
            # Skip if this user was scheduled last week
            if user == last_user:
                continue
                
            # Check if user is available for entire week
            available_all_week = True
            for day in range(7):
                date = week_start + datetime.timedelta(days=day)
                if not self.is_user_available(user, date) or user in daily_assignments[date]:
                    available_all_week = False
                    break
            
            if available_all_week:
                available_from_queue.append(user)
        
        # If we found someone from the queue, use them
        if available_from_queue:
            selected_user = available_from_queue[0]
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
        all_available = []
        for user in all_users:
            # Skip if this user was scheduled last week
            if user == last_user:
                continue
                
            available_all_week = True
            for day in range(7):
                date = week_start + datetime.timedelta(days=day)
                if not self.is_user_available(user, date) or user in daily_assignments[date]:
                    available_all_week = False
                    break
            
            if available_all_week:
                all_available.append(user)
        
        if all_available:
            selected_user = random.choice(all_available)
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
        # If no users available, leave empty
    
    def assign_daily_shifts_with_fairness(self, tier: str, date: datetime.date, 
                                          users: List[str], schedule: Dict, 
                                          daily_assignments: Dict):
        """Assign morning and evening shifts for a tier with fairness consideration"""
        # Get available users for this date
        available_users = [u for u in users 
                         if self.is_user_available(u, date) 
                         and u not in daily_assignments[date]]
        
        if len(available_users) >= 2:
            # Sort by shift count (ascending) to prioritize users with fewer shifts
            available_users.sort(key=lambda u: self.shift_counts[u])
            
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
        # If no users available, leave empty
    
    def print_fairness_report(self):
        """Print a report showing shift distribution"""
        if not self.shift_counts:
            return
        
        print("\n=== Shift Distribution Report ===")
        print(f"{'User':<20} {'Shifts':<10} {'Type'}")
        print("-" * 40)
        
        # Separate by user type
        tier2_counts = {u: self.shift_counts[u] for u in self.tier2_users if u in self.shift_counts}
        tier3_counts = {u: self.shift_counts[u] for u in self.tier3_users if u in self.shift_counts}
        upgrade_counts = {u: self.shift_counts[u] for u in self.upgrade_users if u in self.shift_counts}
        
        # Print Tier 2
        if tier2_counts:
            print("\nTier 2 Users:")
            for user, count in sorted(tier2_counts.items(), key=lambda x: x[1], reverse=True):
                print(f"{user:<20} {count:<10}")
            avg = sum(tier2_counts.values()) / len(tier2_counts)
            print(f"Average: {avg:.1f} shifts")
        
        # Print Tier 3
        if tier3_counts:
            print("\nTier 3 Users:")
            for user, count in sorted(tier3_counts.items(), key=lambda x: x[1], reverse=True):
                print(f"{user:<20} {count:<10}")
            avg = sum(tier3_counts.values()) / len(tier3_counts)
            print(f"Average: {avg:.1f} shifts")
        
        # Print Upgrade
        if upgrade_counts:
            print("\nUpgrade Users:")
            for user, count in sorted(upgrade_counts.items(), key=lambda x: x[1], reverse=True):
                print(f"{user:<20} {count:<10} (weekly)")
            avg = sum(upgrade_counts.values()) / len(upgrade_counts)
            print(f"Average: {avg:.1f} shifts")
        
        print("=" * 40)

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
    
    <script>
        let allUsers = new Set();
        let ptoUsers = new Set();
        let currentSchedule = null;
        
        // Set default month to current month
        document.getElementById('month-year').value = 
            new Date().toLocaleDateString('en-US', {month: '2-digit', year: 'numeric'}).replace('/', '/');
        
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
                            `✓ ${data.count} users loaded`;
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
        
        function loadUsers(tier) {
            const filename = document.getElementById(`${tier}-file`).value;
            if (!filename) {
                alert('Please enter a filename');
                return;
            }
            
            fetch('/load_users', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({tier: tier, filename: filename})
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    alert(data.error);
                } else {
                    document.getElementById(`${tier}-count`).textContent = 
                        `(${data.count} users loaded)`;
                    updatePTOSection();
                }
            });
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
            
            container.innerHTML = '<h3>Enter PTO dates (DD/MM/YYYY-DD/MM/YYYY):</h3>';
            container.innerHTML += '<p style="font-size: 14px; color: #666;">You can enter multiple date ranges separated by commas (e.g., 01/03/2024-05/03/2024, 15/03/2024-20/03/2024)</p>';
            
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
                    
                    // Add shifts
                    if (day.shifts.tier2_morning) {
                        const shift = document.createElement('div');
                        shift.className = 'shift tier2';
                        shift.textContent = `T2 AM: ${day.shifts.tier2_morning}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier2_evening) {
                        const shift = document.createElement('div');
                        shift.className = 'shift tier2';
                        shift.textContent = `T2 PM: ${day.shifts.tier2_evening}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier3_morning) {
                        const shift = document.createElement('div');
                        shift.className = 'shift tier3';
                        shift.textContent = `T3 AM: ${day.shifts.tier3_morning}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.tier3_evening) {
                        const shift = document.createElement('div');
                        shift.className = 'shift tier3';
                        shift.textContent = `T3 PM: ${day.shifts.tier3_evening}`;
                        dayDiv.appendChild(shift);
                    }
                    if (day.shifts.upgrade) {
                        const shift = document.createElement('div');
                        shift.className = 'shift upgrade';
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

@app.route('/load_users', methods=['POST'])
def load_users():
    data = request.json
    tier = data['tier']
    filename = data['filename']
    
    users = scheduler.load_users_from_file(filename)
    if not users:
        return jsonify({'error': f'Could not load users from {filename}'})
    
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
    for user, date_ranges_str in pto_data.items():
        if date_ranges_str:
            dates = set()
            # Split by comma to handle multiple ranges
            ranges = [r.strip() for r in date_ranges_str.split(',')]
            
            for date_range in ranges:
                try:
                    if '-' in date_range:
                        start_str, end_str = date_range.strip().split('-')
                        start = datetime.datetime.strptime(start_str.strip(), '%d/%m/%Y').date()
                        end = datetime.datetime.strptime(end_str.strip(), '%d/%m/%Y').date()
                        
                        current = start
                        while current <= end:
                            dates.add(current)
                            current += datetime.timedelta(days=1)
                except:
                    # Skip invalid date ranges
                    continue
            
            if dates:
                scheduler.pto_dates[user] = dates
    
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
    
    return jsonify({
        'month': month,
        'year': year,
        'month_name': calendar.month_name[month],
        'weeks': weeks_data,
        'schedule': {k.strftime('%Y-%m-%d'): v for k, v in schedule.items()}
    })

@app.route('/export', methods=['POST'])
def export():
    data = request.json
    schedule_data = data['schedule']
    
    # Create an Excel file with multiple sheets instead of CSV
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
    webbrowser.open('http://localhost:5000')

if __name__ == "__main__":
    print("=== SRE On-Call Schedule Generator ===")
    print("\nStarting web interface...")
    print("Opening browser to http://localhost:5000")
    print("\nPress Ctrl+C to stop the server\n")
    
    # Open browser automatically
    threading.Thread(target=open_browser).start()
    
    # Run Flask app
    app.run(debug=False, port=5000)
