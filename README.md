# SRE On-Call Schedule Generator

A web-based tool for generating fair and constraint-aware on-call schedules for Site Reliability Engineering (SRE) teams. Features a visual calendar interface, automatic rotation management, and PTO handling.

## Features

- **Three Schedule Types**:
  - **Tier 2**: Two daily shifts (11am-5pm EST and 5pm-11pm EST)
  - **Tier 3**: Two weekly shifts (11am-5pm EST and 5pm-11pm EST, Monday-Sunday)
  - **Upgrade**: One weekly shift (12pm-8:30pm EST, Monday-Sunday)

- **Smart Scheduling Constraints**:
  - No user has overlapping shifts across schedules
  - No user works more than 1 shift per day
  - Weekly shifts (Tier 3 and Upgrade) have fair rotation with no back-to-back weeks
  - All weekly shift users must get a turn before anyone repeats

- **User-Friendly Web Interface**:
  - File browser for easy user list upload
  - Checkbox-based PTO selection
  - Visual calendar display with color-coded shifts
  - CSV export functionality

- **PTO Management**:
  - Support for multiple PTO date ranges per user
  - Visual interface for entering time off
  - Automatic conflict resolution

## Prerequisites

- Python 3.6 or higher
- Flask web framework
- macOS (or any OS with Python support)

## Installation

1. **Clone or download the script**:
   ```bash
   wget sre_oncall_scheduler.py
   # or save the script to a file named sre_oncall_scheduler.py
   ```

2. **Install Flask**:
   ```bash
   pip install flask
   ```

3. **Make the script executable**:
   ```bash
   chmod +x sre_oncall_scheduler.py
   ```

## User File Format

Create text files for each tier with one username per line:

**tier2_users.txt**:
```
alice
bob
charlie
david
```

**tier3_users.txt**:
```
eve
frank
grace
henry
```

**upgrade_users.txt**:
```
ivan
julia
kevin
lisa
```

## Usage

1. **Start the application**:
   ```bash
   ./sre_oncall_scheduler.py
   ```
   The web interface will automatically open in your default browser at `http://localhost:5000`

2. **Load user files**:
   - Click "Choose File" next to each tier
   - Browse and select your text files
   - You'll see a green checkmark with the number of users loaded

3. **Configure PTO (Optional)**:
   - Check the boxes next to users who need time off
   - Enter date ranges in the format: `DD/MM/YYYY-DD/MM/YYYY`
   - Multiple ranges can be entered separated by commas:
     ```
     01/03/2024-05/03/2024, 15/03/2024-20/03/2024
     ```

4. **Generate Schedule**:
   - Enter the month in `MM/YYYY` format (e.g., `03/2024`)
   - Click "Generate Schedule"
   - The system will show which weeks are included in the schedule

5. **Export Results**:
   - Click "Export to CSV" to download the schedule
   - The CSV includes all shift assignments with dates and times

## Schedule Logic

### Daily Shifts (Tier 2 Only)
- Morning shift: 11am-5pm EST
- Evening shift: 5pm-11pm EST
- Different users are assigned to morning and evening shifts each day
- Random assignment among available users

### Weekly Shifts (Tier 3 & Upgrade)
**Tier 3 Shifts**:
- Morning shift: 11am-5pm EST (Monday-Sunday)
- Evening shift: 5pm-11pm EST (Monday-Sunday)
- **Fair Rotation Rules**:
  1. No back-to-back weeks for the same user
  2. All users must get one assignment before anyone repeats
  3. Separate rotation queues for morning and evening shifts
  4. Only breaks these rules if PTO makes it impossible

**Upgrade Shift**:
- Full shift: 12pm-8:30pm EST (Monday-Sunday)
- Same fair rotation rules as Tier 3

### Conflict Resolution
- The system automatically handles conflicts when users have:
  - PTO dates
  - Existing assignments on other tiers
  - Back-to-back weekly shifts

## Visual Interface

The web interface provides:
- **Color-coded shifts**:
  - Blue: Tier 2 assignments
  - Purple: Tier 3 assignments  
  - Green: Upgrade assignments
- **Week-by-week view** with all 7 days displayed
- **Clear labeling**: T2 AM/PM, T3 AM/PM, Upgrade

## CSV Export Format

The exported CSV includes:
- Date
- Day of week
- Schedule tier (tier2/tier3/upgrade)
- Shift type (morning/evening/full)
- Time range
- Assigned user

Example:
```csv
Date,Day,Schedule,Shift,Time,User
2024-03-01,Friday,tier2,morning,11:00am-5:00pm EST,alice
2024-03-01,Friday,tier2,evening,5:00pm-11:00pm EST,bob
2024-03-01,Friday,upgrade,full,12:00pm-8:30pm EST,ivan
```

## Troubleshooting

### No users appear after file selection
- Ensure the text files contain one username per line
- Check the browser console (F12) for error messages
- Verify the Flask server is running

### Schedule generation fails
- Verify the month format is `MM/YYYY`
- Check that enough users are available (not all on PTO)
- Ensure no typos in PTO date ranges

### Back-to-back warnings
- This occurs when PTO constraints make fair rotation impossible for weekly shifts
- The system will display warnings in the server console
- Warnings specify which shift type (tier3_morning, tier3_evening, or upgrade)
- Review PTO assignments if this happens frequently

## Server Output

The Flask server will show:
- Startup confirmation
- File loading confirmations
- Warnings about scheduling conflicts
- Any errors that occur

Keep the terminal window open to monitor these messages.

## Stopping the Server

Press `Ctrl+C` in the terminal to stop the Flask server.

## Notes

- The scheduler handles months that span multiple weeks (e.g., if the 1st isn't a Monday)
- All times are in EST
- Tier 2 uses daily rotation with random selection among available candidates
- Tier 3 and Upgrade use weekly rotation with fair distribution
- PTO takes precedence over fair rotation rules
- Weekly shifts run Monday through Sunday
