# SRE On-Call Schedule Generator

A Python-based web application for generating fair and balanced on-call schedules for Site Reliability Engineering (SRE) teams with multiple support tiers.

## Features

- **Multi-Tier Support**: Manages three different on-call tiers:
  - **Tier 2**: Daily shifts (morning and evening)
  - **Tier 3**: Weekly shifts (morning and evening)
  - **Upgrade**: Weekly full shifts
  
- **Fair Rotation Algorithm**: 
  - Ensures equitable distribution of shifts among team members
  - Prevents back-to-back weekly assignments
  - Prioritizes users with fewer total shifts for fairness
  - Enforces "everyone gets 1 before anyone gets 2" rule for weekly shifts
  - Maximum 2 weekly shifts per user per month
  - Persistent shift tracking across months for long-term fairness
  
- **PTO Management**: 
  - Support for multiple PTO date ranges per user
  - Automatically excludes users from scheduling during their time off
  - Intelligent date format detection (supports both MM/DD/YYYY and DD/MM/YYYY)
  - PTO is strictly enforced - no scheduling during time off
  - Users with >2 PTO days are excluded from fairness comparisons (no catch-up mechanism)
  
- **Fallback Coverage Strategy**:
  - Automatic fallback when insufficient users are available
  - Cross-tier coverage (except upgrade tier remains restricted)
  - Double shift assignments when necessary (clearly marked)
  - Emergency coverage as last resort
  
- **Comprehensive Validation**:
  - Post-generation validation ensures all requirements are met
  - Checks for PTO violations, double-bookings, and shift imbalances
  - Accounts for PTO when validating fairness
  - Warnings for any constraint violations
  
- **Web-Based Interface**: 
  - Clean, intuitive UI for managing users and generating schedules
  - Visual calendar display with color-coded shifts
  - Clear indicators for double shifts and emergency coverage
  - Real-time validation feedback
  
- **Excel Import/Export**: 
  - Import previous month's schedule to track cumulative shift counts
  - Comprehensive Excel workbook with multiple sheets
  - Main schedule sheet with all assignments
  - Individual sheets for each user showing their specific shifts
  - Color-coded by tier for easy reading
  - Preserves shift history for fairness tracking

## Shift Timings

- **Tier 2 & Tier 3**:
  - Morning: 11:00 AM - 5:00 PM EST
  - Evening: 5:00 PM - 11:00 PM EST
  
- **Upgrade**:
  - Full shift: 12:00 PM - 8:30 PM EST

## Requirements

- Python 3.6+
- Flask
- xlsxwriter
- openpyxl (for Excel import functionality)

## Installation

1. Clone or download the repository
2. Install required dependencies:
```bash
pip install flask xlsxwriter openpyxl
```

## Usage

### Starting the Application

Run the script from command line:
```bash
python sre_oncall_scheduler_ui.py
```

The application will:
1. Start a local web server on port 5000
2. Automatically open your default browser to `http://127.0.0.1:5000`
3. Display the web interface for schedule generation
4. Load any existing shift history from `shift_history.json` if present

### Loading Users

1. Create text files with user names (one per line) for each tier:
   - `tier2_users.txt`
   - `tier3_users.txt`
   - `upgrade_users.txt`

Example file format:
```
john.doe
jane.smith
bob.wilson
sarah.jones
```

2. In the web interface, use the file upload buttons to load users for each tier
3. The interface will display the count of loaded users for confirmation

### Setting PTO Dates

1. After loading users, the PTO section will appear
2. Check the box next to any user who will be on PTO
3. Enter PTO date ranges in `MM/DD/YYYY-MM/DD/YYYY` format (primary format)
4. Multiple date ranges can be entered separated by commas
5. Single dates can also be entered: `MM/DD/YYYY`
6. The system intelligently detects and handles both MM/DD/YYYY and DD/MM/YYYY formats

Examples:
```
03/01/2024-03/05/2024, 03/15/2024-03/20/2024
12/09/2025
09/06/2025-09/09/2025, 09/27/2025-10/13/2025
```

### Importing Previous Month's Schedule (Optional)

1. If you have a previous month's Excel schedule, click "Import Previous Month Schedule"
2. Upload the Excel file to carry forward shift counts
3. This ensures long-term fairness across months

### Generating the Schedule

1. Enter the target month in `MM/YYYY` format (e.g., `03/2024`)
2. Click "Generate Schedule"
3. The schedule will display in a visual calendar format
4. Review the assignments for accuracy and fairness
5. Check for any validation warnings (displayed in red)
6. Look for special indicators:
   - **(DOUBLE)**: User has multiple shifts that day
   - **(EMERGENCY)**: Emergency fallback coverage

### Exporting to Excel

1. After generating a schedule, click the "Export to Excel" button
2. The download will include:
   - Full schedule with all assignments
   - Individual sheets for each user
   - Color coding by tier
   - Shift counts and summaries

## Schedule Rules

### Weekly Assignments (Tier 3 & Upgrade)
- Users are assigned for entire weeks (Monday-Sunday)
- No back-to-back week assignments (when possible)
- Fair rotation ensures all users get equal opportunities
- **Priority Rule**: Everyone gets 1 weekly shift before anyone gets 2
- Maximum 2 weekly shifts per user per month
- Upgrade tier is restricted - only upgrade users can be assigned

### Daily Assignments (Tier 2)
- Assigned on a per-day basis
- Algorithm prioritizes users with fewer total shifts
- Ensures balanced distribution across the month
- Tracks cumulative shifts across months for long-term fairness

### Conflict Prevention
- No user can have multiple shifts on the same day
- PTO dates are strictly respected - no exceptions
- System warns if constraints cannot be met
- Automatic fallback coverage when necessary

### Fairness Philosophy
- Users with more than 2 PTO days are excluded from fairness comparisons
- No catch-up mechanism - users on extended PTO don't need to make up shifts
- Fairness is calculated only among actively available workers
- This prevents unfair burden on those taking legitimate time off

### Fallback Coverage Hierarchy
When insufficient users are available:
1. **Cross-tier coverage**: Users from other tiers cover (except upgrade)
2. **Double shifts**: Existing users take additional shifts (marked as "DOUBLE")
3. **Emergency coverage**: Last resort assignment (marked as "EMERGENCY")
4. Upgrade tier never uses fallback - remains restricted to upgrade users only

## Fairness Report

After generation, the console displays a comprehensive fairness report showing:
- Total shifts per user (including those with 0 shifts due to PTO)
- PTO days taken by each user
- Users with >2 PTO days marked as excluded from fairness metrics
- Average shifts by tier
- Distribution analysis (excludes users with >2 PTO days)
- Validation warnings for any issues detected
- Shift imbalance calculations (only compares users with ≤2 PTO days)

This helps verify equitable scheduling among available workers without penalizing those on extended leave.

## Troubleshooting

### Port Already in Use
If port 5000 is already in use, modify the last line of the script:
```python
app.run(debug=False, port=5001)  # Change to any available port
```

### File Upload Issues
- Ensure text files are properly formatted (one name per line)
- Remove any empty lines or special characters
- Use UTF-8 encoding for files with special characters

### Excel Import/Export Not Working
Ensure required libraries are installed:
```bash
pip install --upgrade xlsxwriter openpyxl
```

### Schedule Generation Errors
- Verify sufficient users are loaded for each tier
- Check that PTO dates don't conflict with minimum staffing requirements
- Review console output for specific error messages
- Check validation warnings for any constraint violations
- If users show 0 shifts, verify they don't have PTO for the entire month

### PTO Not Being Recognized
- Ensure date format is correct (MM/DD/YYYY preferred)
- Check for mixed date formats - system will try to auto-detect
- Verify no extra spaces or special characters in date ranges
- Single dates should not have a dash: `12/09/2025` not `12/09/2025-`

## Customization

### Modifying Shift Times
Edit the shift definitions in the `__init__` method:
```python
self.shifts = {
    'tier2': {
        'morning': {'start': '11:00', 'end': '17:00', 'timezone': 'EST'},
        'evening': {'start': '17:00', 'end': '23:00', 'timezone': 'EST'}
    },
    # ... modify as needed
}
```

### Changing Rotation Logic
The rotation algorithm can be customized in:
- `assign_weekly_shift_with_rotation()` - For weekly assignments
- `assign_daily_shifts_with_fairness()` - For daily assignments

### UI Customization
The HTML template can be modified directly in the `HTML_TEMPLATE` variable for styling or layout changes.

## Architecture

The application consists of:
- **OnCallScheduler Class**: Core scheduling logic and algorithms
- **Flask Web Server**: Handles HTTP requests and serves the UI
- **HTML/JavaScript Frontend**: Interactive user interface
- **RESTful API Endpoints**: Communication between frontend and backend

## API Endpoints

- `GET /` - Serves the main web interface
- `POST /load_users_direct` - Loads users from uploaded files
- `GET /get_all_users` - Returns all loaded users
- `POST /generate` - Generates the schedule with validation
- `POST /export` - Creates and downloads Excel file
- `POST /import_excel` - Imports previous month's Excel schedule

## License

This script is provided as-is for use in SRE team scheduling. Feel free to modify and adapt to your organization's needs.

## Data Persistence

The application maintains persistent data in the following files:
- `cumulative_shift_history.json` - Tracks total cumulative shift counts across all months
  - Automatically loaded at startup
  - Updated after each schedule generation
  - Used to prioritize users with fewer historical shifts
- `tier2_users.txt`, `tier3_users.txt`, `upgrade_users.txt` - User lists for each tier
- Generated Excel files serve as historical records

### Historical Fairness

The system maintains cumulative shift counts across months to ensure long-term fairness:
- Users with fewer total historical shifts are prioritized for new assignments
- Cumulative counts are shown in the fairness report alongside monthly counts
- This prevents any single user from being consistently overloaded across months
- Excel imports also update the cumulative history to maintain continuity

## Contributing

Suggestions and improvements are welcome. Key areas for potential enhancement:
- Integration with calendar systems (Google Calendar, Outlook)
- Slack/email notifications for schedule publication
- Preference-based scheduling (preferred/blocked dates)
- On-call swap management between users
- Dashboard for historical analytics

## Support

For issues or questions:
1. Check the console output for detailed error messages
2. Verify all dependencies are installed correctly
3. Ensure input files are properly formatted
4. Review the fairness report for scheduling insights
