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
  - Prioritizes users with fewer total shifts for tier 2 assignments
  
- **PTO Management**: 
  - Support for multiple PTO date ranges per user
  - Automatically excludes users from scheduling during their time off
  
- **Web-Based Interface**: 
  - Clean, intuitive UI for managing users and generating schedules
  - Visual calendar display with color-coded shifts
  - No installation required for end users - just open in browser
  
- **Excel Export**: 
  - Comprehensive Excel workbook with multiple sheets
  - Main schedule sheet with all assignments
  - Individual sheets for each user showing their specific shifts
  - Color-coded by tier for easy reading

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

## Installation

1. Clone or download the script
2. Install required dependencies:
```bash
pip install flask xlsxwriter
```

## Usage

### Starting the Application

Run the script from command line:
```bash
python oncall_scheduler.py
```

The application will:
1. Start a local web server on port 5000
2. Automatically open your default browser to `http://localhost:5000`
3. Display the web interface for schedule generation

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
3. Enter PTO date ranges in `DD/MM/YYYY-DD/MM/YYYY` format
4. Multiple date ranges can be entered separated by commas

Example:
```
01/03/2024-05/03/2024, 15/03/2024-20/03/2024
```

### Generating the Schedule

1. Enter the target month in `MM/YYYY` format (e.g., `03/2024`)
2. Click "Generate Schedule"
3. The schedule will display in a visual calendar format
4. Review the assignments for accuracy and fairness

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

### Daily Assignments (Tier 2)
- Assigned on a per-day basis
- Algorithm prioritizes users with fewer total shifts
- Ensures balanced distribution across the month

### Conflict Prevention
- No user can have multiple shifts on the same day
- PTO dates are strictly respected
- System warns if constraints cannot be met

## Fairness Report

After generation, the console displays a fairness report showing:
- Total shifts per user
- Average shifts by tier
- Distribution analysis

This helps verify equitable scheduling and identify any imbalances.

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

### Excel Export Not Working
Ensure xlsxwriter is installed:
```bash
pip install --upgrade xlsxwriter
```

### Schedule Generation Errors
- Verify sufficient users are loaded for each tier
- Check that PTO dates don't conflict with minimum staffing requirements
- Review console output for specific error messages

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
- `POST /generate` - Generates the schedule
- `POST /export` - Creates and downloads Excel file

## License

This script is provided as-is for use in SRE team scheduling. Feel free to modify and adapt to your organization's needs.

## Contributing

Suggestions and improvements are welcome. Key areas for enhancement:
- Integration with calendar systems (Google Calendar, Outlook)
- Slack/email notifications
- Historical schedule tracking
- Preference-based scheduling
- On-call swap management

## Support

For issues or questions:
1. Check the console output for detailed error messages
2. Verify all dependencies are installed correctly
3. Ensure input files are properly formatted
4. Review the fairness report for scheduling insights
