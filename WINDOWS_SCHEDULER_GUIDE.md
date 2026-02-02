# Windows Task Scheduler Setup Guide

## Complete Step-by-Step Instructions

### Prerequisites

✅ Python 3.7+ installed
✅ Required packages installed (`pandas`, `openpyxl`)
✅ Database created (`sales_data.db`)
✅ Config file updated (`config.ini`)

---

## Method 1: Using Task Scheduler GUI

### Step 1: Open Task Scheduler

1. Press **Windows Key + R**
2. Type: `taskschd.msc`
3. Press **Enter**

### Step 2: Create Basic Task

1. In the right panel, click **"Create Basic Task..."**
2. Enter Task Details:
   - **Name**: `Daily Sales Report`
   - **Description**: `Automated Excel report generation and email delivery`
3. Click **"Next"**

### Step 3: Set Trigger (When to Run)

1. Select **"Daily"**
2. Click **"Next"**
3. Configure schedule:
   - **Start date**: Select today's date
   - **Start time**: `08:00:00 AM` (or your preferred time)
   - **Recur every**: `1` days
4. Click **"Next"**

### Step 4: Set Action (What to Run)

1. Select **"Start a program"**
2. Click **"Next"**
3. Configure program:
   - **Program/script**: Browse to your Python executable
     - Example: `C:\Python311\python.exe`
     - To find your Python path, open Command Prompt and type: `where python`
   - **Add arguments (optional)**: `generate_report.py`
   - **Start in (optional)**: `C:\path\to\app\` (your project directory)
     - Example: `C:\Users\YourName\Documents\app\`

**Example Configuration**:
```
Program/script: C:\Python311\python.exe
Add arguments: generate_report.py
Start in: C:\Users\John\Documents\app\
```

4. Click **"Next"**

### Step 5: Review and Finish

1. Review all settings
2. Check **"Open the Properties dialog for this task when I click Finish"**
3. Click **"Finish"**

### Step 6: Advanced Settings (Properties Dialog)

1. **General Tab**:
   - Check: **"Run whether user is logged on or not"**
   - Check: **"Run with highest privileges"** (if needed)
   - Configure for: **Windows 10** (or your OS version)

2. **Triggers Tab**:
   - Verify your daily schedule is correct
   - Optional: Click **"Edit"** to add:
     - **"Repeat task every"**: 6 hours (for more frequent reports)
     - **"Stop task if it runs longer than"**: 30 minutes

3. **Actions Tab**:
   - Verify the Python command is correct
   - Double-check the "Start in" path

4. **Conditions Tab**:
   - ❌ **Uncheck**: "Start the task only if the computer is on AC power"
   - ❌ **Uncheck**: "Stop if the computer switches to battery power"
   - ✅ **Check**: "Wake the computer to run this task" (if you want it to run even when sleeping)

5. **Settings Tab**:
   - ✅ **Check**: "Allow task to be run on demand"
   - ✅ **Check**: "Run task as soon as possible after a scheduled start is missed"
   - **If the task fails, restart every**: `5 minutes`, **Attempt to restart up to**: `3` times
   - ✅ **Check**: "If the task is already running, do not start a new instance"

6. Click **"OK"**

7. If prompted, enter your **Windows password** to save the task

---

## Method 2: Using Command Line (PowerShell)

### Quick Setup Command

Open **PowerShell as Administrator** and run:

```powershell
$action = New-ScheduledTaskAction -Execute "C:\Python311\python.exe" -Argument "generate_report.py" -WorkingDirectory "C:\path\to\app"
$trigger = New-ScheduledTaskTrigger -Daily -At "8:00AM"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName "Daily Sales Report" -Action $action -Trigger $trigger -Settings $settings -Description "Automated Excel report generation"
```

**Remember to replace**:
- `C:\Python311\python.exe` with your Python path
- `C:\path\to\app` with your project directory

---

## Testing Your Scheduled Task

### Test Run (Before Scheduling)

1. Open **Task Scheduler**
2. Find your task: **"Daily Sales Report"**
3. Right-click → **"Run"**
4. Check the **"Last Run Result"** column:
   - `0x0` or `The operation completed successfully` = ✅ Success
   - Other codes = ❌ Error

### View Task History

1. Right-click on your task
2. Select **"Properties"**
3. Go to **"History"** tab
4. Review execution logs

### Check Output Files

1. Navigate to your project directory: `C:\path\to\app\`
2. Look for new Excel files: `sales_report_YYYYMMDD_HHMMSS.xlsx`
3. Open the file to verify:
   - Summary sheet with data
   - Data sheet with all records
   - Two charts visible

---

## Troubleshooting Common Issues

### Issue 1: Task Shows "Success" but No File Generated

**Cause**: Python script has errors

**Solution**:
1. Edit your task
2. In **Actions** tab, modify the command to:
   - **Program/script**: `cmd.exe`
   - **Add arguments**: `/c "C:\Python311\python.exe generate_report.py > output.log 2>&1"`
   - **Start in**: Your project directory
3. This will create `output.log` file with error messages

### Issue 2: "The system cannot find the file specified"

**Cause**: Incorrect paths

**Solution**:
1. Verify Python path: Open Command Prompt, type `where python`
2. Verify project path: Use full absolute paths (no relative paths)
3. Ensure "Start in" directory is set correctly

### Issue 3: "Access Denied" or Permission Errors

**Cause**: Insufficient permissions

**Solution**:
1. Right-click task → **Properties**
2. **General** tab → Check **"Run with highest privileges"**
3. Ensure your Windows user account has permissions to:
   - Read the database file
   - Write Excel files to the directory
   - Access Gmail (if firewall/antivirus blocking)

### Issue 4: Email Not Sending

**Cause**: Incorrect Gmail App Password or firewall

**Solution**:
1. Verify Gmail App Password in `config.ini`
2. Test manually: `python generate_report.py`
3. Check firewall settings (allow Python to access network)
4. Verify 2-Step Verification is enabled on Gmail

### Issue 5: Task Runs But Stops Immediately

**Cause**: Python environment issues

**Solution**:
1. Create a batch file wrapper:

**create_report.bat**:
```batch
@echo off
cd /d C:\path\to\app
C:\Python311\python.exe generate_report.py
echo Completed at %date% %time% >> task_log.txt
```

2. Point Task Scheduler to this batch file instead
3. Program/script: `C:\path\to\app\create_report.bat`

---

## Schedule Examples

### Daily at 8 AM
```
Trigger: Daily
Start time: 8:00 AM
Recur every: 1 days
```

### Every Weekday at 9 AM
```
Trigger: Weekly
Recur every: 1 weeks
Days: Monday, Tuesday, Wednesday, Thursday, Friday
Start time: 9:00 AM
```

### Multiple Times Per Day (Every 6 Hours)
```
Trigger: Daily
Start time: 6:00 AM
Advanced settings:
  - Repeat task every: 6 hours
  - For a duration of: 1 day
```

### First Day of Every Month
```
Trigger: Monthly
Months: All
Days: 1
Start time: 8:00 AM
```

---

## Monitoring and Maintenance

### Weekly Checklist

- [ ] Check Task Scheduler history for failed runs
- [ ] Verify Excel reports are being generated
- [ ] Confirm emails are being received
- [ ] Review error logs (if any)
- [ ] Check disk space for report storage

### Monthly Checklist

- [ ] Archive old reports (keep last 90 days)
- [ ] Update sample data if needed
- [ ] Review and optimize SQL queries
- [ ] Test Gmail App Password (passwords can expire)
- [ ] Update Python packages: `pip install --upgrade pandas openpyxl`

---

## Disabling or Deleting the Task

### Temporarily Disable

1. Open Task Scheduler
2. Find your task: **"Daily Sales Report"**
3. Right-click → **"Disable"**

### Permanently Delete

1. Open Task Scheduler
2. Find your task: **"Daily Sales Report"**
3. Right-click → **"Delete"**
4. Confirm deletion

---

## Advanced: Running as Windows Service

For production environments, consider running as a Windows Service:

1. Use **NSSM (Non-Sucking Service Manager)**
2. Download from: https://nssm.cc/download
3. Install service:
   ```
   nssm install SalesReportService "C:\Python311\python.exe" "C:\path\to\app\generate_report.py"
   ```

---

## Command Reference

### Find Python Path
```cmd
where python
```

### Test Script Manually
```cmd
cd C:\path\to\app
python generate_report.py
```

### List All Scheduled Tasks
```cmd
schtasks /query /fo LIST /v
```

### Delete Task via Command Line
```cmd
schtasks /delete /tn "Daily Sales Report" /f
```

---

## Summary

✅ Task Scheduler automates your report generation
✅ No need to run manually every day
✅ Reports delivered to email automatically
✅ Set it and forget it!

**Your automation is now complete!** 🎉
