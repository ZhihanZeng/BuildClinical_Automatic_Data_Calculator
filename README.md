# BuildClinical Weekly Tracker

The BuildClinical Weekly Tracker is a Python program that helps research coordinators monitor clinical study recruitment. It reads BuildClinical export files, compares each lead's current status against the previous week's snapshot, and emails a weekly summary of recruitment progress. Because changes are tracked by participant ID rather than by submission date, the tool catches leads that were submitted weeks ago but only changed status this week.

It reports weekly changes across the recruitment funnel, including:
- Attempted Contact
- Contacted
- Screened
- Consented
- Enrolled
- Ineligible (with the listed reason for each lead)
- New Leads (current total and change since last week)

## Features
- **Excel/CSV Parsing:** Reads `.xlsx` and `.csv` exports from BuildClinical.
- **ID-Based Change Tracking:** Compares every lead's status to its last saved status, so no status change is missed regardless of when the lead was submitted.
- **Funnel Progress Reporting:** Separately counts leads that moved out of New Lead into each downstream stage.
- **Ineligibility Details:** Lists each newly ineligible lead along with its recorded reason.
- **Separate High/Low Risk Tracking:** Keeps an independent snapshot for each study population so their comparisons never mix.
- **Safe State Saving:** Snapshots are written atomically, so a crash mid-save can't corrupt last week's data.
- **Automated Email Reports:** Sends the weekly summary directly from the app via Gmail.
- **Graphical User Interface:** Simple file picker and email form built with tkinter.

## Technologies
- Python 3.x
- pandas (data loading and manipulation)
- openpyxl (Excel support)
- tkinter (GUI)
- smtplib / email (report delivery)
- json (weekly snapshot storage)

## Usage
1. Run the program: `python BuildClinical.py`
2. Select **High Risk** or **Low Risk**.
3. Choose your BuildClinical export file with the file picker.
4. The program compares the file against last week's snapshot and calculates the weekly changes.
5. Enter your Gmail address, app password, and the recipient's email, then click **Send Email**.

> Note: Gmail requires an [App Password](https://support.google.com/accounts/answer/185833) rather than your regular account password.

## Example Output
    Contacted: 4
    Attempted Contact: 7
    Screened: 3
    Consented: 2
    Enrolled: 1
    Ineligible: 2
    New Leads: 18
    Change since last week: -5
    Ineligibility Details:
    ID: 1042 Reason for Ineligibility: Outside age range
    ID: 1057 Reason for Ineligibility: Declined participation

## Data Privacy
The snapshot files (`state_dataset_a.json`, `state_dataset_b.json`) store participant IDs and statuses. They are listed in `.gitignore` and should never be committed. Do not commit real export files either.

## Future Improvements
- Only save the weekly snapshot after a report is successfully sent, so reloading a file doesn't erase the week's changes
- Store the date of each snapshot so the report period reflects when the tool was actually last run
- Display results directly in the GUI
- Error handling for failed email logins and missing files
- Export weekly reports to CSV or Excel for long-term tracking
