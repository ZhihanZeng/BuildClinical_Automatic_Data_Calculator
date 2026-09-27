import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import os

STATE_FILE_A = "state_dataset_a.json"
STATE_FILE_B = "state_dataset_b.json"

root = tk.Tk()
root.title("BuildClinical")
root.geometry("400x300")

sender_email_var = tk.StringVar()
sender_email_pass = tk.StringVar()
reciever_email_var = tk.StringVar()
high_or_low_risk_var = tk.StringVar()

df = None

def select_file():
    global file_path
    global df
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel/CSV files", "*.xlsx *.csv"), ("All files", "*.*")]
    )
    if not file_path:
        print("No file selected.")
        return

    if not high_or_low_risk_var.get():
        messagebox.showwarning(
            "Missing selection",
            "Please select High Risk or Low Risk before loading a file."
        )
        return

    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    show_totals()

def get_state_file():
    """Pick the state file based on a StringVar's current value."""
    if high_or_low_risk_var.get() == "Low Risk":
        return STATE_FILE_A
    elif high_or_low_risk_var.get() == "High Risk":
        return STATE_FILE_B
    else:
        messagebox.showwarning(
            "Missing selection",
            "Please select High Risk or Low Risk."
        )
        return None

def load_previous_value(key, default=0):
    if default is None:
        default = {} if key == "id_status_map" else 0
    state_file = get_state_file()
    if state_file is None:
        return default
    if os.path.exists(state_file):
        try:
            with open(state_file, "r") as f:
                data = json.load(f)
                return data.get(key, default)
        except (json.JSONDecodeError, ValueError):
            return default
    return default

def save_current_value(key, value):
    state_file = get_state_file()
    if state_file is None:
        return

    data = {}
    if os.path.exists(state_file):
        try:
            with open(state_file, "r") as f:
                data = json.load(f)
        except (json.JSONDecodeError, ValueError):
            data = {}

    data[key] = value

    tmp_file = state_file + ".tmp"
    with open(tmp_file, "w") as f:
        json.dump(data, f)
    os.replace(tmp_file, state_file)


def compute_status_changes(current_id_status, previous_statuses):
    """
    ID-based (NOT submission-date-based) weekly change tracker.

    Buckets every ID whose status differs from what was saved for it
    last run, keyed by its NEW status - regardless of what it was
    before or when it was originally submitted. This is what catches
    a lead that was submitted a month ago and only just changed
    status this week.
    """
    buckets = {
        "attempted": [],
        "contacted": [],
        "ineligible": [],
        "screened": [],
        "consented": [],
        "enrolled": [],
    }

    for id_, current_status in current_id_status.items():
        previous_status = previous_statuses.get(id_, "")
        if current_status == previous_status:
            continue  # no change since the last saved snapshot

        cs_lower = current_status.lower()
        if current_status == "Attempted Contact":
            buckets["attempted"].append(id_)
        elif current_status == "Contacted":
            buckets["contacted"].append(id_)
        elif "ineligible" in cs_lower:
            buckets["ineligible"].append(id_)
        elif current_status == "Screened":
            buckets["screened"].append(id_)
        elif current_status == "Consented":
            buckets["consented"].append(id_)
        elif current_status == "Enrolled":
            buckets["enrolled"].append(id_)

    return buckets


def find_progressed_leads(current_id_status, previous_statuses):
    """
    Funnel-specific view: of the IDs that changed, which ones
    specifically moved OUT of New Lead/Eligible into a downstream
    status. current_id_status is passed in (built once in show_totals)
    rather than rebuilt from df here.
    """
    progressed_to_attempted = []
    progressed_to_contacted = []
    progressed_to_ineligible = []
    progressed_to_screened = []
    progressed_to_consented = []
    progressed_to_enrolled = []

    for id_, current_status in current_id_status.items():
        previous_status = previous_statuses.get(id_, "")
        was_new_lead = "new lead" in previous_status.lower() or "eligible" in previous_status.lower()

        if was_new_lead and current_status == "Attempted Contact":
            progressed_to_attempted.append(id_)
        elif was_new_lead and current_status == "Contacted":
            progressed_to_contacted.append(id_)
        elif was_new_lead and current_status == "Ineligible":
            progressed_to_ineligible.append(id_)
        elif was_new_lead and current_status == "Screened":
            progressed_to_screened.append(id_)
        elif was_new_lead and current_status == "Consented":
            progressed_to_consented.append(id_)
        elif was_new_lead and current_status == "Enrolled":
            progressed_to_enrolled.append(id_)

    return (
        len(progressed_to_attempted),
        len(progressed_to_contacted),
        len(progressed_to_ineligible),
        len(progressed_to_screened),
        len(progressed_to_consented),
        len(progressed_to_enrolled),
    )


def show_ineligibility_details(current_id_status, previous_statuses):
    details = []
    for id_, current_status in current_id_status.items():
        if "ineligible" not in current_status.lower():
            continue

        previous_status = previous_statuses.get(id_, "")
        was_new_lead = "new lead" in previous_status.lower()

        if not was_new_lead:
            continue

        row = df[df['ID'].astype(str) == id_].iloc[0]
        reason = row['Reason'] if pd.notna(row['Reason']) else "No Reason"
        details.append(f"ID: {id_} Reason for Ineligibility: {reason}")

    if not details:
        return "None in the past week."

    return "\n" + "\n".join(details)


def show_totals():
    global one_week_ago
    global today
    global contacted_total
    global attempted_contact_total
    global screened_total
    global consented_total
    global enrolled_total
    global ineligible_total
    global new_leads_total
    global difference
    global progressed_to_attempted
    global progressed_to_contacted
    global ineligibility_details_text

    # today/one_week_ago are kept only for labeling the email/report period -
    # they no longer filter which rows count, since we now track changes by
    # ID against the last saved snapshot instead of by Submission Date.
    today = pd.Timestamp.now().normalize()
    one_week_ago = today - pd.Timedelta(days=7)

    current_id_status = dict(zip(df['ID'].astype(str), df['Status'].astype(str)))

    # --- Load last week's snapshot ONE time, before anything gets overwritten ---
    previous_statuses = load_previous_value("id_status_map", {})
    previous_new_leads_total = load_previous_value("new_leads_total", 0)

    # --- ID-based totals: catches an ID submitted long ago that only
    # changed status this week, which date-filtering would have missed ---
    changes = compute_status_changes(current_id_status, previous_statuses)
    attempted_contact_total = len(changes["attempted"])
    contacted_total = len(changes["contacted"])
    screened_total = len(changes["screened"])
    consented_total = len(changes["consented"])
    enrolled_total = len(changes["enrolled"])
    ineligible_total = len(changes["ineligible"])

    # --- Funnel-specific view: only transitions that started at New Lead ---
    (progressed_to_attempted, progressed_to_contacted, progressed_to_ineligible,
     progressed_to_screened, progressed_to_consented, progressed_to_enrolled) = \
        find_progressed_leads(current_id_status, previous_statuses)

    ineligibility_details_text = show_ineligibility_details(current_id_status, previous_statuses)

    # Current total of open new leads (a snapshot count, not date-filtered)
    new_leads_total = sum(
        1 for status in current_id_status.values()
        if "new lead" in status.lower() or "eligible" in status.lower()
    )
    difference = new_leads_total - previous_new_leads_total

    print(f"Contacted: {contacted_total}")
    print(f"Attempted Contact: {attempted_contact_total}")
    print(f"Screened: {screened_total}")
    print(f"Consented: {consented_total}")
    print(f"Enrolled: {enrolled_total}")
    print(f"Ineligible: {ineligible_total}")
    print(f"New Leads: {new_leads_total}")
    print(f"Change since last week: {difference}")
    print(f"New Lead -> Attempted Contact: {progressed_to_attempted}")
    print(f"New Lead -> Contacted: {progressed_to_contacted}")
    print(f"New Lead -> Ineligible: {progressed_to_ineligible}")
    print(f"Ineligibility Details:" + f"{ineligibility_details_text}")
    print(f"New Leads -> Screened: {progressed_to_screened}")
    print(f"New Leads -> Consented: {progressed_to_consented}")
    print(f"New Leads -> Enrolled: {progressed_to_enrolled}")

    # --- Only NOW, after every read of "previous" state is done, do we save ---
    save_current_value("id_status_map", current_id_status)
    save_current_value("new_leads_total", int(new_leads_total))


def send_email(sender_email, sender_password, receiver_email, subject, body):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())


def send_email_button():
    body = (
        "Hello Lupa! "
        + f"This is the weekly update of {one_week_ago.date()} to {today.date()} "
        + f"It is {high_or_low_risk_var.get()} "
        + f"Contacted: {contacted_total} "
        + f"Attempted Contact: {attempted_contact_total} "
        + f"Screened: {screened_total} "
        + f"Consented: {consented_total} "
        + f"Enrolled: {enrolled_total} "
        + f"New Leads: {new_leads_total} "
        + f"Change in new leads: {difference} "
        + f"Ineligible: {ineligible_total} "
        + f"Ineligibility Details: {ineligibility_details_text} "
        + f"New Leads -> Attempted Contact: {progressed_to_attempted} "
        + f"New Leads -> Contacted: {progressed_to_contacted}"
    )
    send_email(
        sender_email=sender_email_var.get(),
        sender_password=sender_email_pass.get(),
        receiver_email=reciever_email_var.get(),
        subject="BuildClinical Weekly Update",
        body=body,
    )


frame = tk.Frame(root)
frame.pack()

container = tk.Frame(root)
container.pack(fill="both", expand=True)

canvas = tk.Canvas(container)
scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="n")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

canvas.bind_all("<MouseWheel>", _on_mousewheel)

tk.Label(scrollable_frame, text="Place BuildClinical File").pack(anchor="center", pady=10)
tk.Button(scrollable_frame, text="Select CSV File", command=select_file).pack(anchor="center", pady=10)

send_email_frame = tk.Frame(scrollable_frame)
send_email_frame.pack(pady=10)

tk.Button(send_email_frame, text="High Risk", command=lambda: high_or_low_risk_var.set("High Risk")).pack(anchor="center", pady=5)
tk.Button(send_email_frame, text="Low Risk", command=lambda: high_or_low_risk_var.set("Low Risk")).pack(anchor="center", pady=5)
tk.Label(send_email_frame, text="Insert Your Email").pack(anchor="center", pady=5)
tk.Entry(send_email_frame, textvariable=sender_email_var).pack(anchor="center", pady=5)
tk.Label(send_email_frame, text="Enter your password").pack(anchor="center", pady=5)
tk.Entry(send_email_frame, textvariable=sender_email_pass, show="*").pack(anchor="center", pady=5)
tk.Label(send_email_frame, text="Enter in reciever email").pack(anchor="center", pady=5)
tk.Entry(send_email_frame, textvariable=reciever_email_var).pack(anchor="center", pady=5)
tk.Button(send_email_frame, text="Send Email", command=send_email_button).pack(anchor="center", pady=5)

root.mainloop()