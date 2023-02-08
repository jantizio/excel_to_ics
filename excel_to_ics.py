import os, uuid, sys, readchar
import pandas as pd
import tkinter as tk
from icalendar import Calendar, Event
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
from tkinter import filedialog


def end_program():
    print("Premi un tasto per chiudere...")
    k = readchar.readchar()
    sys.exit()


def display(cal):
    return cal.to_ical().decode("utf-8").replace("\r\n", "\n").strip()


def is_good_entry(event_info):
    day, work = event_info
    if work not in accepted_works:
        return False

    now = datetime.now()
    if day < now:
        return False

    return True


def parse_work(work_id):
    switch = {"M": "Lavoro Mattina", "P": "Lavoro Pomeriggio", "N": "Lavoro Notte"}
    return switch.get(work_id, "Errore di lettura")


def create_event(event_info):
    day, work = event_info
    day_date = day.date()
    day_after = day_date + timedelta(days=1)

    event = Event()
    event.add("summary", work)
    event.add("dtstamp", day)
    event.add("dtstart", day_date)
    event.add("dtend", day_after)
    event.add("uid", str(uuid.uuid4()))
    return event


days_index = 4
work_turn_index = 7
accepted_works = ("M", "N", "P")
line_offset = 2
current_year = datetime.now().year

# define the necessary file path
# Determine the location of the script/executable
if getattr(sys, "frozen", False):
    # The script is running as an executable
    exec_dir = os.path.dirname(sys.executable)
else:
    # The script is running as a script
    exec_dir = os.path.dirname(os.path.abspath(__file__))
# excel_file_path = os.path.join(exec_dir, "Febbraio Marzo 2023.xlsx")
calendar_file_path = os.path.join(exec_dir, "lavoro.ics")

# open a gui to select the excel file
root = tk.Tk()
root.withdraw()

excel_file_path = filedialog.askopenfilename()

# open the file and save it into a pandas DataFrame
try:
    df = pd.read_excel(excel_file_path)
except FileNotFoundError:
    print("Il file specificato non è stato trovato.")
    end_program()
except Exception as e:
    print("Si è verificato un errore")
    end_program()

work_turn_index = (
    int(input("Inserisci la riga del file excel di cui vuoi il calendario: "))
    - line_offset
)

# save the date and the type of work into separate lists
days = df.loc[days_index].tolist()
days = [
    day.replace(year=current_year) for day in days
]  # set the year to current year to prevent mistake in the excel
works = df.loc[work_turn_index].tolist()

# create a list of couples (day, work)
calendar_list = list(zip(days, works))
# filter the list to remove non working days and past dates
calendar_list = list(filter(is_good_entry, calendar_list))
# convert the letter in into the corresponding text
calendar_list = [(day, parse_work(work)) for (day, work) in calendar_list]

# Print the list
# print(calendar_list)

# calendar init
cal = Calendar()
cal.add("version", "2.0")
cal.add("calscale", "GREGORIAN")

for event_info in calendar_list:
    event = create_event(event_info)
    cal.add_component(event)


# print(display(cal))

# print the calendar to file
with open(calendar_file_path, "wb") as f:
    f.write(cal.to_ical())
