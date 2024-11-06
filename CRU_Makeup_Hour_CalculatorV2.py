import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date, timedelta, datetime

# Function to convert hours input, accepting "HOURS:MINUTES" format
def parse_hours_input(input_str):
    try:
        if ':' in input_str:
            hours_str, minutes_str = input_str.split(':')
            hours = float(hours_str)
            minutes = float(minutes_str)
            total_hours = hours + minutes / 60
        else:
            total_hours = float(input_str)
        return total_hours
    except ValueError:
        raise ValueError(f"Invalid hours format: {input_str}")

# Function to perform the calculations
def calculate():
    try:
        # Fetch inputs and parse hours
        current_hours_input = entry_current_hours.get()
        current_hours = parse_hours_input(current_hours_input)

        missing_hours_input = entry_missing_hours.get()
        missing_hours = parse_hours_input(missing_hours_input)

        total_program_hours_input = entry_total_program_hours.get()
        total_program_hours = parse_hours_input(total_program_hours_input)

        # Optional target hours behind
        target_hours_behind_input = entry_target_hours_behind.get()
        if target_hours_behind_input.strip() == '':
            target_hours_behind = 0  # Default to 0 if no input is given
        else:
            target_hours_behind = parse_hours_input(target_hours_behind_input)
            # Ensure the target is not negative or greater than missing_hours
            target_hours_behind = min(max(target_hours_behind, 0), missing_hours)

        adjusted_missing_hours = missing_hours - target_hours_behind

        # Start date
        start_date_input = entry_start_date.get()
        start_date = datetime.strptime(start_date_input, "%m-%d-%Y").date()

        # Makeup hours schedule
        makeup_hours_schedule = {
            'Monday': parse_hours_input(entry_makeup_monday.get()),
            'Tuesday': parse_hours_input(entry_makeup_tuesday.get()),
            'Wednesday': parse_hours_input(entry_makeup_wednesday.get()),
            'Thursday': parse_hours_input(entry_makeup_thursday.get()),
            'Friday': parse_hours_input(entry_makeup_friday.get()),
        }

        # Normal hours schedule
        normal_hours_schedule = {
            'Monday': parse_hours_input(entry_normal_monday.get()),
            'Tuesday': parse_hours_input(entry_normal_tuesday.get()),
            'Wednesday': parse_hours_input(entry_normal_wednesday.get()),
            'Thursday': parse_hours_input(entry_normal_thursday.get()),
            'Friday': parse_hours_input(entry_normal_friday.get()),
        }

        # Collect holidays from the text field
        holidays_input = text_holidays.get("1.0", tk.END)
        holidays_lines = holidays_input.strip().split('\n')
        holidays = []
        for line in holidays_lines:
            line = line.strip()
            if line:
                try:
                    holiday_date = datetime.strptime(line, "%m-%d-%Y").date()
                    holidays.append(holiday_date)
                except ValueError:
                    messagebox.showerror("Input Error", f"Invalid holiday date format: {line}")
                    return

        # Initialize variables
        current_date = start_date
        makeup_hours_remaining = adjusted_missing_hours
        total_hours_earned = current_hours
        total_hours_needed = total_program_hours - current_hours
        makeup_completion_date = None
        graduation_date = None
        dates_hours = []
        day_count = 0

        while total_hours_earned < total_program_hours:
            # Check if current_date is a weekday (Monday=0, Sunday=6) and not a holiday
            if current_date.weekday() < 5 and current_date not in holidays:
                weekday_name = current_date.strftime("%A")
                # Get normal hours for this weekday
                normal_hours = normal_hours_schedule.get(weekday_name, 0)

                # Initialize today's hours
                today_hours = normal_hours

                # If makeup hours are still remaining, add makeup hours
                if makeup_hours_remaining > 0:
                    makeup_hours = makeup_hours_schedule.get(weekday_name, 0)
                    makeup_hours = min(makeup_hours, makeup_hours_remaining)
                    today_hours += makeup_hours
                    makeup_hours_remaining -= makeup_hours

                    # Check if makeup hours are completed
                    if makeup_hours_remaining == 0 and makeup_completion_date is None:
                        makeup_completion_date = current_date
                else:
                    makeup_hours = 0  # No makeup hours today

                # Add today's hours to total
                total_hours_earned += today_hours

                # Increment day count
                day_count += 1

                # Append to dates_hours
                dates_hours.append((day_count, current_date, today_hours, total_hours_earned))

                # Check if total program hours are met
                if total_hours_earned >= total_program_hours and graduation_date is None:
                    graduation_date = current_date
            else:
                # Non-working day or holiday
                pass

            # Move to the next day
            current_date += timedelta(days=1)

        # Display results
        result_text = ""
        if makeup_completion_date:
            result_text += f"Makeup hours will be completed on: {makeup_completion_date.strftime('%m-%d-%Y')}\n"
        else:
            result_text += "Makeup hours will not be completed within the calculated period or no makeup hours were scheduled.\n"

        if graduation_date:
            result_text += f"Graduation will be completed on: {graduation_date.strftime('%m-%d-%Y')}\n"
            result_text += f"Total number of days until graduation: {day_count}\n"
        else:
            result_text += "Graduation will not be completed within the calculated period.\n"

        # Create textual representation
        result_text += "\nDay\tDate\t\tHours Earned\tCumulative Hours\n"
        for entry in dates_hours:
            day_number = entry[0]
            date_str = entry[1].strftime("%m-%d-%Y")
            hours_earned = entry[2]
            cumulative_hours = entry[3]
            result_text += f"{day_number}\t{date_str}\t{hours_earned}\t\t{cumulative_hours}\n"

        # Display in text widget
        text_output.delete(1.0, tk.END)
        text_output.insert(tk.END, result_text)

    except Exception as e:
        messagebox.showerror("Input Error", f"An error occurred: {str(e)}")

# Predefined normal hours schedules
predefined_schedules = {
    "Day 34 Hour": {
        'Monday': 7,
        'Tuesday': 7,
        'Wednesday': 7,
        'Thursday': 7,
        'Friday': 6,
    },
    "Day 28 Hour": {
        'Monday': 5.5,
        'Tuesday': 5.5,
        'Wednesday': 5.5,
        'Thursday': 5.5,
        'Friday': 6,
    },
    "Night 24 Hour": {
        'Monday': 6,
        'Tuesday': 6,
        'Wednesday': 6,
        'Thursday': 6,
        'Friday': 0,
    },
    "Night 20 Hour": {
        'Monday': 5,
        'Tuesday': 5,
        'Wednesday': 5,
        'Thursday': 5,
        'Friday': 0,
    },
    "Day/Night 24 Hour": {
        'Monday': 5,
        'Tuesday': 5,
        'Wednesday': 5,
        'Thursday': 5,
        'Friday': 4,
    },
}

# Predefined holidays (internal variable)
predefined_holidays = [
    date(2024, 11, 7),    # Holiday
    date(2024, 11, 28),   # Holiday
    date(2024, 11, 29),   # Holiday
    date(2024, 12, 23),
    date(2024, 12, 24),
    date(2024, 12, 25),
    date(2024, 12, 26),
    date(2024, 12, 27),
    date(2024, 12, 28),
    date(2024, 12, 29),
    date(2024, 12, 30),
    date(2024, 12, 31),
    date(2025, 1, 1),
    date(2025, 1, 20),
    date(2025, 2, 17),
    date(2025, 5, 26),
    date(2025, 6, 19),
    date(2025, 7, 4),
    date(2025, 7, 5),
    date(2025, 9, 2),
    date(2025, 10, 14),
    # Add other fixed holidays here
]

# Function to load predefined holidays into the text widget
def load_predefined_holidays():
    holidays_text = ''
    for holiday in predefined_holidays:
        holidays_text += holiday.strftime("%m-%d-%Y") + '\n'
    text_holidays.insert(tk.END, holidays_text)

# Function to load a predefined schedule into the normal hours entries
def load_schedule(schedule_name):
    schedule = predefined_schedules[schedule_name]
    entry_normal_monday.delete(0, tk.END)
    entry_normal_monday.insert(0, str(schedule['Monday']))
    entry_normal_tuesday.delete(0, tk.END)
    entry_normal_tuesday.insert(0, str(schedule['Tuesday']))
    entry_normal_wednesday.delete(0, tk.END)
    entry_normal_wednesday.insert(0, str(schedule['Wednesday']))
    entry_normal_thursday.delete(0, tk.END)
    entry_normal_thursday.insert(0, str(schedule['Thursday']))
    entry_normal_friday.delete(0, tk.END)
    entry_normal_friday.insert(0, str(schedule['Friday']))

# Create main window
root = tk.Tk()
root.title("CRU Makeup Hours Forecasting")

# Set default window size
root.geometry("1000x800")  # Width x Height in pixels

# Create a canvas with scrollbars
main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=1)

canvas = tk.Canvas(main_frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

scrollbar_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

scrollbar_x = ttk.Scrollbar(root, orient=tk.HORIZONTAL, command=canvas.xview)
scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Create another frame inside the canvas
content_frame = ttk.Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

# Now, build the GUI inside content_frame
# Create frames
frame_inputs = ttk.Frame(content_frame, padding="10")
frame_inputs.grid(row=0, column=0, sticky="W")

frame_schedule = ttk.Frame(content_frame, padding="10")
frame_schedule.grid(row=1, column=0, sticky="W")

frame_output = ttk.Frame(content_frame, padding="10")
frame_output.grid(row=2, column=0, sticky="W")

# Input fields
ttk.Label(frame_inputs, text="Enter the current hours of the student (HOURS or HOURS:MINUTES):").grid(row=0, column=0, sticky="W")
entry_current_hours = ttk.Entry(frame_inputs)
entry_current_hours.grid(row=0, column=1)

ttk.Label(frame_inputs, text="Enter the missing hours of the student (HOURS or HOURS:MINUTES):").grid(row=1, column=0, sticky="W")
entry_missing_hours = ttk.Entry(frame_inputs)
entry_missing_hours.grid(row=1, column=1)

ttk.Label(frame_inputs, text="Enter the total program hours for the student (HOURS or HOURS:MINUTES):").grid(row=2, column=0, sticky="W")
entry_total_program_hours = ttk.Entry(frame_inputs)
entry_total_program_hours.grid(row=2, column=1)

ttk.Label(frame_inputs, text="Enter the target number of hours to be behind after makeup (optional):").grid(row=3, column=0, sticky="W")
entry_target_hours_behind = ttk.Entry(frame_inputs)
entry_target_hours_behind.grid(row=3, column=1)

ttk.Label(frame_inputs, text="Enter the start date for the plan (MM-DD-YYYY):").grid(row=4, column=0, sticky="W")
entry_start_date = ttk.Entry(frame_inputs)
entry_start_date.grid(row=4, column=1)

# Holidays input
ttk.Label(frame_inputs, text="Holidays (MM-DD-YYYY, one per line):").grid(row=5, column=0, sticky="NW")
text_holidays = tk.Text(frame_inputs, width=20, height=10)
text_holidays.grid(row=5, column=1, sticky="W")

# Load predefined holidays into the text widget
load_predefined_holidays()

# Makeup hours schedule
ttk.Label(frame_schedule, text="Makeup Hours for Each Weekday (HOURS or HOURS:MINUTES):").grid(row=0, column=0, columnspan=2, sticky="W")

ttk.Label(frame_schedule, text="Monday:").grid(row=1, column=0, sticky="W")
entry_makeup_monday = ttk.Entry(frame_schedule)
entry_makeup_monday.grid(row=1, column=1)

ttk.Label(frame_schedule, text="Tuesday:").grid(row=2, column=0, sticky="W")
entry_makeup_tuesday = ttk.Entry(frame_schedule)
entry_makeup_tuesday.grid(row=2, column=1)

ttk.Label(frame_schedule, text="Wednesday:").grid(row=3, column=0, sticky="W")
entry_makeup_wednesday = ttk.Entry(frame_schedule)
entry_makeup_wednesday.grid(row=3, column=1)

ttk.Label(frame_schedule, text="Thursday:").grid(row=4, column=0, sticky="W")
entry_makeup_thursday = ttk.Entry(frame_schedule)
entry_makeup_thursday.grid(row=4, column=1)

ttk.Label(frame_schedule, text="Friday:").grid(row=5, column=0, sticky="W")
entry_makeup_friday = ttk.Entry(frame_schedule)
entry_makeup_friday.grid(row=5, column=1)

# Normal hours schedule
ttk.Label(frame_schedule, text="Normal Hours for Each Weekday (HOURS or HOURS:MINUTES):").grid(row=6, column=0, columnspan=2, sticky="W")

ttk.Label(frame_schedule, text="Monday:").grid(row=7, column=0, sticky="W")
entry_normal_monday = ttk.Entry(frame_schedule)
entry_normal_monday.grid(row=7, column=1)

ttk.Label(frame_schedule, text="Tuesday:").grid(row=8, column=0, sticky="W")
entry_normal_tuesday = ttk.Entry(frame_schedule)
entry_normal_tuesday.grid(row=8, column=1)

ttk.Label(frame_schedule, text="Wednesday:").grid(row=9, column=0, sticky="W")
entry_normal_wednesday = ttk.Entry(frame_schedule)
entry_normal_wednesday.grid(row=9, column=1)

ttk.Label(frame_schedule, text="Thursday:").grid(row=10, column=0, sticky="W")
entry_normal_thursday = ttk.Entry(frame_schedule)
entry_normal_thursday.grid(row=10, column=1)

ttk.Label(frame_schedule, text="Friday:").grid(row=11, column=0, sticky="W")
entry_normal_friday = ttk.Entry(frame_schedule)
entry_normal_friday.grid(row=11, column=1)

# Predefined schedule buttons
ttk.Label(frame_schedule, text="Preload Normal Hours Schedule:").grid(row=12, column=0, columnspan=2, sticky="W", pady=(10, 0))

btn_day_34 = ttk.Button(frame_schedule, text="Day 34 Hour", command=lambda: load_schedule("Day 34 Hour"))
btn_day_34.grid(row=13, column=0, sticky="W", pady=2)

btn_day_28 = ttk.Button(frame_schedule, text="Day 28 Hour", command=lambda: load_schedule("Day 28 Hour"))
btn_day_28.grid(row=13, column=1, sticky="W", pady=2)

btn_night_24 = ttk.Button(frame_schedule, text="Night 24 Hour", command=lambda: load_schedule("Night 24 Hour"))
btn_night_24.grid(row=14, column=0, sticky="W", pady=2)

btn_night_20 = ttk.Button(frame_schedule, text="Night 20 Hour", command=lambda: load_schedule("Night 20 Hour"))
btn_night_20.grid(row=14, column=1, sticky="W", pady=2)

btn_day_night_24 = ttk.Button(frame_schedule, text="Day/Night 24 Hour", command=lambda: load_schedule("Day/Night 24 Hour"))
btn_day_night_24.grid(row=15, column=0, sticky="W", pady=2)

# Calculate button
btn_calculate = ttk.Button(frame_schedule, text="Calculate", command=calculate)
btn_calculate.grid(row=16, column=0, columnspan=2, pady=10)

# Output text widget
ttk.Label(frame_output, text="Calculation Results:").grid(row=0, column=0, sticky="W")
text_output = tk.Text(frame_output, width=80, height=20)
text_output.grid(row=1, column=0)

# Start the main loop
root.mainloop()
