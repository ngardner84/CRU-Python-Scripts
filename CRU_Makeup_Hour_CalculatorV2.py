import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date, timedelta, datetime

# Function to perform the calculations
def calculate():
    try:
        # Fetch inputs
        current_hours = float(entry_current_hours.get())
        missing_hours = float(entry_missing_hours.get())
        total_program_hours = float(entry_total_program_hours.get())
        
        # Optional target hours behind
        target_hours_behind_input = entry_target_hours_behind.get()
        if target_hours_behind_input.strip() == '':
            target_hours_behind = 0  # Default to 0 if no input is given
        else:
            target_hours_behind = float(target_hours_behind_input)
            # Ensure the target is not negative or greater than missing_hours
            target_hours_behind = min(max(target_hours_behind, 0), missing_hours)
        
        adjusted_missing_hours = missing_hours - target_hours_behind
        
        # Start date
        start_date_input = entry_start_date.get()
        start_date = datetime.strptime(start_date_input, "%Y-%m-%d").date()
        
        # Makeup hours schedule
        makeup_hours_schedule = {
            'Monday': float(entry_makeup_monday.get()),
            'Tuesday': float(entry_makeup_tuesday.get()),
            'Wednesday': float(entry_makeup_wednesday.get()),
            'Thursday': float(entry_makeup_thursday.get()),
            'Friday': float(entry_makeup_friday.get()),
        }
        
        # Normal hours schedule
        normal_hours_schedule = {
            'Monday': float(entry_normal_monday.get()),
            'Tuesday': float(entry_normal_tuesday.get()),
            'Wednesday': float(entry_normal_wednesday.get()),
            'Thursday': float(entry_normal_thursday.get()),
            'Friday': float(entry_normal_friday.get()),
        }
        
        # Holidays (predefined holidays)
        holidays = [
            date(2023, 12, 25),  # Christmas Day
            date(2024, 1, 1),    # New Year's Day
            # Add other fixed holidays here
        ]
        
        # User-entered holidays
        holidays_input = text_holidays.get("1.0", tk.END)
        holidays_lines = holidays_input.strip().split('\n')
        for line in holidays_lines:
            line = line.strip()
            if line:
                try:
                    holiday_date = datetime.strptime(line, "%Y-%m-%d").date()
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
            result_text += f"Makeup hours will be completed on: {makeup_completion_date.strftime('%Y-%m-%d')}\n"
        else:
            result_text += "Makeup hours will not be completed within the calculated period or no makeup hours were scheduled.\n"
        
        if graduation_date:
            result_text += f"Graduation will be completed on: {graduation_date.strftime('%Y-%m-%d')}\n"
            result_text += f"Total number of days until graduation: {day_count}\n"
        else:
            result_text += "Graduation will not be completed within the calculated period.\n"
        
        # Create textual representation
        result_text += "\nDay\tDate\t\tHours Earned\tCumulative Hours\n"
        for entry in dates_hours:
            day_number = entry[0]
            date_str = entry[1].strftime("%Y-%m-%d")
            hours_earned = entry[2]
            cumulative_hours = entry[3]
            result_text += f"{day_number}\t{date_str}\t{hours_earned}\t\t{cumulative_hours}\n"
        
        # Display in text widget
        text_output.delete(1.0, tk.END)
        text_output.insert(tk.END, result_text)
        
    except Exception as e:
        messagebox.showerror("Input Error", f"An error occurred: {str(e)}")

# Create main window
root = tk.Tk()
root.title("CRU Makeup Hours Forecasting")

# Create frames
frame_inputs = ttk.Frame(root, padding="10")
frame_inputs.grid(row=0, column=0, sticky="W")

frame_schedule = ttk.Frame(root, padding="10")
frame_schedule.grid(row=1, column=0, sticky="W")

frame_output = ttk.Frame(root, padding="10")
frame_output.grid(row=2, column=0, sticky="W")

# Input fields
ttk.Label(frame_inputs, text="Enter the current hours of the student:").grid(row=0, column=0, sticky="W")
entry_current_hours = ttk.Entry(frame_inputs)
entry_current_hours.grid(row=0, column=1)

ttk.Label(frame_inputs, text="Enter the missing hours of the student:").grid(row=1, column=0, sticky="W")
entry_missing_hours = ttk.Entry(frame_inputs)
entry_missing_hours.grid(row=1, column=1)

ttk.Label(frame_inputs, text="Enter the total program hours for the student:").grid(row=2, column=0, sticky="W")
entry_total_program_hours = ttk.Entry(frame_inputs)
entry_total_program_hours.grid(row=2, column=1)

ttk.Label(frame_inputs, text="Enter the target number of hours to be behind after makeup (optional):").grid(row=3, column=0, sticky="W")
entry_target_hours_behind = ttk.Entry(frame_inputs)
entry_target_hours_behind.grid(row=3, column=1)

ttk.Label(frame_inputs, text="Enter the start date for the plan (YYYY-MM-DD):").grid(row=4, column=0, sticky="W")
entry_start_date = ttk.Entry(frame_inputs)
entry_start_date.grid(row=4, column=1)

# Holidays input
ttk.Label(frame_inputs, text="Enter holidays (YYYY-MM-DD, one per line):").grid(row=5, column=0, sticky="NW")
text_holidays = tk.Text(frame_inputs, width=20, height=5)
text_holidays.grid(row=5, column=1, sticky="W")

# Makeup hours schedule
ttk.Label(frame_schedule, text="Makeup Hours for Each Weekday (can be zero if none):").grid(row=0, column=0, columnspan=2, sticky="W")

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
ttk.Label(frame_schedule, text="Normal Hours for Each Weekday:").grid(row=6, column=0, columnspan=2, sticky="W")

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

# Calculate button
btn_calculate = ttk.Button(frame_schedule, text="Calculate", command=calculate)
btn_calculate.grid(row=12, column=0, columnspan=2, pady=10)

# Output text widget
text_output = tk.Text(frame_output, width=80, height=20)
text_output.grid(row=0, column=0)

# Start the main loop
root.mainloop()
