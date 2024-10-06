# Program to calculate barber school makeup hours and the finishing date
import datetime
import json
import calendar
from tabulate import tabulate

def print_month_calendar(year, month, daily_hours, initial_hours_missed):
    _, num_days = calendar.monthrange(year, month)
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Remaining Hours"]
    calendar_data = [[] for _ in range(6)]  # Assuming maximum 6 weeks in any month

    # Find starting position in the first week
    first_day_of_month = datetime.date(year, month, 1)
    start_index = first_day_of_month.weekday()  # Monday is 0, Sunday is 6 (ISO)
    
    week_row = 0
    total_hours_used = 0
    for day in range(1, num_days + 1):
        current_date = datetime.date(year, month, day)
        if current_date in daily_hours:
            day_hours = daily_hours[current_date]  # Fetch hours for the current date
            total_hours_used += day_hours
        else:
            day_hours = 0
        
        remaining_hours = max(0, initial_hours_missed - sum(total_hours_used for total_hours_used in daily_hours.values()))

        # Fill the row till starting day
        if day == 1:
            calendar_data[week_row].extend(["    "] * (start_index + 1))  # Add empty strings for days before the first of the month
        
        # Add day and hours to calendar data, and remaining hours
        calendar_data[week_row].append(f"{day} ({day_hours})")
        
        # On the last day of the week or the last day of the month, add remaining hours
        if (start_index + day) % 7 == 0 or day == num_days:
            calendar_data[week_row].append(f"{remaining_hours:.1f}")
            week_row += 1  # Move to the next week row if not the end of the month

    # Fill remaining cells in the last row with empty spaces if needed
    if len(calendar_data[week_row - 1]) < 8:
        calendar_data[week_row - 1].extend(["    "] * (8 - len(calendar_data[week_row - 1])))

    # Print calendar using tabulate
    print(f"\n{calendar.month_name[month]} {year}")
    print(tabulate(calendar_data, headers=days, tablefmt="fancy_grid", stralign='right', numalign='center'))
    
# Define all student schedules as dictionaries
day_34 = { 
    "Sunday": 0,
    "Monday": 7,
    "Tuesday": 7,
    "Wednesday": 7,
    "Thursday": 7,
    "Friday": 6,
    "Saturday": 0
}

day_28 = {
    "Sunday": 0,
    "Monday": 5.5,
    "Tuesday": 5.5,
    "Wednesday": 5.5,
    "Thursday": 5.5,
    "Friday": 5.5,
    "Saturday": 0
}

day_night_24 = {
    "Sunday": 0,
    "Monday": 5,
    "Tuesday": 5,
    "Wednesday": 5,
    "Thursday": 5,
    "Friday": 4,
    "Saturday": 0
}

night_24 = {
    "Sunday": 0, 
    "Monday": 6,
    "Tuesday": 6,
    "Wednesday": 6,
    "Thursday": 6,
    "Friday": 0,
    "Saturday": 0
}

night_20 = {
    "Sunday": 0,
    "Monday": 5,
    "Tuesday": 5,
    "Wednesday": 5,
    "Thursday": 5,
    "Friday": 0,
    "Saturday": 0
}

# Define holidays
holidays = {datetime.datetime.strptime(date, "%m-%d-%Y") for date in [
    "06-19-2024", "07-04-2024", "07-05-2024", "09-02-2024", "10-14-2024", 
    "11-11-2024", "11-28-2024", "11-29-2024", "12-23-2024", "12-25-2024"
]}

possible_schedules = [day_34, day_28, day_night_24, night_24, night_20]
# Get the start date
start_date = input("Enter the start date for the makeup hour plan (MM-DD-YYYY): ")
start_date = datetime.datetime.strptime(start_date, "%m-%d-%Y")
calendar_start_date = start_date

current_hours = float(input("Enter the current hours of the student: "))

# Get the student hours schedule
student_hours = int(input("Enter the schedule of the student: 1. 34 hours, 2. 28 hours, 3. 24 hours (Day/Night), 4. 24 hours (Night), 5. 20 hours (Night): "))
student_schedule = possible_schedules[student_hours - 1]
print("Student schedule: " + json.dumps(student_schedule))

# Get the number of hours missed
hours_missed = float(input("Enter the number of hours missed: "))
calendar_hours_missed = hours_missed


# Input weekly makeup hours
weekly_makeup_hours = {}
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
for day in days:
    if student_schedule[day] > 0:  # Only ask for input if there are originally scheduled hours
        weekly_makeup_hours[day] = float(input(f"Enter the number of makeup hours on {day}: "))
    else:
        weekly_makeup_hours[day] = 0


daily_hours_done = {}
# Calculate the end date
while hours_missed > 0 and current_hours < 1200:
    day_name = start_date.strftime("%A")
    if start_date not in holidays and student_schedule[day_name] > 0:
        day_name = start_date.strftime("%A")
        daily_hours_done[start_date.date()] = student_schedule[day_name] + weekly_makeup_hours[day_name]
        daily_hours = student_schedule[day_name]
        hours_missed -= weekly_makeup_hours[day_name]
        current_hours += daily_hours + weekly_makeup_hours[day_name]
    start_date += datetime.timedelta(days=1)
    
end_date = start_date - datetime.timedelta(days=1)
# Print each month's calendar from the start date to the end date
current_month = end_date.month
current_year = end_date.year

calendar_month = current_month
calendar_year = current_year


while calendar_year < end_date.year or (calendar_year == end_date.year and calendar_month <= end_date.month):
    print_month_calendar(calendar_start_date.year, calendar_start_date.month, daily_hours_done, calendar_hours_missed)
    calendar_month += 1
    if calendar_month == 13:
        calendar_month = 1
        calendar_year += 1

print(f"The student will finish on {start_date.strftime('%m-%d-%Y')}")
print(f"The student will have {current_hours} hours")
# End of program

