import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date, timedelta


# Function to display message box
def show_message(msg_text):
    root = tk.Tk()  # Create a root window
    root.withdraw()  # Hide the root window
    messagebox.showinfo("Message", msg_text)  # Display the message box
    root.destroy()  # Close the root window when the message box is closed


def get_previous_month_dates():
    """Return the first and last day of the previous month as DD.MM.YYYY."""
    today = date.today()
    first_this_month = today.replace(day=1)
    last_prev_month = first_this_month - timedelta(days=1)
    first_prev_month = last_prev_month.replace(day=1)
    return first_prev_month.strftime("%d.%m.%Y"), last_prev_month.strftime("%d.%m.%Y")


def get_last_six_months_boundaries():
    """Return the first and last month of the last six months period as MM.YYYY."""
    today = date.today()
    # Last month is always previous month
    last_month = today.month - 1 or 12
    last_year = today.year if today.month > 1 else today.year - 1

    # First month = 5 months before last_month
    first_month = last_month - 5
    first_year = last_year
    while first_month <= 0:
        first_month += 12
        first_year -= 1

    return f"{first_month:02d}.{first_year}", f"{last_month:02d}.{last_year}"


def adjust_previous_month_dates_by_user(message):
    """Show a small form to let user adjust first and last day of previous month."""
    def on_ok():
        nonlocal first_day, last_day
        first_day = entry_first.get()
        last_day = entry_last.get()
        root.destroy()

    # Default values
    first_default, last_default = get_previous_month_dates()
    first_month_of_last_six_years, last_month_of_last_six_years = get_last_six_months_boundaries()

    # Create window
    root = tk.Tk()
    root.title("Select Previous Month Dates")
    root.geometry("300x350")

    # Labels and inputs
    ttk.Label(root, text=f"{message}\n").pack(pady=5)

    ttk.Label(root, text="First day of previous month:").pack(pady=5)
    entry_first = ttk.Entry(root)
    entry_first.insert(0, first_default)
    entry_first.pack()

    ttk.Label(root, text="Last day of previous month:").pack(pady=5)
    entry_last = ttk.Entry(root)
    entry_last.insert(0, last_default)
    entry_last.pack()

    # gap
    ttk.Label(root, text="").pack(pady=5)

    ttk.Label(root, text="ZEK 1 range from:").pack(pady=5)
    six_months_first = ttk.Entry(root)
    six_months_first.insert(0, first_month_of_last_six_years)
    six_months_first.pack()

    ttk.Label(root, text="ZEK 1 range to:").pack(pady=5)
    six_months_last = ttk.Entry(root)
    six_months_last.insert(0, last_month_of_last_six_years)
    six_months_last.pack()

    # gap
    ttk.Label(root, text="").pack(pady=5)

    # OK button
    ttk.Button(root, text="OK", command=on_ok).pack(pady=10)

    # Variables to be filled in by form
    first_day = ""
    last_day = ""
    first_month_of_last_six_years = ""
    last_month_of_last_six_years = ""

    root.mainloop()
    return first_day, last_day, first_month_of_last_six_years, last_month_of_last_six_years
