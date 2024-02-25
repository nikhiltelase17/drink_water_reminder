import tkinter as tk
from tkinter import messagebox
import win32com.client

engine = win32com.client.Dispatch("SAPI.SpVoice")


class WaterReminderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Water Reminder")
        self.root.minsize(width=300, height=400)

        self.label = tk.Label(root, text="Drink Water Reminder", font=("Helvetica", 16))
        self.label.pack(pady=10)

        self.my_label = tk.Label(root, text="Crated by Nikhil Telase")
        self.my_label.pack(pady=2)
        self.interval_label = tk.Label(root, text="Enter reminder interval in minutes:")
        self.interval_label.pack()

        self.interval_entry = tk.Entry(root)
        self.interval_entry.pack(pady=10)

        self.start_button = tk.Button(root, text="Start Reminders", command=self.start_reminders)
        self.start_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Stop Reminders", command=self.stop_reminders)
        self.stop_button.pack(pady=10)
        self.stop_button.config(state=tk.DISABLED)

        self.time_remaining_label = tk.Label(root, text="Time Remaining: --:--", font=("Helvetica", 12))
        self.time_remaining_label.pack(pady=10)

        self.reminder_interval = None
        self.running = False

    def start_reminders(self):
        try:
            self.reminder_interval = int(self.interval_entry.get())
            if self.reminder_interval <= 0:
                messagebox.showerror("Error", "Invalid interval. Please enter a positive value.")
                return

            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)

            messagebox.showinfo("Reminder Started", f"Reminders will be sent every {self.reminder_interval} minutes.")

            self.running = True
            self.run_reminders()

        except ValueError:
            messagebox.showerror("Error", "Invalid interval. Please enter a valid number.")

    def run_reminders(self, time_remaining=None):
        if time_remaining is None:
            time_remaining = self.reminder_interval * 60

        mins, secs = divmod(time_remaining, 60)
        time_str = f"{mins:02d}:{secs:02d}"
        self.time_remaining_label.config(text=f"Time Remaining: {time_str}")

        if time_remaining > 0 and self.running:
            self.root.after(1000, self.run_reminders, time_remaining - 1)
        elif time_remaining == 0 and self.running:
            self.show_reminder()

    def show_reminder(self):
        messagebox.showinfo("Drink Water Reminder", "Hey Nikhil, it's time to drink water.")
        engine.Speak("Hey Nikhil, It's time to drink water.")
        if self.running:
            self.run_reminders()

    def stop_reminders(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.running = False
        self.time_remaining_label.config(text="Time Remaining: --:--")
        messagebox.showinfo("Reminder Stopped", "Reminders stopped. Bye!")


if __name__ == "__main__":
    root = tk.Tk()
    app = WaterReminderApp(root)
    root.mainloop()
