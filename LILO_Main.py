import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
from datetime import datetime, timedelta
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class TrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation LILO- Tracker")
        self.root.geometry("800x400")
        self.root.configure(bg="blue")

        # Initialize variables
        self.work_mode_var = tk.StringVar(value="WFO")
        self.break_times = []
        self.login_time = None
        self.logout_time = None
        self.total_working_hours = None
        self.register_user()

        # Ensure tracker_log.xlsx exists and is initialized
        self.log_file = "tracker_log.xlsx"
        self.initialize_tracker_log()

        # Create GUI
        self.create_widgets()

    def initialize_tracker_log(self):
        # Create the Excel file if it doesn't exist
        if not os.path.exists(self.log_file):
            columns = [
                "Date", "Username", "User ID", "Work Mode", "Login Time",
                "Each Break Time", "Logout Time", "Total Working Hours"
            ]
            df = pd.DataFrame(columns=columns)
            df.to_excel(self.log_file, index=False)

    def register_user(self):
        # Register user details if it's the first time running the app
        if not os.path.exists("user_info.txt"):
            self.username = simpledialog.askstring("Name", "Please enter your name:")
            self.user_id = simpledialog.askstring("User ID", "Please enter your User ID:")
            self.work_mode = self.work_mode_var.get()

            # Save user details for future use
            with open("user_info.txt", "w") as file:
                file.write(f"{self.username}\n{self.user_id}\n{self.work_mode}")
        else:
            # Read user info from the file if it already exists
            with open("user_info.txt", "r") as file:
                self.username, self.user_id, self.work_mode = file.read().splitlines()

    def create_widgets(self):
        # Welcome Label
        welcome_label = tk.Label(
            self.root, text=f"Welcome {self.username}", bg="blue", fg="white", font=("Helvetica", 16)
        )
        welcome_label.pack()

        # Info Label
        info_label = tk.Label(
            self.root, text=f"{self.user_id}\n{datetime.now().strftime('%m/%d/%Y')}", bg="blue", fg="white", font=("Helvetica", 14)
        )
        info_label.pack()

        # WFO/WFH selection
        wfo_wfh_frame = tk.Frame(self.root, bg="blue")
        wfo_wfh_frame.pack()

        wfo_radio = tk.Radiobutton(
            wfo_wfh_frame, text="WFO", variable=self.work_mode_var, value="WFO",
            bg="blue", fg="white", font=("Helvetica", 14), command=self.on_radio_button_change
        )
        wfo_radio.pack(side="left", padx=10)

        wfh_radio = tk.Radiobutton(
            wfo_wfh_frame, text="WFH", variable=self.work_mode_var, value="WFH",
            bg="blue", fg="white", font=("Helvetica", 14), command=self.on_radio_button_change
        )
        wfh_radio.pack(side="left", padx=10)

        # Buttons
        button_frame = tk.Frame(self.root, bg="blue")
        button_frame.pack(pady=20)

        self.create_button(button_frame, "Login", self.login, 0, 0, "green")
        self.create_button(button_frame, "Break Start", self.break_start, 0, 1, "orange")
        self.create_button(button_frame, "Break End", self.break_end, 0, 2, "orange")
        self.create_button(button_frame, "Logout", self.logout, 0, 3, "red")

        email_button = tk.Button(
            self.root, text="Send Log", command=self.send_email, font=("Helvetica", 14),
            bg="orange", width=20
        )
        email_button.pack(pady=10)

        exit_button = tk.Button(
            self.root, text="Exit", command=self.root.quit, font=("Helvetica", 14), width=10
        )
        exit_button.pack(pady=10)

    def create_button(self, parent, text, command, row, column, color):
        button = tk.Button(parent, text=text, command=command, font=("Helvetica", 12), bg=color, fg="white", width=15)
        button.grid(row=row, column=column, padx=10, pady=10)

    def on_radio_button_change(self):
        self.work_mode = self.work_mode_var.get()

    def log_to_excel(self):
        data = {
            "Date": [datetime.now().strftime('%Y-%m-%d')],
            "Username": [self.username],
            "User ID": [self.user_id],
            "Work Mode": [self.work_mode],
            "Login Time": [self.login_time],
            "Each Break Time": [", ".join(self.break_times)],
            "Logout Time": [self.logout_time],
            "Total Working Hours": [self.total_working_hours]
        }
        new_df = pd.DataFrame(data)

        # Append to the log file
        with pd.ExcelWriter(self.log_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            existing_df = pd.read_excel(self.log_file)
            updated_df = pd.concat([existing_df, new_df], ignore_index=True)
            updated_df.to_excel(writer, index=False)

    def calculate_total_hours(self):
        login_dt = datetime.strptime(self.login_time, '%Y-%m-%d %H:%M:%S')
        logout_dt = datetime.strptime(self.logout_time, '%Y-%m-%d %H:%M:%S')
        total_hours = logout_dt - login_dt

        break_duration = timedelta()
        for break_time in self.break_times:
            start, end = map(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S'), break_time.split(" - "))
            break_duration += (end - start)

        total_work_duration = total_hours - break_duration
        self.total_working_hours = str(total_work_duration)

    def login(self):
        self.login_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        messagebox.showinfo("Info", "Logged In")

    def break_start(self):
        start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.break_times.append(f"{start_time} - ")
        messagebox.showinfo("Info", "Break Started")

    def break_end(self):
        if not self.break_times or self.break_times[-1].endswith(" - "):
            end_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.break_times[-1] = f"{self.break_times[-1].strip()} {end_time}"
            messagebox.showinfo("Info", "Break Ended")
        else:
            messagebox.showwarning("Warning", "No break in progress!")

    def logout(self):
        self.logout_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.calculate_total_hours()
        self.log_to_excel()
        messagebox.showinfo("Info", "Logged Out")

    def send_email(self):
        email = simpledialog.askstring("Email", "Enter recipient's email:")
        if not email:
            messagebox.showerror("Error", "Email address is required!")
            return

        subject = "Daily Activity Log"
        body = "Attached is the daily tracker log."
        sender_email = "bvhss20@gmail.com"
        sender_password = "yqup nkss xket bpkt"  

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with open(self.log_file, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(self.log_file)}")
            msg.attach(part)

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email, msg.as_string())
            messagebox.showinfo("Info", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TrackerApp(root)
    root.mainloop()
