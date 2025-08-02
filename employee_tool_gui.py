import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.simpledialog import Dialog
import pandas as pd
import os
from fpdf import FPDF
import smtplib
from email.message import EmailMessage

class EmailCredentialsDialog(Dialog):
    def body(self, master):
        self.title("Send Email")

        tk.Label(master, text="Sender Gmail:").grid(row=0)
        tk.Label(master, text="App Password:").grid(row=1)
        tk.Label(master, text="Recipient Email:").grid(row=2)

        self.sender_entry = tk.Entry(master, width=40)
        self.app_password_entry = tk.Entry(master, width=40, show='*')
        self.recipient_entry = tk.Entry(master, width=40)

        self.sender_entry.grid(row=0, column=1, padx=5, pady=5)
        self.app_password_entry.grid(row=1, column=1, padx=5, pady=5)
        self.recipient_entry.grid(row=2, column=1, padx=5, pady=5)

        return self.sender_entry

    def apply(self):
        self.sender = self.sender_entry.get()
        self.password = self.app_password_entry.get()
        self.recipient = self.recipient_entry.get()

class EmployeeToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Comparison Tool")
        self.root.geometry("1100x700")
        self.root.configure(bg="#2c3e50")

        self.master_file_path = tk.StringVar()
        self.changes_file_path = tk.StringVar()

        style = ttk.Style() 
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10), padding=6, background="#3498db", foreground="white")
        style.configure("TLabel", font=("Segoe UI", 10), background="#2c3e50", foreground="white")
        style.configure("Treeview", background="#ecf0f1", foreground="black", rowheight=25, fieldbackground="#ecf0f1", borderwidth=1, relief="solid")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#34495e", foreground="white")

        header = tk.Label(self.root, text="Employee Comparison Dashboard", font=("Segoe UI", 16, "bold"), bg="#2c3e50", fg="#f1c40f")
        header.pack(pady=10)

        file_frame = tk.Frame(root, bg="#2c3e50")
        file_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(file_frame, text="Master CSV File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(file_frame, textvariable=self.master_file_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_master).grid(row=0, column=2, padx=5)

        ttk.Label(file_frame, text="Changes CSV File:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(file_frame, textvariable=self.changes_file_path, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_changes).grid(row=1, column=2, padx=5)

        button_frame = tk.Frame(root, bg="#2c3e50")
        button_frame.pack(pady=15)

        ttk.Button(button_frame, text="Run Comparison", command=self.run_comparison).grid(row=0, column=0, padx=10)
        ttk.Button(button_frame, text="View Changes File", command=self.show_changes).grid(row=0, column=1, padx=10)
        ttk.Button(button_frame, text="View Counts File", command=self.show_counts).grid(row=0, column=2, padx=10)
        ttk.Button(button_frame, text="View New Joines File", command=self.show_new_joiners).grid(row=0, column=3, padx=10)
        ttk.Button(button_frame, text="Clear Table", command=self.clear_table).grid(row=0, column=4, padx=10)
        ttk.Button(button_frame, text="Send Email", command=self.send_email_with_reports).grid(row=0, column=5, padx=10)

        tree_frame = tk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        self.tree = ttk.Treeview(tree_frame, show="headings", yscrollcommand=tree_scroll_y.set)
        self.tree.pack(fill="both", expand=True)
        tree_scroll_y.config(command=self.tree.yview)

    def browse_master(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.master_file_path.set(file_path)

    def browse_changes(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.changes_file_path.set(file_path)

    def run_comparison(self):
        # [omitted for brevity, keep your logic unchanged]
        pass

    def generate_pdf_report(self, df):
        # [omitted for brevity, keep your logic unchanged]
        pass

    def send_email_with_reports(self):
        dialog = EmailCredentialsDialog(self.root)
        sender_email = dialog.sender
        app_password = dialog.password
        receiver_email = dialog.recipient

        if not all([sender_email, app_password, receiver_email]):
            messagebox.showwarning("Input Error", "All fields must be filled.")
            return

        try:
            msg = EmailMessage()
            msg["Subject"] = "Employee Reports"
            msg["From"] = sender_email
            msg["To"] = receiver_email
            msg.set_content("Attached are the employee comparison reports.")

            attachments = ["changes_report.csv", "count_report.csv", "new_joiners_report.csv", "changes_report.pdf"]
            for file in attachments:
                with open(file, "rb") as f:
                    file_data = f.read()
                    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file)

            with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
                smtp.starttls()
                smtp.login(sender_email, app_password)
                smtp.send_message(msg)

            messagebox.showinfo("Email Sent", f"Reports sent to {receiver_email} successfully.")
        except Exception as e:
            messagebox.showerror("Email Error", str(e))

    def show_changes(self):
        self.show_file("changes_report.csv")

    def show_counts(self):
        self.show_file("count_report.csv")

    def show_new_joiners(self):
        self.show_file("new_joiners_report.csv")

    def show_file(self, filename):
        try:
            df = pd.read_csv(filename)
            df.replace(["nan", "NaN", pd.NA, None], "-", inplace=True)
            df.fillna("-", inplace=True)
            self.clear_table()
            self.tree["columns"] = list(df.columns)
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=150, anchor="center")
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=list(row))
        except:
            messagebox.showerror("Error", f"{filename} not found.")

    def clear_table(self):
        self.tree.delete(*self.tree.get_children())

if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeToolApp(root)
    root.mainloop()
