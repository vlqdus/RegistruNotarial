import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import openpyxl

# Database
conn = sqlite3.connect("registry.db")
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS registry (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        act_number TEXT,
        date TEXT,
        full_name TEXT,
        birth_date TEXT,
        personal_id TEXT,
        act_type TEXT,
        state_fee TEXT,
        assistance_payment TEXT,
        notes TEXT
    )
''')
conn.commit()

# Functions
def save_record():
    values = tuple(entry.get().strip() for entry in entries)
    try:
        cursor.execute('''
            INSERT INTO registry (act_number, date, full_name, birth_date, personal_id,
                                  act_type, state_fee, assistance_payment, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', values)
        conn.commit()
        status_label.config(text="✅ Saved!")
        display_records()

        # Clear input fields after saving
        for entry in entries:
            if isinstance(entry, tk.Entry):
                entry.delete(0, tk.END)
            elif isinstance(entry, ttk.Combobox):
                entry.set("")

    except Exception as e:
        status_label.config(text=f"❌ Error: {e}")


def search_records():
    query = search_entry.get().strip()
    for row in table.get_children():
        table.delete(row)
    cursor.execute("""
        SELECT * FROM registry WHERE 
            full_name LIKE ? OR 
            birth_date LIKE ? OR 
            personal_id LIKE ? OR 
            act_number LIKE ?
    """, (f"%{query}%",)*4)
    for row in cursor.fetchall():
        table.insert('', 'end', values=row)


def display_records():
    for row in table.get_children():
        table.delete(row)
    cursor.execute("SELECT * FROM registry")
    for row in cursor.fetchall():
        table.insert('', 'end', values=row)


def delete_record():
    selected = table.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a record to delete.")
        return
    record_id = table.item(selected[0])['values'][0]
    cursor.execute("DELETE FROM registry WHERE id=?", (record_id,))
    conn.commit()
    display_records()
    status_label.config(text="✅ Deleted successfully!")


def export_to_excel():
    cursor.execute("SELECT * FROM registry")
    rows = cursor.fetchall()
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Notarial Registry"
    worksheet.append(columns)
    for row in rows:
        worksheet.append(row)
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        workbook.save(filepath)
        messagebox.showinfo("Exported", "Data successfully exported!")


# Application
app = tk.Tk()
app.title("Notarial Registry")
app.geometry("1280x720")
app.configure(bg="#2c2f33")

# Fonts
font_base = ("Segoe UI", 12)
font_bold = ("Segoe UI", 12, "bold")
text_color = "#ffffff"

# -------------- FORM --------------
form_frame = tk.Frame(app, bg="#2c2f33")
form_frame.pack(padx=20, pady=20, anchor="nw")

labels_text = [
    "Act Number", "Date", "Full Name", "Birth Date", "Personal ID",
    "Act Type", "State Fee", "Assistance Payment", "Notes"
]
entries = []

act_types = [
    "Sales Contract",
    "Donation Contract",
    "Loan Contract",
    "Lease Contract",
    "Power of Attorney",
    "Will",
    "Certificate",
    "Other"
]

for i, text in enumerate(labels_text):
    tk.Label(form_frame, text=text, font=font_base, bg="#2c2f33", fg=text_color)\
        .grid(row=i, column=0, sticky="w", pady=4)

    if text == "Act Type":
        combo = ttk.Combobox(form_frame, values=act_types, font=font_base, width=57)
        combo.grid(row=i, column=1, pady=4, padx=10)
        combo.set("")  # default empty
        entries.append(combo)
    else:
        entry = tk.Entry(form_frame, width=60, font=font_base, bd=1, relief="solid",
                         bg="#2c2f33", fg="white", insertbackground="white")
        entry.grid(row=i, column=1, pady=4, padx=10)
        entries.append(entry)

tk.Button(form_frame, text="💾 Save", command=save_record,
          bg="#007BFF", fg="white", font=font_bold,
          relief="flat", padx=15, pady=5).grid(row=len(labels_text), column=0, columnspan=2, pady=15)

status_label = tk.Label(form_frame, text="", bg="#2c2f33", fg="lightgreen", font=font_base)
status_label.grid(row=len(labels_text)+1, column=0, columnspan=2)

# -------------- SEARCH --------------
search_frame = tk.Frame(app, bg="#2c2f33")
search_frame.pack(padx=20, anchor="nw")

tk.Label(search_frame, text="🔍 Search:", font=font_base, bg="#2c2f33", fg=text_color).pack(side="left")
search_entry = tk.Entry(search_frame, width=40, font=font_base, bd=1, relief="solid",
                         bg="#444", fg="white", insertbackground="white")
search_entry.pack(side="left", padx=10)

tk.Button(search_frame, text="Search", command=search_records,
          bg="#28a745", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

tk.Button(search_frame, text="🗑️ Delete", command=delete_record,
          bg="#dc3545", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

tk.Button(search_frame, text="📤 Export", command=export_to_excel,
          bg="#17a2b8", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

# -------------- TABLE WITH SCROLLBAR --------------
table_frame = tk.Frame(app, bg="#2c2f33")
table_frame.pack(padx=20, pady=20, fill="both", expand=True)

columns = (
    "ID", "Act Number", "Date", "Full Name", "Birth Date", "Personal ID",
    "Act Type", "State Fee", "Assistance Payment", "Notes"
)

scroll_y = tk.Scrollbar(table_frame, orient="vertical")
scroll_y.pack(side="right", fill="y")

table = ttk.Treeview(table_frame, columns=columns, show='headings', yscrollcommand=scroll_y.set)

style = ttk.Style()
style.theme_use("default")
style.configure("Treeview",
                background="#444",
                foreground="white",
                rowheight=25,
                fieldbackground="#444",
                font=font_base)
style.configure("Treeview.Heading",
                background="#343a40",
                foreground="white",
                font=font_bold)

for col in columns:
    table.heading(col, text=col)
    table.column(col, width=150, anchor="w")

table.pack(fill="both", expand=True)
scroll_y.config(command=table.yview)

display_records()
app.mainloop()
