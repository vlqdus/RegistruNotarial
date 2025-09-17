import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import openpyxl

# Bază de date
conn = sqlite3.connect("registru.db")
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS registru (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nr_act TEXT,
        data TEXT,
        nume_prenume TEXT,
        data_nasterii TEXT,
        idnp TEXT,
        denumire_act TEXT,
        taxa_stat TEXT,
        plata_asistenta TEXT,
        mentiuni TEXT
    )
''')
conn.commit()

# Funcții
def salveaza():
    valori = tuple(entry.get().strip() for entry in entries)
    try:
        cursor.execute('''
            INSERT INTO registru (nr_act, data, nume_prenume, data_nasterii, idnp,
                                  denumire_act, taxa_stat, plata_asistenta, mentiuni)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', valori)
        conn.commit()
        status_label.config(text="✅ Salvat!")
        afiseaza()

        #curăță câmpurile după salvare
        for entry in entries:
            if isinstance(entry, tk.Entry):
                entry.delete(0, tk.END)
            elif isinstance(entry, ttk.Combobox):
                entry.set("")

    except Exception as e:
        status_label.config(text=f"❌ Eroare: {e}")


def cauta():
    valoare = entry_cautare.get().strip()
    for row in tabel.get_children():
        tabel.delete(row)
    cursor.execute("""
        SELECT * FROM registru WHERE 
            nume_prenume LIKE ? OR 
            data_nasterii LIKE ? OR 
            idnp LIKE ? OR 
            nr_act LIKE ?
    """, (f"%{valoare}%",)*4)
    for row in cursor.fetchall():
        tabel.insert('', 'end', values=row)

def afiseaza():
    for row in tabel.get_children():
        tabel.delete(row)
    cursor.execute("SELECT * FROM registru")
    for row in cursor.fetchall():
        tabel.insert('', 'end', values=row)

def sterge():
    selected = tabel.selection()
    if not selected:
        messagebox.showwarning("Atenție", "Selectează o înregistrare de șters.")
        return
    id_sters = tabel.item(selected[0])['values'][0]
    cursor.execute("DELETE FROM registru WHERE id=?", (id_sters,))
    conn.commit()
    afiseaza()
    status_label.config(text="✅ Șters cu succes!")

def exporta_excel():
    cursor.execute("SELECT * FROM registru")
    rows = cursor.fetchall()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registru Notarial"
    ws.append(coloane)
    for row in rows:
        ws.append(row)
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        wb.save(filepath)
        messagebox.showinfo("Exportat", "Datele au fost exportate cu succes!")

# Aplicație
app = tk.Tk()
app.title("Registru Notarial")
app.geometry("1280x720")
app.configure(bg="#2c2f33")

# Font implicit
font_base = ("Segoe UI", 12)
font_bold = ("Segoe UI", 12, "bold")
text_color = "#ffffff"

# -------------- FORMULAR --------------
frame_formular = tk.Frame(app, bg="#2c2f33")
frame_formular.pack(padx=20, pady=20, anchor="nw")

labels_text = [
    "Nr. act", "Data", "Nume și prenume", "Data nașterii", "IDNP",
    "Denumire act", "Taxă de stat", "Plata asistență", "Mențiuni"
]
entries = []

contracte = [
    "Contract de vânzare-cumpărare",
    "Contract de donație",
    "Contract de împrumut",
    "Contract de închiriere",
    "Procura",
    "Testament",
    "Certificat",
    "Altele"
]

for i, text in enumerate(labels_text):
    tk.Label(frame_formular, text=text, font=font_base, bg="#2c2f33", fg=text_color)\
        .grid(row=i, column=0, sticky="w", pady=4)

    if text == "Denumire act":
        combo = ttk.Combobox(frame_formular, values=contracte, font=font_base, width=57)
        combo.grid(row=i, column=1, pady=4, padx=10)
        combo.set("")  # implicit gol
        entries.append(combo)
    else:
        entry = tk.Entry(frame_formular, width=60, font=font_base, bd=1, relief="solid",
                         bg="#2c2f33", fg="white", insertbackground="white")
        entry.grid(row=i, column=1, pady=4, padx=10)
        entries.append(entry)

tk.Button(frame_formular, text="💾 Salvează", command=salveaza,
          bg="#007BFF", fg="white", font=font_bold,
          relief="flat", padx=15, pady=5).grid(row=len(labels_text), column=0, columnspan=2, pady=15)

status_label = tk.Label(frame_formular, text="", bg="#2c2f33", fg="lightgreen", font=font_base)
status_label.grid(row=len(labels_text)+1, column=0, columnspan=2)

# -------------- CĂUTARE --------------
frame_cautare = tk.Frame(app, bg="#2c2f33")
frame_cautare.pack(padx=20, anchor="nw")

tk.Label(frame_cautare, text="🔍 Caută:", font=font_base, bg="#2c2f33", fg=text_color).pack(side="left")
entry_cautare = tk.Entry(frame_cautare, width=40, font=font_base, bd=1, relief="solid",
                         bg="#444", fg="white", insertbackground="white")
entry_cautare.pack(side="left", padx=10)

tk.Button(frame_cautare, text="Caută", command=cauta,
          bg="#28a745", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

tk.Button(frame_cautare, text="🗑️ Șterge", command=sterge,
          bg="#dc3545", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

tk.Button(frame_cautare, text="📤 Exportă", command=exporta_excel,
          bg="#17a2b8", fg="white", font=font_bold,
          relief="flat", padx=10, pady=3).pack(side="left", padx=10)

# -------------- TABEL CU SCROLL --------------
frame_tabel = tk.Frame(app, bg="#2c2f33")
frame_tabel.pack(padx=20, pady=20, fill="both", expand=True)

coloane = (
    "ID", "Nr. act", "Data", "Nume și prenume", "Data nașterii", "IDNP",
    "Denumire act", "Taxă de stat", "Plata asistență", "Mențiuni"
)

scroll_y = tk.Scrollbar(frame_tabel, orient="vertical")
scroll_y.pack(side="right", fill="y")

tabel = ttk.Treeview(frame_tabel, columns=coloane, show='headings', yscrollcommand=scroll_y.set)

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

for col in coloane:
    tabel.heading(col, text=col)
    tabel.column(col, width=150, anchor="w")

tabel.pack(fill="both", expand=True)
scroll_y.config(command=tabel.yview)

afiseaza()
app.mainloop()