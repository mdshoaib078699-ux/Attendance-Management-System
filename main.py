import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import mysql.connector
from datetime import date, datetime
from openpyxl import Workbook
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

DELETE_PASSWORD = "admin123"
LOGOUT_USERNAME = "shoaib"
LOGOUT_PASSWORD = "123"

def get_connection():
    return mysql.connector.connect(host="localhost", user="root", password="", database="attendance_management_db")

def ensure_contact_column():
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SHOW COLUMNS FROM students LIKE 'contact_no'")
        if not cur.fetchone():
            cur.execute("ALTER TABLE students ADD COLUMN contact_no VARCHAR(20) NULL")
            con.commit()
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def next_id(cur, table):
    if table not in ("students", "attendance"):
        raise ValueError("Invalid table")
    cur.execute(f"SELECT COALESCE(MAX(id),0)+1 FROM {table}")
    return cur.fetchone()[0]

def confirm_password():
    password = simpledialog.askstring("Password Required", "Enter delete password:", show="*", parent=root)
    if password is None:
        return False
    if password != DELETE_PASSWORD:
        messagebox.showerror("Error", "Wrong password")
        return False
    return True

# --- LOGIN WINDOW CODE ---
login_root = tk.Tk()
login_root.title("Login - Attendance Management System")
login_root.geometry("400x250")
login_root.minsize(400, 250)
login_root.resizable(False, False)
login_root.configure(bg="#edf2f7")

# Center the login window on screen
login_root.update_idletasks()
width = login_root.winfo_width()
height = login_root.winfo_height()
x = (login_root.winfo_screenwidth() // 2) - (width // 2)
y = (login_root.winfo_screenheight() // 2) - (height // 2)
login_root.geometry(f"{width}x{height}+{x}+{y}")

tk.Label(login_root, text="SYSTEM LOGIN", font=("Arial", 14, "bold"), bg="#1d3557", fg="white", pady=10).pack(fill="x")

login_frame = tk.Frame(login_root, bg="#edf2f7", padx=20, pady=20)
login_frame.pack(fill="both", expand=True)

tk.Label(login_frame, text="Username:", font=("Arial", 10, "bold"), bg="#edf2f7").pack(anchor="w", pady=(0, 2))
username_entry = tk.Entry(login_frame, font=("Arial", 10), width=28)
username_entry.pack(anchor="w", pady=(0, 10))
username_entry.focus()

tk.Label(login_frame, text="Password:", font=("Arial", 10, "bold"), bg="#edf2f7").pack(anchor="w", pady=(0, 2))
password_entry = tk.Entry(login_frame, font=("Arial", 10), show="*", width=28)
password_entry.pack(anchor="w", pady=(0, 15))

authenticated = False

def attempt_login(event=None):
    global authenticated
    u = username_entry.get().strip()
    p = password_entry.get()
    if u == LOGOUT_USERNAME and p == LOGOUT_PASSWORD:
        authenticated = True
        login_root.destroy()
    else:
        messagebox.showerror("Login Failed", "Invalid username or password.", parent=login_root)
        password_entry.delete(0, tk.END)

login_btn = tk.Button(login_frame, text="Login", command=attempt_login, bg="#2a9d8f", fg="white", font=("Arial", 10, "bold"), width=15)
login_btn.pack(anchor="center")

password_entry.bind("<Return>", attempt_login)
username_entry.bind("<Return>", lambda e: password_entry.focus())

login_root.mainloop()

# If user closed login window without authenticating, exit program completely
if not authenticated:
    exit()

# --- MAIN APPLICATION CODE ---
root = tk.Tk()
root.title("Attendance Management System")
root.geometry("1350x800")
root.minsize(1100, 650)
root.configure(bg="#edf2f7")
selected_id = None
active_tree = None
school_id_var = tk.StringVar()
name_var = tk.StringVar()
roll_var = tk.StringVar()
class_var = tk.StringVar()
contact_var = tk.StringVar()
status_var = tk.StringVar(value="Present")
date_var = tk.StringVar(value=str(date.today()))
attendance_list_date_var = tk.StringVar(value=str(date.today()))
attendance_list_mode = "today"
percentage_var = tk.StringVar(value="Select a student to view percentage")
selected_student_details_var = tk.StringVar(value="Select a student in Student Attendance Summary to view details")
registered_class_values = []
all_attendance_class_boxes = []
stats_var = tk.StringVar(value="Total Students: 0\nPresent: 0    Absent: 0\nAttendance Percentage: 0.00%")
class_filter_var = tk.StringVar(value="All Classes")
class_stats_var = tk.StringVar(value="Class Students: 0\nPresent: 0   Absent: 0")

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", rowheight=22, font=("Arial", 9))
style.configure("Treeview.Heading", background="#1d3557", foreground="white", font=("Arial", 9, "bold"))

canvas = tk.Canvas(root, bg="#edf2f7", highlightthickness=0)
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)
main_frame = tk.Frame(canvas, bg="#edf2f7")
canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")

def update_scroll(event=None):
    canvas.configure(scrollregion=canvas.bbox("all"))

def resize_frame(event):
    canvas.itemconfigure(canvas_window, width=event.width)

def mouse_scroll(event):
    canvas.yview_scroll(int(-event.delta/120), "units")

main_frame.bind("<Configure>", update_scroll)
canvas.bind("<Configure>", resize_frame)
canvas.bind_all("<MouseWheel>", mouse_scroll)

tk.Label(main_frame, text="ATTENDANCE MANAGEMENT SYSTEM", font=("Arial", 19, "bold"), bg="#1d3557", fg="white", pady=7).pack(fill="x", padx=8, pady=5)

top_frame = tk.Frame(main_frame, bg="#edf2f7")
top_frame.pack(fill="x", padx=8, pady=2)
top_frame.grid_columnconfigure(0, weight=3, uniform="top")
top_frame.grid_columnconfigure(1, weight=2, uniform="top")
form = tk.LabelFrame(top_frame, text="Student Details", font=("Arial", 10, "bold"), bg="white", padx=4, pady=2)
form.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
fields = [("School ID", school_id_var), ("Student Name", name_var), ("Roll No", roll_var), ("Class Name", class_var), ("Contact No", contact_var)]
for i, (label, var) in enumerate(fields):
    r, c = divmod(i, 2)
    tk.Label(form, text=label, bg="white").grid(row=r, column=c*2, padx=3, pady=1, sticky="w")
    tk.Entry(form, textvariable=var, width=16).grid(row=r, column=c*2+1, padx=3, pady=1, sticky="ew")
for i in range(4):
    form.grid_columnconfigure(i, weight=1)

student_action_frame = tk.Frame(form, bg="white")
student_action_frame.grid(row=3, column=0, columnspan=4, sticky="w", pady=(2, 0))
tk.Button(student_action_frame, text="Add Student", command=lambda: add_student(), bg="#2a9d8f", fg="white", padx=6, pady=3).pack(side="left", padx=2)
tk.Button(student_action_frame, text="Update Student", command=lambda: update_student(), bg="#457b9d", fg="white", padx=6, pady=3).pack(side="left", padx=2)
tk.Button(student_action_frame, text="Delete Student", command=lambda: delete_student(), bg="#e76f51", fg="white", padx=6, pady=3).pack(side="left", padx=2)
tk.Button(student_action_frame, text="Clear", command=lambda: clear_fields(), bg="#6c757d", fg="white", padx=6, pady=3).pack(side="left", padx=2)

percentage_panel = tk.LabelFrame(top_frame, text="Attendance Percentage", font=("Arial", 10, "bold"), bg="white", padx=5, pady=3)
percentage_panel.grid(row=0, column=1, sticky="nsew")
percentage_content = tk.Frame(percentage_panel, bg="white")
percentage_content.pack(fill="both", expand=True)
percentage_content.grid_columnconfigure(0, weight=1, uniform="percentage_info")
percentage_content.grid_columnconfigure(1, weight=1, uniform="percentage_info")
stats_column = tk.Frame(percentage_content, bg="white")
stats_column.grid(row=0, column=0, sticky="nsew", padx=(1, 4))
selected_details_column = tk.LabelFrame(percentage_content, text="Selected Student Details", bg="white", padx=4, pady=2)
selected_details_column.grid(row=0, column=1, sticky="nsew", padx=(4, 1))
tk.Label(stats_column, textvariable=stats_var, bg="white", fg="green", font=("Arial", 9, "bold"), justify="left", wraplength=230).pack(anchor="w", pady=1)
tk.Label(selected_details_column, textvariable=selected_student_details_var, bg="white", fg="#1d3557", font=("Arial", 9), justify="left", anchor="nw", wraplength=230).pack(fill="both", expand=True, anchor="w")

def open_calendar(target_var, parent=None):
    """Calendar popup: choose year first, then month, then a valid day."""
    parent = parent or root
    try:
        initial = datetime.strptime(target_var.get(), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        initial = date.today()

    import calendar as _calendar
    win = tk.Toplevel(parent)
    win.title("Select Attendance Date")
    win.transient(parent)
    win.grab_set()
    win.resizable(False, False)
    win.configure(bg="white")

    head = tk.Frame(win, bg="white", padx=10, pady=10)
    head.pack(fill="x")
    tk.Label(head, text="Year:", bg="white", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 4))
    year_var = tk.StringVar(value=str(initial.year))
    year_box = ttk.Combobox(head, textvariable=year_var, values=[str(y) for y in range(1900, 2101)], width=7)
    year_box.pack(side="left", padx=(0, 12))
    tk.Label(head, text="Month:", bg="white", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 4))
    month_var = tk.StringVar(value=_calendar.month_name[initial.month])
    month_box = ttk.Combobox(head, textvariable=month_var, values=list(_calendar.month_name)[1:], state="readonly", width=12)
    month_box.pack(side="left", padx=(0, 5))

    day_frame = tk.Frame(win, bg="white", padx=10, pady=5)
    day_frame.pack()
    selected_note = tk.Label(win, text="Choose a date from the calendar", bg="white", fg="#1d3557")
    selected_note.pack(pady=(0, 8))

    def draw_days(event=None):
        for child in day_frame.winfo_children():
            child.destroy()
        try:
            year = int(year_var.get())
            month = list(_calendar.month_name).index(month_var.get())
            if year < 1900 or year > 2100 or month < 1:
                raise ValueError
        except (ValueError, TypeError):
            tk.Label(day_frame, text="Choose a year from 1900 to 2100 and a month.", bg="white", fg="red").grid(row=0, column=0, columnspan=7, padx=6, pady=8)
            return

        for col, day_name in enumerate(("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")):
            tk.Label(day_frame, text=day_name, width=5, bg="#1d3557", fg="white", font=("Arial", 9, "bold")).grid(row=0, column=col, padx=1, pady=1)
        for row_index, week in enumerate(_calendar.monthcalendar(year, month), start=1):
            for col, number in enumerate(week):
                if number:
                    current = (year == initial.year and month == initial.month and number == initial.day)
                    tk.Button(
                        day_frame, text=str(number), width=4,
                        bg="#2a9d8f" if current else "#f1f5f9",
                        fg="white" if current else "#1d3557",
                        command=lambda day=number, yr=year, mo=month: choose_date(yr, mo, day)
                    ).grid(row=row_index, column=col, padx=1, pady=1)

    def choose_date(year, month, day):
        target_var.set(date(year, month, day).isoformat())
        win.destroy()
        if target_var is date_var:
            update_class_stats()
        elif target_var is attendance_list_date_var:
            view_attendance_date()

    year_box.bind("<<ComboboxSelected>>", draw_days)
    year_box.bind("<Return>", draw_days)
    year_box.bind("<FocusOut>", draw_days)
    month_box.bind("<<ComboboxSelected>>", draw_days)
    draw_days()

attendance_form = tk.Frame(main_frame, bg="#edf2f7")
attendance_form.pack(fill="x", padx=8, pady=2)
tk.Label(attendance_form, text="Attendance Date:", bg="#edf2f7").pack(side="left", padx=5)
date_entry = tk.Entry(attendance_form, textvariable=date_var, width=13)
date_entry.pack(side="left", padx=5)
tk.Button(attendance_form, text="📅 Calendar", command=lambda: open_calendar(date_var), bg="#457b9d", fg="white", padx=5).pack(side="left", padx=2)
date_entry.bind("<FocusOut>", lambda event: update_class_stats())
tk.Label(attendance_form, text="Status:", bg="#edf2f7").pack(side="left", padx=5)
ttk.Combobox(attendance_form, textvariable=status_var, values=("Present", "Absent", "Late", "Leave"), state="readonly", width=12).pack(side="left", padx=5)

def set_active_tree(tree):
    global active_tree
    active_tree = tree

def make_table(parent, columns, height=5):
    box = tk.Frame(parent, bg="white")
    box.pack(fill="both", expand=True, padx=2, pady=2)
    tree = ttk.Treeview(box, columns=columns, show="headings", height=height, selectmode="extended")
    vs = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
    hs = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    vs.pack(side="right", fill="y")
    hs.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, minwidth=70, anchor="center")
    tree.bind("<Button-1>", lambda e, t=tree: set_active_tree(t), add="+")
    tree.bind("<FocusIn>", lambda e, t=tree: set_active_tree(t), add="+")
    return tree

tree_rows_cache = {}
search_class_boxes = {}
search_bars = {}
search_class_vars = {}
search_text_vars = {}
search_roll_vars = {}
def cache_tree_rows(tree):
    tree_rows_cache[tree] = [tuple(tree.item(item, "values")) for item in tree.get_children()]
    if tree in search_class_boxes:
        search_class_boxes[tree]["values"] = ["All Classes"] + list(registered_class_values)

def search_tree(tree, search_var, class_var=None, roll_var=None):
    """Search a table by general text, class and roll number."""
    term = search_var.get().strip().casefold() if search_var else ""
    class_name = class_var.get().strip() if class_var else "All Classes"
    roll_number = roll_var.get().strip().casefold() if roll_var else ""
    rows = tree_rows_cache.get(
        tree, [tuple(tree.item(item, "values")) for item in tree.get_children()]
    )

    if tree == student_tree:
        roll_index, class_index = 3, 4
    elif tree == attendance_tree:
        roll_index, class_index = 2, 3
    else:
        roll_index, class_index = 3, 4

    tree.delete(*tree.get_children())
    for row in rows:
        if class_name and class_name != "All Classes":
            if str(row[class_index]).strip().casefold() != class_name.casefold():
                continue
        if roll_number and roll_number not in str(row[roll_index]).casefold():
            continue
        if term and not any(term in str(value).casefold() for value in row):
            continue
        tree.insert("", "end", values=row)

def add_search_bar(parent, tree, refresh_command):
    bar = tk.Frame(parent, bg="#edf2f7")
    search_bars[tree] = bar
    bar.pack(fill="x", padx=3, pady=2)
    var = tk.StringVar()
    roll_var = tk.StringVar()
    class_var = tk.StringVar(value="All Classes")
    search_text_vars[tree] = var
    search_roll_vars[tree] = roll_var
    search_class_vars[tree] = class_var

    class_index = 3 if tree == attendance_tree else 4
    initial_rows = tree_rows_cache.get(tree, [])
    classes = sorted({str(row[class_index]) for row in initial_rows if len(row) > class_index and str(row[class_index]).strip()})

    tk.Label(bar, text="Class:", bg="#edf2f7").pack(side="left", padx=(2, 2))
    class_box = ttk.Combobox(
        bar, textvariable=class_var, values=["All Classes"] + list(registered_class_values),
        state="readonly", width=14
    )
    class_box.pack(side="left", padx=2)
    search_class_boxes[tree] = class_box

    tk.Label(bar, text="Roll No:", bg="#edf2f7").pack(side="left", padx=(3, 2))
    roll_entry = tk.Entry(bar, textvariable=roll_var, width=10)
    roll_entry.pack(side="left", padx=2)

    if tree != attendance_tree:
        tk.Label(bar, text="Search:", bg="#edf2f7").pack(side="left", padx=(3, 2))
    entry = tk.Entry(bar, textvariable=var, width=15 if tree != attendance_tree else 18)
    entry.pack(side="left", padx=2, fill="x", expand=True)

    apply_filter = lambda: search_tree(tree, var, class_var, roll_var)
    class_box.bind("<<ComboboxSelected>>", lambda event: apply_filter())
    tk.Button(
        bar, text="Search", command=apply_filter,
        bg="#457b9d", fg="white", padx=5
    ).pack(side="left", padx=2)

    def show_all():
        if tree == attendance_tree:
            attendance_list_date_var.set(str(date.today()))
            var.set("")
            roll_var.set("")
            class_var.set("All Classes")
            fetch_attendance()
            search_tree(attendance_tree, var, class_var, roll_var)
            return
        for current_tree in (student_tree, summary_tree):
            if current_tree in search_text_vars:
                search_text_vars[current_tree].set("")
                search_roll_vars[current_tree].set("")
                search_class_vars[current_tree].set("All Classes")
        class_filter_var.set("All Classes")
        refresh_all()
        for current_tree in (student_tree, summary_tree):
            if current_tree in search_text_vars:
                search_tree(current_tree, search_text_vars[current_tree], search_class_vars[current_tree], search_roll_vars[current_tree])
        class_box["values"] = ["All Classes"] + list(registered_class_values)

    if tree != attendance_tree:
        tk.Button(
            bar, text="Show All", command=show_all,
            bg="#6c757d", fg="white", padx=5
        ).pack(side="left", padx=2)

    def update_class_values(event=None):
        rows = tree_rows_cache.get(tree, [])
        vals = sorted({str(row[class_index]) for row in rows if len(row) > class_index and str(row[class_index]).strip()})
        class_box["values"] = ["All Classes"] + list(registered_class_values)
        apply_filter()

    class_box.bind("<<ComboboxSelected>>", lambda event: apply_filter())
    entry.bind("<Return>", lambda event: apply_filter())
    roll_entry.bind("<Return>", lambda event: apply_filter())
    return var

summary_panel = tk.LabelFrame(main_frame, text="Student Attendance Summary", font=("Arial", 11, "bold"), bg="#edf2f7")
summary_toolbar = tk.Frame(summary_panel, bg="#edf2f7")
summary_toolbar.pack(fill="x", padx=3, pady=3)
tk.Button(summary_toolbar, text="Monthly Graph", command=lambda: monthly_graph(), bg="#6a4c93", fg="white", padx=8, pady=4).pack(side="left", padx=3)
summary_tree = make_table(summary_panel, ("ID", "School ID", "Student", "Roll No", "Class", "Present", "Absent", "Late", "Leave", "Percentage"), height=9)

def fetch_students():
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        student_tree.delete(*student_tree.get_children())
        selected_class = class_filter_var.get()
        if selected_class and selected_class != "All Classes":
            cur.execute("SELECT id,contact_no,student_name,roll_no,class_name FROM students WHERE class_name=%s ORDER BY id DESC", (selected_class,))
        else:
            cur.execute("SELECT id,contact_no,student_name,roll_no,class_name FROM students ORDER BY id DESC")
        for row in cur.fetchall():
            student_tree.insert("", "end", values=(row[0], row[1] or "", row[2], row[3], row[4]))
        cache_tree_rows(student_tree)
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def fetch_attendance():
    global attendance_list_mode
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        attendance_tree.delete(*attendance_tree.get_children())
        query = """SELECT a.id,s.student_name,s.roll_no,s.class_name,a.attendance_date,a.status
                   FROM attendance a JOIN students s ON s.id=a.student_id"""
        if attendance_list_mode == "previous":
            try:
                selected_date = datetime.strptime(attendance_list_date_var.get().strip(), "%Y-%m-%d").date()
            except ValueError:
                attendance_list_mode = "today"
                attendance_list_date_var.set(str(date.today()))
                selected_date = date.today()
            year_start = date(selected_date.year, 1, 1)
            query += " WHERE a.attendance_date >= %s AND a.attendance_date < %s"
            params = (year_start, selected_date)
        elif attendance_list_mode == "date":
            try:
                selected_date = datetime.strptime(attendance_list_date_var.get().strip(), "%Y-%m-%d").date()
            except ValueError:
                selected_date = date.today()
                attendance_list_date_var.set(str(selected_date))
            query += " WHERE a.attendance_date=%s"
            params = (selected_date,)
        else:
            attendance_list_mode = "today"
            attendance_list_date_var.set(str(date.today()))
            query += " WHERE a.attendance_date=%s"
            params = (date.today(),)
        query += " ORDER BY a.attendance_date DESC,a.id DESC"
        cur.execute(query, params)
        for row in cur.fetchall():
            attendance_tree.insert("", "end", values=row)
        cache_tree_rows(attendance_tree)
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def view_attendance_date():
    global attendance_list_mode
    try:
        datetime.strptime(attendance_list_date_var.get().strip(), "%Y-%m-%d")
    except ValueError:
        messagebox.showwarning("Date Format", "Select a valid date from the calendar.")
        return
    attendance_list_mode = "date"
    fetch_attendance()
    if attendance_tree in search_text_vars:
        search_tree(attendance_tree, search_text_vars[attendance_tree], search_class_vars[attendance_tree], search_roll_vars[attendance_tree])

def show_previous_attendance():
    global attendance_list_mode
    try:
        datetime.strptime(attendance_list_date_var.get().strip(), "%Y-%m-%d")
    except ValueError:
        messagebox.showwarning("Select Date", "Select a date from the calendar before viewing previous attendance.")
        return
    attendance_list_mode = "previous"
    if attendance_tree in search_text_vars:
        search_text_vars[attendance_tree].set("")
        search_roll_vars[attendance_tree].set("")
        search_class_vars[attendance_tree].set("All Classes")
    fetch_attendance()
    if attendance_tree in search_text_vars:
        search_tree(attendance_tree, search_text_vars[attendance_tree], search_class_vars[attendance_tree], search_roll_vars[attendance_tree])

def reset_attendance_list():
    """Return to today's attendance only; show an empty table when today has no records."""
    global attendance_list_mode
    attendance_list_mode = "today"
    attendance_list_date_var.set(str(date.today()))
    if attendance_tree in search_text_vars:
        search_text_vars[attendance_tree].set("")
        search_roll_vars[attendance_tree].set("")
        search_class_vars[attendance_tree].set("All Classes")
    fetch_attendance()
    if attendance_tree in search_text_vars:
        search_tree(attendance_tree, search_text_vars[attendance_tree], search_class_vars[attendance_tree], search_roll_vars[attendance_tree])

def fetch_summary():
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        summary_tree.delete(*summary_tree.get_children())
        cur.execute("""SELECT s.id,s.school_id,s.student_name,s.roll_no,s.class_name,
                       COALESCE(SUM(a.status='Present'),0),
                       COALESCE(SUM(a.status='Absent'),0),
                       COALESCE(SUM(a.status='Late'),0),
                       COALESCE(SUM(a.status='Leave'),0),
                       CASE WHEN COUNT(a.id)=0 THEN 0
                       ELSE ROUND(SUM(a.status='Present')*100.0/COUNT(a.id),2) END
                       FROM students s LEFT JOIN attendance a ON s.id=a.student_id
                       GROUP BY s.id,s.school_id,s.student_name,s.roll_no,s.class_name
                       ORDER BY s.id DESC""")
        for row in cur.fetchall():
            values = list(row)
            values[9] = f"{float(values[9]):.2f}%"
            summary_tree.insert("", "end", values=values)
        cache_tree_rows(summary_tree)
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def update_global_stats():
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT COUNT(*) FROM students")
        total_students = cur.fetchone()[0]
        cur.execute("""SELECT COUNT(*), COALESCE(SUM(status='Present'),0),
                       COALESCE(SUM(status='Absent'),0), COALESCE(SUM(status='Late'),0),
                       COALESCE(SUM(status='Leave'),0) FROM attendance""")
        total_marked, present, absent, late, leave = cur.fetchone()
        percentage = (present * 100 / total_marked) if total_marked else 0
        stats_var.set(
            f"Total Students: {total_students}\n"
            f"Total Attendance Marked: {total_marked}\n"
            f"Present: {present}    Absent: {absent}\n"
            f"Attendance Percentage: {percentage:.2f}%"
        )
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def update_class_stats():
    """Show student count and attendance counts for the selected class and attendance date."""
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        selected_class = class_filter_var.get()
        try:
            att_date = datetime.strptime(date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            class_stats_var.set("Enter date as YYYY-MM-DD")
            return

        if selected_class and selected_class != "All Classes":
            cur.execute("""SELECT COUNT(*) FROM students WHERE class_name=%s""", (selected_class,))
            total_students = cur.fetchone()[0]
            cur.execute("""SELECT
                COALESCE(SUM(a.status='Present'),0),
                COALESCE(SUM(a.status='Absent'),0)
                FROM students s LEFT JOIN attendance a
                ON a.student_id=s.id AND a.attendance_date=%s
                WHERE s.class_name=%s""", (att_date, selected_class))
        else:
            cur.execute("SELECT COUNT(*) FROM students")
            total_students = cur.fetchone()[0]
            cur.execute("""SELECT
                COALESCE(SUM(status='Present'),0),
                COALESCE(SUM(status='Absent'),0)
                FROM attendance WHERE attendance_date=%s""", (att_date,))
        present, absent = cur.fetchone()
        class_stats_var.set(
            f"Total Students: {total_students}\nPresent: {present}   Absent: {absent}\nDate: {att_date}"
        )
    except mysql.connector.Error as e:
        class_stats_var.set("Could not load class statistics")
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def refresh_all():
    try:
        load_class_options()
        fetch_students()
        fetch_attendance()
        fetch_summary()
        update_global_stats()
        update_class_stats()
    except mysql.connector.Error as e:
        messagebox.showerror("Database Error", str(e))

def clear_fields():
    global selected_id
    selected_id = None
    school_id_var.set("")
    name_var.set("")
    roll_var.set("")
    class_var.set("")
    contact_var.set("")
    percentage_var.set("Select a student to view percentage")
    student_tree.selection_remove(student_tree.selection())

def select_student(event=None):
    global selected_id
    selection = student_tree.selection()
    if not selection:
        return
    row = student_tree.item(selection[0], "values")
    selected_id = row[0]
    name_var.set(row[2])
    roll_var.set(row[3])
    class_var.set(row[4])
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT school_id,contact_no FROM students WHERE id=%s", (selected_id,))
        student_row = cur.fetchone()
        school_id_var.set((student_row[0] or "") if student_row else "")
        contact_var.set((student_row[1] or "") if student_row else "")
        cur.execute("""SELECT COUNT(*),COALESCE(SUM(status='Present'),0),
                       COALESCE(SUM(status='Absent'),0),COALESCE(SUM(status='Late'),0),
                       COALESCE(SUM(status='Leave'),0)
                       FROM attendance WHERE student_id=%s""", (selected_id,))
        total, present, absent, late, leave = cur.fetchone()
        percentage = present*100/total if total else 0
        selected_student_details_var.set(
            f"Student: {row[2]}\nSchool ID: {school_id_var.get()}  |  Roll No: {row[3]}\nClass: {row[4]}\n"
            f"Contact No: {contact_var.get()}\nMarked: {total}  |  Present: {present}  Absent: {absent}\n"
            f"Late: {late}  Leave: {leave}  |  Attendance: {percentage:.2f}%"
        )
    except mysql.connector.Error as e:
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def add_student():
    values = (school_id_var.get().strip(), name_var.get().strip(), roll_var.get().strip(), class_var.get().strip(), contact_var.get().strip())
    if not all(values[:4]):
        messagebox.showwarning("Warning", "Fill all student details")
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT id FROM students WHERE school_id=%s AND roll_no=%s AND class_name=%s", (values[0], values[2], values[3]))
        if cur.fetchone():
            messagebox.showwarning("Duplicate", "Student already exists")
            return
        sid = next_id(cur, "students")
        cur.execute("INSERT INTO students (id,school_id,student_name,roll_no,class_name,contact_no) VALUES (%s,%s,%s,%s,%s,%s)", (sid,)+values)
        con.commit()
        clear_fields()
        refresh_all()
        messagebox.showinfo("Success", "Student added")
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def update_student():
    if selected_id is None:
        messagebox.showwarning("Warning", "Select a student first")
        return
    values = (school_id_var.get().strip(), name_var.get().strip(), roll_var.get().strip(), class_var.get().strip(), contact_var.get().strip())
    if not all(values[:4]):
        messagebox.showwarning("Warning", "Fill all details")
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("UPDATE students SET school_id=%s,student_name=%s,roll_no=%s,class_name=%s,contact_no=%s WHERE id=%s", values+(selected_id,))
        con.commit()
        clear_fields()
        refresh_all()
        messagebox.showinfo("Success", "Student updated")
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def delete_student():
    if selected_id is None:
        messagebox.showwarning("Warning", "Select a student first")
        return
    if not confirm_password() or not messagebox.askyesno("Confirm", "Delete student and attendance?"):
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("DELETE FROM attendance WHERE student_id=%s", (selected_id,))
        cur.execute("DELETE FROM students WHERE id=%s", (selected_id,))
        con.commit()
        clear_fields()
        refresh_all()
        messagebox.showinfo("Success", "Student deleted")
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def save_attendance(student_id, att_date, status, cur):
    cur.execute("SELECT id FROM attendance WHERE student_id=%s AND attendance_date=%s", (student_id, att_date))
    existing = cur.fetchone()
    if existing:
        cur.execute("UPDATE attendance SET status=%s WHERE id=%s", (status, existing[0]))
    else:
        aid = next_id(cur, "attendance")
        cur.execute("INSERT INTO attendance (id,student_id,attendance_date,status) VALUES (%s,%s,%s,%s)", (aid, student_id, att_date, status))

def mark_attendance():
    mark_all_attendance()

def mark_all_attendance():
    """Confirm counts, then mark selected students Present and others in scope Absent."""
    try:
        att_date = datetime.strptime(date_var.get().strip(), "%Y-%m-%d").date()
    except ValueError:
        messagebox.showerror("Invalid Date", "Enter date as YYYY-MM-DD")
        return

    selected_items = student_tree.selection()
    if not selected_items:
        messagebox.showwarning(
            "Select Students",
            "Select the students to mark Present in the Student List. "
            "Press Ctrl+A to select all students in the current class."
        )
        return

    selected_ids = {str(student_tree.item(item, "values")[0]) for item in selected_items}
    selected_class = class_filter_var.get()
    scope = f"Class: {selected_class}" if selected_class and selected_class != "All Classes" else "Class: All Classes"

    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        if selected_class and selected_class != "All Classes":
            cur.execute("SELECT id FROM students WHERE class_name=%s", (selected_class,))
        else:
            cur.execute("SELECT id FROM students")
        all_ids = [str(row[0]) for row in cur.fetchall()]

        selected_ids.intersection_update(all_ids)
        total_students = len(all_ids)
        present_count = len(selected_ids)
        absent_count = total_students - present_count

        if total_students == 0:
            messagebox.showwarning("No Students", "There are no students in the selected class.")
            return

        msg = (
            f"{scope}\nDate: {att_date}\n\n"
            f"Total Students: {total_students}\n"
            f"Present: {present_count}\n"
            f"Absent: {absent_count}\n\n"
            f"Do you want to save attendance?"
        )
        if not messagebox.askyesno("Confirm Attendance", msg):
            return

        for sid in all_ids:
            status = "Present" if sid in selected_ids else "Absent"
            save_attendance(sid, att_date, status, cur)
        con.commit()
        refresh_all()
        messagebox.showinfo(
            "Success",
            f"Attendance saved for {att_date}.\nPresent: {present_count}\nAbsent: {absent_count}"
        )
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def delete_selected_all():
    if active_tree not in (student_tree, attendance_tree):
        messagebox.showwarning("Warning", "Select rows in Student List or Attendance List")
        return
    selected = active_tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Select rows first")
        return
    if not confirm_password() or not messagebox.askyesno("Confirm Delete", "Delete selected records?"):
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        for item in selected:
            row = active_tree.item(item, "values")
            if active_tree == student_tree:
                cur.execute("DELETE FROM attendance WHERE student_id=%s", (row[0],))
                cur.execute("DELETE FROM students WHERE id=%s", (row[0],))
            else:
                cur.execute("DELETE FROM attendance WHERE id=%s", (row[0],))
        con.commit()
        clear_fields()
        refresh_all()
        messagebox.showinfo("Success", "Selected records deleted")
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def delete_entire_system():
    if not confirm_password():
        return
    if not messagebox.askyesno("Confirm Delete", "Delete ALL students and attendance records?"):
        return
    if not messagebox.askyesno("Final Confirmation", "This cannot be undone. Continue?"):
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("DELETE FROM attendance")
        cur.execute("DELETE FROM students")
        con.commit()
        clear_fields()
        refresh_all()
        messagebox.showinfo("Success", "All records deleted")
    except mysql.connector.Error as e:
        if con:
            con.rollback()
        messagebox.showerror("Database Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def export_selected_student():
    selection = student_tree.selection()
    if not selection:
        messagebox.showwarning("Warning", "Select a student from Student List")
        return
    row = student_tree.item(selection[0], "values")
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Student_Attendance.xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
    if not path:
        return
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT id,school_id,student_name,roll_no,class_name,contact_no FROM students WHERE id=%s", (row[0],))
        details = cur.fetchone()
        cur.execute("SELECT attendance_date,status FROM attendance WHERE student_id=%s ORDER BY attendance_date DESC", (row[0],))
        records = cur.fetchall()
        cur.execute("SELECT COUNT(*),COALESCE(SUM(status='Present'),0) FROM attendance WHERE student_id=%s", (row[0],))
        total, present = cur.fetchone()
        wb = Workbook()
        ws = wb.active
        ws.title = "Student Details"
        ws.append(["ID", "School ID", "Student Name", "Roll No", "Class", "Contact No"])
        ws.append(list(details))
        ws.append([])
        ws.append(["Total Marked", total, "Present", present, "Percentage", f"{present*100/total:.2f}%" if total else "0.00%"])
        wa = wb.create_sheet("Attendance")
        wa.append(["Attendance Date", "Status"])
        for record in records:
            wa.append([str(record[0]), record[1]])
        wb.save(path)
        messagebox.showinfo("Success", "Student details and attendance exported")
    except (mysql.connector.Error, OSError) as e:
        messagebox.showerror("Export Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def export_summary():
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Attendance_Summary.xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
    if not path:
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["ID", "School ID", "Student", "Roll No", "Class", "Present", "Absent", "Late", "Leave", "Percentage"])
    for item in summary_tree.get_children():
        ws.append(list(summary_tree.item(item, "values")))
    try:
        wb.save(path)
        messagebox.showinfo("Success", "Summary exported")
    except OSError as e:
        messagebox.showerror("Export Error", str(e))

def edit_summary_attendance():
    """Edit an individual student's saved attendance from the summary list."""
    selection = attendance_tree.selection()
    if not selection:
        messagebox.showwarning(
            "Select Attendance",
            "Select an attendance row in the Attendance List first."
        )
        return

    values = attendance_tree.item(selection[0], "values")
    attendance_id = values[0]
    student_name = values[1]
    student_roll = values[2]
    student_class = values[3]
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT student_id FROM attendance WHERE id=%s", (attendance_id,))
        student_row = cur.fetchone()
        if not student_row:
            messagebox.showerror("Not Found", "The selected attendance record was not found.")
            return
        student_id = student_row[0]
    except mysql.connector.Error as e:
        messagebox.showerror("Database Error", str(e))
        return
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

    win = tk.Toplevel(root)
    win.title("Edit Student Attendance")
    win.geometry("850x500")
    win.minsize(700, 380)
    win.configure(bg="#edf2f7")

    tk.Label(
        win,
        text=f"Edit Attendance: {student_name}  |  Roll No: {student_roll}  |  Class: {student_class}",
        font=("Arial", 12, "bold"), bg="#1d3557", fg="white", pady=8
    ).pack(fill="x", padx=8, pady=6)

    tk.Label(
        win,
        text="Select an attendance record. Status changes are saved automatically.",
        bg="#edf2f7", wraplength=800
    ).pack(anchor="w", padx=10, pady=4)

    table_frame = tk.Frame(win, bg="white")
    table_frame.pack(fill="both", expand=True, padx=8, pady=5)
    cols = ("Attendance ID", "Date", "Status")
    edit_tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=12)
    edit_vbar = ttk.Scrollbar(table_frame, orient="vertical", command=edit_tree.yview)
    edit_tree.configure(yscrollcommand=edit_vbar.set)
    edit_vbar.pack(side="right", fill="y")
    edit_tree.pack(side="left", fill="both", expand=True)
    for col in cols:
        edit_tree.heading(col, text=col)
        edit_tree.column(col, width=180, anchor="center")

    editor = tk.Frame(win, bg="#edf2f7")
    editor.pack(fill="x", padx=8, pady=6)
    tk.Label(editor, text="Selected Status:", bg="#edf2f7").pack(side="left", padx=5)
    edit_status_var = tk.StringVar(value="Present")
    status_box = ttk.Combobox(
        editor, textvariable=edit_status_var,
        values=("Present", "Absent", "Late", "Leave"),
        state="readonly", width=14
    )
    status_box.pack(side="left", padx=5)

    def load_student_records():
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            cur.execute(
                "SELECT id, attendance_date, status FROM attendance WHERE student_id=%s ORDER BY attendance_date DESC, id DESC",
                (student_id,)
            )
            edit_tree.delete(*edit_tree.get_children())
            for record in cur.fetchall():
                edit_tree.insert("", "end", values=(record[0], str(record[1]), record[2]))
        except mysql.connector.Error as e:
            messagebox.showerror("Database Error", str(e), parent=win)
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()

    def pick_record(event=None):
        chosen = edit_tree.selection()
        if chosen:
            edit_status_var.set(edit_tree.item(chosen[0], "values")[2])

    def save_record_status(event=None):
        chosen = edit_tree.selection()
        if not chosen:
            messagebox.showwarning("Select Record", "Select an attendance record first.", parent=win)
            return
        record = edit_tree.item(chosen[0], "values")
        record_id = record[0]
        new_status = edit_status_var.get()
        if new_status not in ("Present", "Absent", "Late", "Leave"):
            messagebox.showwarning("Status", "Select a valid status.", parent=win)
            return
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            cur.execute(
                "UPDATE attendance SET status=%s WHERE id=%s AND student_id=%s",
                (new_status, record_id, student_id)
            )
            con.commit()
            load_student_records()
            refresh_all()
        except mysql.connector.Error as e:
            if con:
                con.rollback()
            messagebox.showerror("Database Error", str(e), parent=win)
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()

    edit_tree.bind("<<TreeviewSelect>>", pick_record)
    status_box.bind("<<ComboboxSelected>>", save_record_status)
    tk.Button(
        editor, text="Close", command=win.destroy,
        bg="#6c757d", fg="white", padx=10, pady=4
    ).pack(side="left", padx=4)
    load_student_records()

def monthly_graph():
    """Show monthly Present/Absent graph and totals for the selected summary student."""
    selected_student_id = None
    selected_student_name = None

    summary_selection = summary_tree.selection()
    if summary_selection:
        summary_values = summary_tree.item(summary_selection[0], "values")
        selected_student_id = str(summary_values[0])
        selected_student_name = str(summary_values[2])
    else:
        selected_items = student_tree.selection()
        if len(selected_items) == 1:
            student_values = student_tree.item(selected_items[0], "values")
            selected_student_id = str(student_values[0])
            selected_student_name = str(student_values[2])

    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        year = datetime.now().year
        selected_class = class_filter_var.get()
        conditions = ["YEAR(a.attendance_date)=%s"]
        params = [year]
        if selected_student_id:
            conditions.append("a.student_id=%s")
            params.append(selected_student_id)
            graph_title = f"{selected_student_name} (ID: {selected_student_id})"
        elif selected_class and selected_class != "All Classes":
            conditions.append("s.class_name=%s")
            params.append(selected_class)
            graph_title = f"Class {selected_class}"
        else:
            graph_title = "All Students"

        cur.execute(f"""SELECT MONTH(a.attendance_date),
                             COALESCE(SUM(a.status='Present'),0),
                             COALESCE(SUM(a.status='Absent'),0),
                             COALESCE(SUM(a.status='Late'),0),
                             COALESCE(SUM(a.status='Leave'),0)
                      FROM attendance a JOIN students s ON s.id=a.student_id
                      WHERE {' AND '.join(conditions)}
                      GROUP BY MONTH(a.attendance_date)
                      ORDER BY MONTH(a.attendance_date)""", tuple(params))
        rows = cur.fetchall()
        month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        present_values, absent_values = [], []
        late_values, leave_values = [], []
        row_by_month = {int(row[0]): row for row in rows}
        for month in range(1, 13):
            row = row_by_month.get(month, (month, 0, 0, 0, 0))
            present_values.append(int(row[1] or 0))
            absent_values.append(int(row[2] or 0))
            late_values.append(int(row[3] or 0))
            leave_values.append(int(row[4] or 0))

        total_present = sum(present_values)
        total_absent = sum(absent_values)
        total_late = sum(late_values)
        total_leave = sum(leave_values)
        if selected_student_id:
            totals_text = (f"{graph_title} | {year} totals: Present {total_present}, "
                           f"Absent {total_absent}, Late {total_late}, Leave {total_leave}")
        else:
            totals_text = (f"{graph_title} | {year} totals: Present {total_present}, "
                           f"Absent {total_absent}, Late {total_late}, Leave {total_leave}")

        x = list(range(12))
        width = 0.36
        plt.figure(figsize=(11, 5.5))
        plt.bar([i - width / 2 for i in x], present_values, width=width, label="Present")
        plt.bar([i + width / 2 for i in x], absent_values, width=width, label="Absent")
        plt.xticks(x, month_names)
        plt.xlabel("Month")
        plt.ylabel("Attendance Days / Records")
        plt.title(f"Monthly Attendance - {graph_title} ({year})\n{totals_text}")
        plt.legend()
        plt.grid(axis="y", linestyle="--", alpha=0.4)
        plt.tight_layout()
        plt.show()
    except mysql.connector.Error as e:
        messagebox.showerror("Graph Error", str(e))
    except Exception as e:
        messagebox.showerror("Graph Error", str(e))
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def open_all_attendance_window():
    """Display attendance by calendar date and/or class, with a Show All reset."""
    win = tk.Toplevel(root)
    win.title("View All Attendance")
    win.geometry("1200x650")
    win.minsize(900, 480)
    win.configure(bg="#edf2f7")
    tk.Label(win, text="ALL ATTENDANCE RECORDS", font=("Arial", 15, "bold"),
             bg="#1d3557", fg="white", pady=7).pack(fill="x", padx=8, pady=6)

    filters = tk.LabelFrame(win, text="Calendar and Class Filter", bg="#edf2f7", padx=6, pady=5)
    filters.pack(fill="x", padx=8, pady=4)
    row = tk.Frame(filters, bg="#edf2f7")
    row.pack(fill="x", pady=3)
    tk.Label(row, text="Attendance Date:", bg="#edf2f7").pack(side="left", padx=3)
    day_var = tk.StringVar()
    tk.Entry(row, textvariable=day_var, width=13).pack(side="left", padx=3)
    tk.Button(row, text="📅 Calendar", command=lambda: open_calendar(day_var, win),
              bg="#457b9d", fg="white", padx=5).pack(side="left", padx=3)

    tk.Label(row, text="Class:", bg="#edf2f7").pack(side="left", padx=(12, 3))
    class_var_filter = tk.StringVar(value="All Classes")
    class_box = ttk.Combobox(row, textvariable=class_var_filter,
                             values=["All Classes"] + list(registered_class_values),
                             state="readonly", width=14)
    class_box.pack(side="left", padx=3)
    all_attendance_class_boxes.append(class_box)
    count_var = tk.StringVar(value="Records: 0")
    tk.Label(row, textvariable=count_var, bg="#edf2f7", fg="#1d3557",
             font=("Arial", 9, "bold")).pack(side="right", padx=8)

    columns = ("ID", "Student", "Roll No", "Class", "Date", "Status")
    box = tk.Frame(win, bg="white")
    box.pack(fill="both", expand=True, padx=8, pady=5)
    tree = ttk.Treeview(box, columns=columns, show="headings", height=18)
    vbar = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
    hbar = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
    vbar.pack(side="right", fill="y")
    hbar.pack(side="bottom", fill="x")
    tree.pack(side="left", fill="both", expand=True)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=145, minwidth=80, anchor="center")

    def query_records(show_everything=False):
        day = day_var.get().strip()
        if day and not show_everything:
            try:
                datetime.strptime(day, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Date Format", "Select a valid date from the calendar.", parent=win)
                return
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            query = ("SELECT a.id,s.student_name,s.roll_no,s.class_name,a.attendance_date,a.status "
                     "FROM attendance a JOIN students s ON s.id=a.student_id")
            conditions, params = [], []
            if not show_everything:
                if day:
                    conditions.append("a.attendance_date=%s")
                    params.append(day)
                if class_var_filter.get() != "All Classes":
                    conditions.append("s.class_name=%s")
                    params.append(class_var_filter.get())
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            query += " ORDER BY a.attendance_date DESC,a.id DESC"
            cur.execute(query, tuple(params))
            rows = cur.fetchall()
            tree.delete(*tree.get_children())
            for record in rows:
                tree.insert("", "end", values=tuple(str(v) if hasattr(v, "isoformat") else v for v in record))
            count_var.set(f"Records: {len(rows)}")
        except mysql.connector.Error as e:
            messagebox.showerror("Database Error", str(e), parent=win)
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()

    tk.Button(row, text="Apply Filter", command=lambda: query_records(False),
              bg="#2a9d8f", fg="white", padx=7).pack(side="left", padx=5)
    tk.Button(row, text="Show All", command=lambda: reset_and_show_all(),
              bg="#457b9d", fg="white", padx=7).pack(side="left", padx=3)

    def reset_and_show_all():
        day_var.set("")
        class_var_filter.set("All Classes")
        query_records(True)

    class_box.bind("<<ComboboxSelected>>", lambda event: query_records(False))
    query_records(True)

def select_all(event=None):
    global active_tree
    widget = root.focus_get()
    if widget in (student_tree, attendance_tree, summary_tree):
        active_tree = widget
    if active_tree in (student_tree, attendance_tree, summary_tree):
        active_tree.selection_set(active_tree.get_children())
    return "break"

def show_summary_student(event=None):
    """Keep the selected student's concise information in the existing compact panel."""
    global selected_id
    selection = summary_tree.selection()
    if not selection:
        return
    selected_rows = [summary_tree.item(item, "values") for item in selection]

    if len(selected_rows) == 1:
        values = selected_rows[0]
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            cur.execute(
                "SELECT id,school_id,student_name,roll_no,class_name,contact_no FROM students WHERE id=%s",
                (values[0],)
            )
            student = cur.fetchone()
            if student:
                selected_id = student[0]
                school_id_var.set(student[1])
                name_var.set(student[2])
                roll_var.set(student[3])
                class_var.set(student[4])
                contact_var.set(student[5] or "")
                total = sum(int(values[i] or 0) for i in range(5, 9))
                selected_student_details_var.set(
                    f"Student: {student[2]}\nSchool ID: {student[1]}  |  Roll No: {student[3]}\n"
                    f"Class: {student[4]}\nContact No: {student[5] or ''}\nMarked: {total}  |  Present: {values[5]}  Absent: {values[6]}\n"
                    f"Late: {values[7]}  Leave: {values[8]}  |  Attendance: {values[9]}"
                )
        except mysql.connector.Error as e:
            messagebox.showerror("Database Error", str(e))
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()
    else:
        selected_id = None
        total_present = sum(int(row[5] or 0) for row in selected_rows)
        total_absent = sum(int(row[6] or 0) for row in selected_rows)
        total_late = sum(int(row[7] or 0) for row in selected_rows)
        total_leave = sum(int(row[8] or 0) for row in selected_rows)
        marked = total_present + total_absent + total_late + total_leave
        percentage = total_present * 100 / marked if marked else 0
        names = ", ".join(str(row[2]) for row in selected_rows[:3])
        if len(selected_rows) > 3:
            names += f" +{len(selected_rows) - 3} more"
        selected_student_details_var.set(
            f"Selected Students: {len(selected_rows)}\n{names}\n"
            f"Present: {total_present}  Absent: {total_absent}\n"
            f"Late: {total_late}  Leave: {total_leave}  |  Attendance: {percentage:.2f}%"
        )

_active_summary_report = None
_summary_report_after_id = None

def show_selected_summary_report(event=None):
    """Open one report window for the selected student(s), or refresh the existing window."""
    global _active_summary_report
    selection = summary_tree.selection()
    if not selection:
        return "break" if event else None

    if _active_summary_report is not None and _active_summary_report.winfo_exists():
        refresh_callback = getattr(_active_summary_report, "_refresh_selection", None)
        if refresh_callback:
            refresh_callback()
        return "break" if event else None

    summary_rows = [summary_tree.item(item, "values") for item in selection]
    student_ids = [int(row[0]) for row in summary_rows]
    report_state = {"summary_rows": summary_rows, "student_ids": student_ids}
    initial_report_selection = {"summary_rows": list(summary_rows), "student_ids": list(student_ids)}
    selected_year_var = tk.StringVar(value=str(date.today().year))
    selected_month_var = tk.StringVar(value="All Months")

    win = tk.Toplevel(root)
    _active_summary_report = win
    win.title("Selected Student Attendance")
    win.geometry("1400x900")
    win.minsize(1050, 700)
    win.configure(bg="#edf2f7")
    tk.Label(
        win, text="STUDENT DETAILS  |  MONTHLY ATTENDANCE  |  ATTENDANCE RECORDS",
        font=("Arial", 12, "bold"), bg="#1d3557", fg="white", pady=5
    ).pack(fill="x", padx=5, pady=3)

    report_search_bar = tk.Frame(win, bg="#edf2f7")
    report_search_bar.pack(fill="x", padx=7, pady=(4, 2))
    tk.Label(report_search_bar, text="Search Student:", bg="#edf2f7").pack(side="left", padx=(2, 5))
    report_search_var = tk.StringVar()
    report_search_entry = tk.Entry(report_search_bar, textvariable=report_search_var, width=32)
    report_search_entry.pack(side="left", padx=2)
    report_search_entry.bind("<Return>", lambda event: search_report_students())
    tk.Button(
        report_search_bar, text="Search", command=lambda: search_report_students(),
        bg="#457b9d", fg="white", padx=8
    ).pack(side="left", padx=3)
    tk.Button(
        report_search_bar, text="Show Selected", command=lambda: show_all_report_students(),
        bg="#6c757d", fg="white", padx=8
    ).pack(side="left", padx=3)

    top_report = tk.Frame(win, bg="#edf2f7")
    top_report.pack(fill="x", padx=5, pady=2)
    top_report.grid_columnconfigure(0, minsize=250, weight=0)
    top_report.grid_columnconfigure(1, weight=1)

    details_panel = tk.LabelFrame(
        top_report, text="Student Details", bg="white",
        font=("Arial", 9, "bold"), padx=5, pady=3, width=250, height=165
    )
    details_panel.grid(row=0, column=0, sticky="nw", padx=(0, 3))
    details_panel.grid_propagate(False)
    details_text = tk.Text(details_panel, height=8, width=27, wrap="word", bg="white",
                           relief="flat", font=("Arial", 9), state="normal")
    details_text.pack(fill="both", expand=True)
    details_text.insert("1.0", "\n\n".join(
        f"Student: {row[2]}\nSchool ID: {row[1]}\nRoll No: {row[3]}\nClass: {row[4]}\n"
        f"Present: {row[5]}  Absent: {row[6]}\nLate: {row[7]}  Leave: {row[8]}\nAttendance: {row[9]}"
        for row in summary_rows
    ))
    details_text.configure(state="disabled")

    graph_panel = tk.LabelFrame(
        top_report, text="Monthly Attendance Graph", bg="white",
        font=("Arial", 9, "bold"), padx=4, pady=3, height=365
    )
    graph_panel.grid(row=0, column=1, sticky="nsew", padx=(3, 0))
    graph_panel.grid_propagate(False)
    graph_controls = tk.Frame(graph_panel, bg="white")
    graph_controls.pack(fill="x", padx=3, pady=1)
    tk.Label(graph_controls, text="Year:", bg="white").pack(side="left", padx=(2, 4))
    year_box = ttk.Combobox(
        graph_controls, textvariable=selected_year_var,
        values=[str(y) for y in range(1900, 2101)], width=8, state="readonly"
    )
    year_box.pack(side="left", padx=2)
    tk.Label(graph_controls, text="Month:", bg="white").pack(side="left", padx=(8, 3))
    month_names = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    month_box = ttk.Combobox(
        graph_controls, textvariable=selected_month_var,
        values=["All Months"] + month_names, width=12, state="readonly"
    )
    month_box.pack(side="left", padx=2)

    fig = Figure(figsize=(10.5, 3.2), dpi=95)
    ax = fig.add_subplot(111)
    graph_canvas = FigureCanvasTkAgg(fig, master=graph_panel)
    graph_canvas.get_tk_widget().pack(fill="both", expand=True, padx=3, pady=1)

    attendance_panel_report = tk.LabelFrame(
        win, text="Attendance List",
        bg="#edf2f7", font=("Arial", 9, "bold"), padx=3, pady=2
    )
    attendance_panel_report.pack(fill="both", expand=True, padx=5, pady=2)
    att_cols = ("Attendance ID", "Student", "Roll No", "Class", "Date", "Status")
    table_frame = tk.Frame(attendance_panel_report, bg="white")
    table_frame.pack(fill="both", expand=True, padx=2, pady=2)
    att_tree = ttk.Treeview(table_frame, columns=att_cols, show="headings", height=18, selectmode="browse")
    att_vbar = ttk.Scrollbar(table_frame, orient="vertical", command=att_tree.yview)
    att_hbar = ttk.Scrollbar(table_frame, orient="horizontal", command=att_tree.xview)
    att_tree.configure(yscrollcommand=att_vbar.set, xscrollcommand=att_hbar.set)
    att_vbar.pack(side="right", fill="y")
    att_hbar.pack(side="bottom", fill="x")
    att_tree.pack(side="left", fill="both", expand=True)
    for col in att_cols:
        att_tree.heading(col, text=col)
        att_tree.column(col, width=125, minwidth=75, anchor="center")

    def load_attendance_rows(select_record_id=None):
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            placeholders = ",".join(["%s"] * len(report_state["student_ids"]))
            cur.execute(
                f"""SELECT a.id,s.student_name,s.roll_no,s.class_name,a.attendance_date,a.status
                    FROM attendance a JOIN students s ON s.id=a.student_id
                    WHERE a.student_id IN ({placeholders})
                    ORDER BY a.attendance_date DESC,a.id DESC""",
                tuple(report_state["student_ids"])
            )
            rows = cur.fetchall()
            att_tree.delete(*att_tree.get_children())
            for record in rows:
                values = tuple(str(v) if hasattr(v, "isoformat") else v for v in record)
                att_tree.insert("", "end", iid=str(record[0]), values=values)
            if select_record_id and att_tree.exists(str(select_record_id)):
                att_tree.selection_set(str(select_record_id))
                att_tree.focus(str(select_record_id))
        except mysql.connector.Error as e:
            messagebox.showerror("Database Error", str(e), parent=win)
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()

    def draw_monthly_graph(event=None):
        try:
            year = int(selected_year_var.get())
        except ValueError:
            year = date.today().year
            selected_year_var.set(str(year))
        month_choice = selected_month_var.get()
        con = cur = None
        try:
            con = get_connection()
            cur = con.cursor()
            placeholders = ",".join(["%s"] * len(report_state["student_ids"]))
            if month_choice == "All Months":
                cur.execute(
                    f"""SELECT MONTH(attendance_date),
                               COALESCE(SUM(status='Present'),0),
                               COALESCE(SUM(status='Absent'),0),
                               COALESCE(SUM(status='Late'),0),
                               COALESCE(SUM(status='Leave'),0)
                        FROM attendance
                        WHERE student_id IN ({placeholders}) AND YEAR(attendance_date)=%s
                        GROUP BY MONTH(attendance_date) ORDER BY MONTH(attendance_date)""",
                    tuple(report_state["student_ids"]) + (year,)
                )
                rows = cur.fetchall()
                by_period = {int(row[0]): row[1:] for row in rows}
                labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                positions = list(range(1, 13))
                counts = [by_period.get(m, (0, 0, 0, 0)) for m in positions]
                chart_title = "Monthly Attendance"
            else:
                month_number = month_names.index(month_choice) + 1
                cur.execute(
                    f"""SELECT DAY(attendance_date),
                               COALESCE(SUM(status='Present'),0),
                               COALESCE(SUM(status='Absent'),0),
                               COALESCE(SUM(status='Late'),0),
                               COALESCE(SUM(status='Leave'),0)
                        FROM attendance
                        WHERE student_id IN ({placeholders}) AND YEAR(attendance_date)=%s
                          AND MONTH(attendance_date)=%s
                        GROUP BY DAY(attendance_date) ORDER BY DAY(attendance_date)""",
                    tuple(report_state["student_ids"]) + (year, month_number)
                )
                rows = cur.fetchall()
                by_period = {int(row[0]): row[1:] for row in rows}
                import calendar as _calendar
                days_in_month = _calendar.monthrange(year, month_number)[1]
                positions = list(range(1, days_in_month + 1))
                labels = [str(day) for day in positions]
                counts = [by_period.get(day, (0, 0, 0, 0)) for day in positions]
                chart_title = f"Daily Attendance — {month_choice}"

            present = [int(row[0] or 0) for row in counts]
            absent = [int(row[1] or 0) for row in counts]
            late = [int(row[2] or 0) for row in counts]
            leave = [int(row[3] or 0) for row in counts]
            x = list(range(len(positions)))
            width = min(0.19, 0.75 / 4)
            ax.clear()
            ax.bar([i - 1.5 * width for i in x], present, width=width, label="Present")
            ax.bar([i - 0.5 * width for i in x], absent, width=width, label="Absent")
            ax.bar([i + 0.5 * width for i in x], late, width=width, label="Late")
            ax.bar([i + 1.5 * width for i in x], leave, width=width, label="Leave")
            ax.set_xticks(x, labels)
            ax.set_ylabel("Records")
            names = ", ".join(str(r[2]) for r in report_state["summary_rows"][:3])
            if len(report_state["summary_rows"]) > 3:
                names += f" +{len(report_state['summary_rows'])-3} more"
            ax.set_title(f"{chart_title} — {names} ({year})", fontsize=8)
            ax.legend(ncol=4, loc="upper right", fontsize=6)
            ax.tick_params(axis="both", labelsize=7)
            ax.grid(axis="y", linestyle="--", alpha=0.3)
            fig.tight_layout(pad=1.0)
            graph_canvas.draw_idle()
        except mysql.connector.Error as e:
            messagebox.showerror("Graph Error", str(e), parent=win)
        finally:
            if cur:
                cur.close()
            if con and con.is_connected():
                con.close()

    def refresh_report_selection():
        current_selection = summary_tree.selection()
        if not current_selection:
            return
        current_rows = [summary_tree.item(item, "values") for item in current_selection]
        report_state["summary_rows"] = current_rows
        report_state["student_ids"] = [int(row[0]) for row in current_rows]
        update_report_details(current_rows)
        load_attendance_rows()
        draw_monthly_graph()

    def search_report_students():
        """Search all registered students by name, ID, roll number, or class."""
        term = report_search_var.get().strip().casefold()
        if not term:
            messagebox.showinfo("Search Student", "Enter a student name, School ID, roll number, or class.", parent=win)
            return
        rows = tree_rows_cache.get(summary_tree, [])
        matches = [row for row in rows if any(term in str(value).casefold() for value in row[:5])]
        if not matches:
            messagebox.showinfo("Search Student", "No students were found. Check your search text.", parent=win)
            return
        report_state["summary_rows"] = matches
        report_state["student_ids"] = [int(row[0]) for row in matches]
        update_report_details(matches)
        load_attendance_rows()
        draw_monthly_graph()

    def show_all_report_students():
        """Restore the student selection that originally opened this report."""
        rows = list(initial_report_selection["summary_rows"])
        if not rows:
            return
        report_search_var.set("")
        report_state["summary_rows"] = rows
        report_state["student_ids"] = list(initial_report_selection["student_ids"])
        update_report_details(rows)
        load_attendance_rows()
        draw_monthly_graph()

    def update_report_details(current_rows):
        details_text.configure(state="normal")
        details_text.delete("1.0", "end")
        details_text.insert("1.0", "\n\n".join(
            f"Student: {row[2]}\nSchool ID: {row[1]}\nRoll No: {row[3]}\nClass: {row[4]}\n"
            f"Present: {row[5]}  Absent: {row[6]}\nLate: {row[7]}  Leave: {row[8]}\nAttendance: {row[9]}"
            for row in current_rows
        ))
        details_text.configure(state="disabled")

    win._refresh_selection = refresh_report_selection

    def close_report_window():
        global _active_summary_report
        if _active_summary_report is win:
            _active_summary_report = None
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", close_report_window)
    load_attendance_rows()
    draw_monthly_graph()
    return "break" if event else None

workspace_frame = tk.Frame(main_frame, bg="#edf2f7")
workspace_frame.pack(fill="both", expand=True, padx=6, pady=2)
workspace_frame.grid_columnconfigure(0, weight=0, minsize=205)
workspace_frame.grid_columnconfigure(1, weight=1)
workspace_frame.grid_rowconfigure(0, weight=1)

lists_frame = tk.Frame(workspace_frame, bg="#edf2f7")
lists_frame.grid_columnconfigure(0, weight=1, uniform="lists")
lists_frame.grid_columnconfigure(1, weight=1, uniform="lists")
student_panel = tk.LabelFrame(lists_frame, text="Student List", font=("Arial", 11, "bold"), bg="#edf2f7")
attendance_panel = tk.LabelFrame(lists_frame, text="Attendance List", font=("Arial", 11, "bold"), bg="#edf2f7")
student_panel.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
attendance_panel.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
student_export_frame = tk.Frame(student_panel, bg="#edf2f7")
student_export_frame.pack(fill="x", padx=3, pady=3)
tk.Button(student_export_frame, text="Export Selected Student + Attendance", command=lambda: export_selected_student(), bg="#386641", fg="white", padx=6, pady=5).pack(side="left", padx=3)
attendance_export_frame = tk.Frame(attendance_panel, bg="#edf2f7")
attendance_export_frame.pack(fill="x", padx=3, pady=3)
attendance_date_controls = tk.Frame(attendance_export_frame, bg="#edf2f7")
attendance_date_controls.pack(fill="x", pady=1)
tk.Label(attendance_date_controls, text="Date:", bg="#edf2f7").pack(side="left", padx=(2, 2))
tk.Entry(attendance_date_controls, textvariable=attendance_list_date_var, width=11).pack(side="left", padx=2)
tk.Button(attendance_date_controls, text="Calendar", command=lambda: open_calendar(attendance_list_date_var), bg="#457b9d", fg="white", padx=5).pack(side="left", padx=2)
tk.Button(attendance_date_controls, text="View All Attendance", command=lambda: open_all_attendance_window(), bg="#1d3557", fg="white", padx=5, pady=4).pack(side="left", padx=3)
tk.Button(attendance_date_controls, text="Previous Attendance", command=show_previous_attendance, bg="#6a4c93", fg="white", padx=5, pady=4).pack(side="left", padx=3)
attendance_action_controls = tk.Frame(attendance_export_frame, bg="#edf2f7")
attendance_action_controls.pack(fill="x", pady=1)
tk.Button(attendance_action_controls, text="Reset", command=reset_attendance_list, bg="#6c757d", fg="white", padx=10, pady=3).pack(side="left", padx=2)
tk.Button(attendance_action_controls, text="Edit Attendance", command=edit_summary_attendance, bg="#457b9d", fg="white", padx=10, pady=3).pack(side="left", padx=3)
student_tree = make_table(student_panel, ("ID", "Contact No", "Student Name", "Roll No", "Class"))
attendance_tree = make_table(attendance_panel, ("ID", "Student", "Roll No", "Class", "Date", "Status"))
student_search_var = add_search_bar(student_panel, student_tree, fetch_students)
attendance_search_var = add_search_bar(attendance_panel, attendance_tree, fetch_attendance)
for panel, tree in ((student_panel, student_tree), (attendance_panel, attendance_tree)):
    children = panel.pack_slaves()
    search_widget = children[-1]
    search_widget.pack_forget()
    search_widget.pack(fill="x", padx=3, pady=2, before=tree.master)
lists_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)

control_panel = tk.Frame(workspace_frame, bg="#edf2f7", width=205)
control_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 5), pady=0)
control_panel.grid_propagate(False)

class_filter_frame = tk.LabelFrame(control_panel, text="Filter Students by Class", bg="#edf2f7", padx=5, pady=4)
class_filter_frame.pack(fill="x", pady=(0, 4))
tk.Label(class_filter_frame, text="Select Class:", bg="#edf2f7").pack(anchor="w", padx=2, pady=(0, 2))
class_combo = ttk.Combobox(class_filter_frame, textvariable=class_filter_var, width=17, state="readonly", values=["All Classes"] + list(registered_class_values))
class_combo.pack(fill="x", padx=2, pady=2)

def load_class_options():
    """Load every class entered in Add Student into all class filters."""
    global registered_class_values
    con = cur = None
    try:
        con = get_connection()
        cur = con.cursor()
        cur.execute("SELECT DISTINCT class_name FROM students WHERE class_name IS NOT NULL AND TRIM(class_name)<>''")
        classes = [str(row[0]).strip() for row in cur.fetchall() if str(row[0]).strip()]
        def class_sort_key(value):
            try:
                return (0, int(value))
            except ValueError:
                return (1, value.casefold())
        registered_class_values = sorted(set(classes), key=class_sort_key)
        all_values = ["All Classes"] + registered_class_values
        if 'class_combo' in globals():
            class_combo["values"] = all_values
            if class_filter_var.get() not in all_values:
                class_filter_var.set("All Classes")
        for box in search_class_boxes.values():
            old_value = box.get()
            box["values"] = all_values
            if old_value not in all_values:
                box.set("All Classes")
        for box in list(all_attendance_class_boxes):
            try:
                old_value = box.get()
                box["values"] = all_values
                if old_value not in all_values:
                    box.set("All Classes")
            except tk.TclError:
                all_attendance_class_boxes.remove(box)
    except mysql.connector.Error as e:
        if 'class_combo' in globals():
            class_combo["values"] = ["All Classes"] + list(registered_class_values)
    finally:
        if cur:
            cur.close()
        if con and con.is_connected():
            con.close()

def apply_class_filter(event=None):
    selected_class = class_filter_var.get()
    search_class_vars[student_tree].set(selected_class)
    search_text_vars[student_tree].set("")
    search_roll_vars[student_tree].set("")
    fetch_students()
    search_tree(
        student_tree,
        search_text_vars[student_tree],
        search_class_vars[student_tree],
        search_roll_vars[student_tree]
    )
    update_class_stats()
    return "break"

class_combo.bind("<<ComboboxSelected>>", apply_class_filter)

attendance_buttons = tk.LabelFrame(control_panel, text="Attendance Actions", bg="#edf2f7", padx=4, pady=4)
attendance_buttons.pack(fill="x", pady=(0, 4))
for text, command, color in [
    ("Mark Attendance", mark_attendance, "#2a9d8f"),
    ("Delete Selected All", delete_selected_all, "#b23a48"),
    ("Delete Entire System", delete_entire_system, "#7f1d1d"),
    ("Refresh All Lists", refresh_all, "#2a9d8f")
]:
    tk.Button(attendance_buttons, text=text, command=command, bg=color, fg="white", padx=3, pady=4).pack(fill="x", padx=2, pady=2)

summary_search_var = add_search_bar(summary_panel, summary_tree, fetch_summary)
summary_search_widget = summary_panel.pack_slaves()[-1]
summary_search_widget.pack_forget()
summary_search_widget.pack(fill="x", padx=3, pady=2, before=summary_tree.master)
summary_panel.pack(fill="both", expand=True, padx=8, pady=2)

student_tree.bind("<<TreeviewSelect>>", select_student)

def handle_summary_selection(event=None):
    global _summary_report_after_id
    if event is not None and hasattr(event, "y") and not summary_tree.identify_row(event.y):
        return
    show_summary_student(event)
    if _summary_report_after_id is not None:
        try:
            root.after_cancel(_summary_report_after_id)
        except tk.TclError:
            pass
    def open_or_refresh_report():
        global _summary_report_after_id
        _summary_report_after_id = None
        if summary_tree.selection():
            show_selected_summary_report()
    _summary_report_after_id = root.after(650, open_or_refresh_report)

summary_tree.bind("<<TreeviewSelect>>", handle_summary_selection)
summary_tree.bind("<ButtonRelease-1>", handle_summary_selection, add="+")

def logout():
    username = simpledialog.askstring("Logout", "Enter username:", parent=root)
    if username is None:
        return
    password = simpledialog.askstring("Logout", "Enter password:", show="*", parent=root)
    if password is None:
        return
    if username.strip() != LOGOUT_USERNAME or password != LOGOUT_PASSWORD:
        messagebox.showerror("Logout Failed", "Wrong username or password.", parent=root)
        return
    if messagebox.askyesno("Logout", "Are you sure you want to logout?", parent=root):
        root.destroy()

menu_bar = tk.Menu(root)
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Export Selected Student", command=export_selected_student)
file_menu.add_command(label="Export All Summary", command=export_summary)
file_menu.add_command(label="Monthly Graph", command=monthly_graph)
file_menu.add_command(label="Logout", command=logout)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.destroy)
menu_bar.add_cascade(label="File", menu=file_menu)
edit_menu = tk.Menu(menu_bar, tearoff=0)
edit_menu.add_command(label="Select All Rows", command=select_all)
edit_menu.add_command(label="Delete Selected All", command=delete_selected_all)
edit_menu.add_command(label="Delete Entire System", command=delete_entire_system)
edit_menu.add_command(label="Clear Student Details", command=clear_fields)
menu_bar.add_cascade(label="Edit", menu=edit_menu)
root.config(menu=menu_bar)

for tree in (student_tree, attendance_tree, summary_tree):
    tree.bind("<Control-a>", select_all)
    tree.bind("<Control-A>", select_all)
root.bind_all("<Control-a>", select_all)
root.bind_all("<Control-A>", select_all)
root.bind_all("<Control-s>", lambda event: mark_attendance())
root.bind_all("<Control-S>", lambda event: mark_attendance())

try:
    ensure_contact_column()
except mysql.connector.Error as e:
    messagebox.showerror("Database Error", f"Contact No column could not be prepared: {e}")
load_class_options()
refresh_all()
root.mainloop()