import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
from datetime import datetime, timedelta
from PIL import Image, ImageTk
from tkcalendar import DateEntry

# MAIN WINDOW
root = tk.Tk()
root.title("Library Management System")
root.geometry("1100x650")

# Hide system before login
root.withdraw()

# DATA STORAGE
books = {}
borrow_records = {}
excel_file_path = ""

# ADMIN LOGIN
ADMIN_USERNAME = "1"
ADMIN_PASSWORD = "1"

# CLOSE PROGRAM FUNCTION
def close_program():

    answer = messagebox.askyesno(
        "Exit",
        "Are you sure you want to exit?"
    )

    if answer:
        root.destroy()

# ============= LOGIN FUNCTION =======================
def login():

    username = entry_username.get()
    password = entry_password.get()

    if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:

        messagebox.showinfo(
            "Success",
            "Login Successful!"
        )
        login_window.destroy()
        root.deiconify()

    else:

        messagebox.showerror(
            "Error",
            "Invalid Username or Password"
        )

# LOGIN WINDOW
login_window = tk.Toplevel()
login_window.title("Admin Login Page")
login_window.geometry("1100x650")
login_window.config(bg="#bae6fd")

# LOGIN IMAGE
image = Image.open("gambar.png")

image = image.resize((500, 150))

photo = ImageTk.PhotoImage(image)

image_label = tk.Label(
    login_window,
    image=photo,
    bg="#bae6fd"
)

image_label.image = photo
image_label.pack(pady=20)

# LOGIN TITLE
tk.Label(
    login_window,
    text="📚 Library Management Login",
    font=("Arial", 22, "bold"),
    bg="#bae6fd",
    fg="#1e3a8a"
).pack(pady=20)

# USERNAME
tk.Label(
    login_window,
    text="Username",
    bg="#bae6fd",
    font=("Arial", 12)
).pack()

entry_username = tk.Entry(
    login_window,
    width=30,
    font=("Arial", 12)
)
entry_username.pack(pady=10)

# PASSWORD
tk.Label(
    login_window,
    text="Password",
    bg="#bae6fd",
    font=("Arial", 12)
).pack()

entry_password = tk.Entry(
    login_window,
    width=30,
    show="*",
    font=("Arial", 12)
)
entry_password.pack(pady=10)

# LOGIN BUTTON
tk.Button(
    login_window,
    text="Login",
    width=20,
    bg="#2563eb",
    fg="white",
    font=("Arial", 11, "bold"),
    command=login
).pack(pady=20)

# =========================================================
# After Sucessfull Login
# SIDEBAR
# =========================================================

sidebar = tk.Frame(
    root,
    bg="#2c3e50",
    width=220
)

sidebar.pack(side="left", fill="y")

# =========================================================
# MAIN CONTENT
# =========================================================

main_content = tk.Frame(
    root,
    bg="white"
)

main_content.pack(
    side="right",
    expand=True,
    fill="both"
)

# =========================================================
# CLEAR PAGE
# =========================================================

def clear_page():

    for widget in main_content.winfo_children():
        widget.destroy()

# =========================================================
# HOME PAGE
# =========================================================

def show_home():

    clear_page()

    tk.Label(
        main_content,
        text="📚 Welcome To Library Management System",
        font=("Arial", 28, "bold"),
        bg="white",
        fg="#2c3e50"
    ).pack(pady=120)

    tk.Label(
        main_content,
        text="Manage books, borrowing and returning efficiently",
        font=("Arial", 14),
        bg="white"
    ).pack()

# =========================================================
# ADD BOOK FUNCTION
# =========================================================

def add_book():

    book_name = entry_book.get()
    author_name = entry_author.get()

    if book_name == "" or author_name == "":

        messagebox.showerror(
            "Error",
            "Please fill all fields"
        )

        return

    if book_name in books:

        messagebox.showerror(
            "Error",
            "Book already exists"
        )

        return

    books[book_name] = {
        "author": author_name,
        "status": "Available",
        "borrow_date": "",
        "due_date": ""
    }

    messagebox.showinfo(
        "Success",
        f"{book_name} added successfully!"
    )

    entry_book.delete(0, tk.END)
    entry_author.delete(0, tk.END)

# =========================================================
# CHOOSE EXCEL FILE
# =========================================================

def choose_excel():

    global excel_file_path

    excel_file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if excel_file_path:

        excel_file_label.config(
            text=excel_file_path.split("/")[-1]
        )

# =========================================================
# IMPORT BOOKS FROM EXCEL
# =========================================================

def import_excel_books():

    global excel_file_path

    if excel_file_path == "":

        messagebox.showerror(
            "Error",
            "Please choose an Excel file"
        )

        return

    try:

        df = pd.read_excel(
        excel_file_path,
        skiprows=4
        )

        for index, row in df.iterrows():

            book_name = str(row["Book Name"])
            author_name = str(row["Author"])

            if book_name not in books:

                books[book_name] = {
                    "author": author_name,
                    "status": "Available",
                    "borrow_date": "",
                    "due_date": ""
                }

        messagebox.showinfo(
            "Success",
            "Books imported successfully!"
        )

        show_all_books()

    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )
# =========================================================
# ADD BOOK PAGE
# =========================================================

def show_add_book():

    clear_page()

    tk.Label(
        main_content,
        text="Add Book",
        font=("Arial", 24, "bold"),
        bg="white"
    ).pack(pady=20)

    container = tk.Frame(
        main_content,
        bg="white"
    )

    container.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=20
    )

    # =====================================================
    # LEFT FRAME
    # =====================================================

    left_frame = tk.Frame(
        container,
        bg="#ecf0f1",
        bd=2,
        relief="solid"
    )

    left_frame.pack(
        side="left",
        fill="both",
        expand=True,
        padx=20
    )

    tk.Label(
        left_frame,
        text="Add Book Manually",
        font=("Arial", 18, "bold"),
        bg="#ecf0f1"
    ).pack(pady=20)

    tk.Label(
        left_frame,
        text="Book Name",
        bg="#ecf0f1"
    ).pack()

    global entry_book

    entry_book = tk.Entry(
        left_frame,
        width=30
    )

    entry_book.pack(pady=10)

    tk.Label(
        left_frame,
        text="Author Name",
        bg="#ecf0f1"
    ).pack()

    global entry_author

    entry_author = tk.Entry(
        left_frame,
        width=30
    )

    entry_author.pack(pady=10)

    tk.Button(
        left_frame,
        text="Add Book",
        width=20,
        bg="#2563eb",
        fg="white",
        command=add_book
    ).pack(pady=20)

    # =====================================================
    # RIGHT FRAME
    # =====================================================

    right_frame = tk.Frame(
        container,
        bg="#dfe6e9",
        bd=2,
        relief="solid"
    )

    right_frame.pack(
        side="right",
        fill="both",
        expand=True,
        padx=20
    )

    tk.Label(
        right_frame,
        text="Import Books From Excel",
        font=("Arial", 18, "bold"),
        bg="#dfe6e9"
    ).pack(pady=20)

    tk.Label(
        right_frame,
        text="Upload Excel File (.xlsx)",
        bg="#dfe6e9"
    ).pack(pady=10)

    global excel_file_label

    excel_file_label = tk.Label(
        right_frame,
        text="No file selected",
        bg="#dfe6e9"
    )

    excel_file_label.pack(pady=10)

    tk.Button(
        right_frame,
        text="Choose Excel File",
        width=20,
        command=choose_excel
    ).pack(pady=10)

    tk.Button(
        right_frame,
        text="Import Books",
        width=20,
        bg="#16a34a",
        fg="white",
        command=import_excel_books
    ).pack(pady=20)

# =========================================================
# BORROW BOOK FUNCTION
# =========================================================

def borrow_book():

    student_name = entry_student_name.get()
    matrix_id = entry_matrix.get()
    book_name = entry_borrow_book.get()

    if (
        student_name == ""
        or matrix_id == ""
        or book_name == ""
    ):

        messagebox.showerror(
            "Error",
            "Please fill all fields"
        )

        return

    if book_name not in books:

        messagebox.showerror(
            "Error",
            "Book does not exist"
        )

        return

    if books[book_name]["status"] == "Borrowed":

        messagebox.showerror(
            "Error",
            "Book already borrowed"
        )

        return

    borrow_date = start_date.get_date()
    due_date = end_date.get_date()

    if due_date <= borrow_date :
        messagebox.showerror(
            "Error",
            "End Date must be after Start Date"
        )
        return
    
    books[book_name]["status"] = "Borrowed"
    books[book_name]["borrow_date"] = borrow_date.strftime("%Y-%m-%d")
    books[book_name]["due_date"] = due_date.strftime("%Y-%m-%d")
    books[book_name]["borrower_name"] = student_name
    books[book_name]["matrix_id"] = matrix_id

    if matrix_id not in borrow_records:

        borrow_records[matrix_id] = {
            "name": student_name,
            "books": []
        }

    borrow_records[matrix_id]["books"].append(book_name)

    messagebox.showinfo(
        "Success",
        f"{book_name} borrowed successfully!\nReturn before {due_date.strftime('%Y-%m-%d')}"
    )
    show_borrow()

# =========================================================
# BORROW PAGE
# =========================================================

def show_borrow():

    clear_page()

    tk.Label(
        main_content,
        text="Borrow Book",
        font=("Arial", 24, "bold"),
        bg="white",
        fg="#2c3e50"
    ).pack(pady=(30, 10))

    tk.Label(
        main_content,
        text="Manage student borrowing records",
        font=("Arial", 11),
        bg="white",
        fg="gray"
    ).pack()

    card = tk.Frame(
        main_content,
        bg="white",
        bd=1,
        relief="solid"
    )

    card.pack(
        pady=30,
        ipadx=50,
        ipady=25
    )

    card.grid_columnconfigure(0, weight=1)

    # Student Name
    tk.Label(
        card,
        text="Student Name",
        bg="white",
        font=("Arial", 11, "bold")
    ).grid(
        row=0,
        column=0,
        padx=30,
        pady=(10, 5)
    )

    global entry_student_name

    entry_student_name = tk.Entry(
        card,
        font=("Arial", 11)
    )

    entry_student_name.grid(
        row=1,
        column=0,
        padx=30,
        sticky="ew"
    )

    # Matrix ID
    tk.Label(
        card,
        text="Matrix ID",
        bg="white",
        font=("Arial", 11, "bold")
    ).grid(
        row=2,
        column=0,
        padx=30,
        pady=(15, 5)
    )

    global entry_matrix

    entry_matrix = tk.Entry(
        card,
        font=("Arial", 11)
    )

    entry_matrix.grid(
        row=3,
        column=0,
        padx=30,
        sticky="ew"
    )

    # Book Name
    tk.Label(
        card,
        text="Book Name",
        bg="white",
        font=("Arial", 11, "bold")
    ).grid(
        row=4,
        column=0,
        padx=30,
        pady=(15, 5)
    )

    global entry_borrow_book

    entry_borrow_book = ttk.Combobox(
        card,
        state="readonly"
    )

    available_books = []

    for book, details in books.items():
        if details["status"] == "Available":
            available_books.append(book)

    entry_borrow_book["values"] = available_books

    entry_borrow_book.grid(
        row=5,
        column=0,
        padx=30,
        sticky="ew"
    )

    # Start Date
    tk.Label(
        card,
        text="Start Date",
        bg="white",
        font=("Arial", 11, "bold")
    ).grid(
        row=6,
        column=0,
        padx=30,
        pady=(15, 5)
    )

    global start_date

    start_date = DateEntry(
        card,
        date_pattern="yyyy-mm-dd",
        state="readonly"
    )

    start_date.grid(
        row=7,
        column=0,
        padx=30,
        sticky="ew"
    )

    # End Date
    tk.Label(
        card,
        text="End Date",
        bg="white",
        font=("Arial", 11, "bold")
    ).grid(
        row=8,
        column=0,
        padx=30,
        pady=(15, 5)
    )

    global end_date

    end_date = DateEntry(
        card,
        date_pattern="yyyy-mm-dd",
        state="readonly"
    )

    end_date.grid(
        row=9,
        column=0,
        padx=30,
        sticky="ew"
    )

    # Button
    tk.Button(
        card,
        text="Borrow Book",
        bg="#2563eb",
        fg="white",
        font=("Arial", 10, "bold"),
        width=20,
        command=borrow_book
    ).grid(
        row=10,
        column=0,
        pady=25
    )

# =========================================================
# LOAD STUDENT DATA
# =========================================================

def load_student_data():

    matrix_id = entry_return_matrix.get()

    if matrix_id not in borrow_records:

        messagebox.showerror(
            "Error",
            "Student record not found"
        )

        return

    student = borrow_records[matrix_id]

    label_student_name.config(
        text=f"Student Name: {student['name']}"
    )

    combo_books["values"] = student["books"]

# =========================================================
# RETURN BOOK FUNCTION
# =========================================================

def return_book():

    matrix_id = entry_return_matrix.get()

    selected_book = combo_books.get()

    if selected_book == "":

        messagebox.showerror(
            "Error",
            "Please select a book"
        )

        return

    due_date = datetime.strptime(
        books[selected_book]["due_date"],
        "%Y-%m-%d"
    )

    today = datetime.combine(
        return_date.get_date(),
        datetime.min.time()
    )

    fine = 0

    if today > due_date:

        late_days = (today - due_date).days

        fine = late_days * 1

    books[selected_book]["status"] = "Available"
    books[selected_book]["borrow_date"] = ""
    books[selected_book]["due_date"] = ""
    books[selected_book]["borrower_name"] = ""
    books[selected_book]["matrix_id"] = ""

    borrow_records[matrix_id]["books"].remove(selected_book)

    combo_books["values"] = borrow_records[matrix_id]["books"]

    if fine > 0:

        messagebox.showwarning(
            "Late Return",
            f"Book returned late!\nFine = RM{fine}"
        )

    else:

        messagebox.showinfo(
            "Success",
            "Book returned successfully!"
        )
    show_return()

# =========================================================
# RETURN PAGE
# =========================================================

def show_return():

    clear_page()
    FIELD_WIDTH = 28

    # ==========================
    # PAGE TITLE
    # ==========================
    tk.Label(
        main_content,
        text="Return Book",
        font=("Arial", 24, "bold"),
        bg="white",
        fg="#2c3e50"
    ).pack(pady=(30, 10))

    tk.Label(
        main_content,
        text="Manage book returns and update library records",
        font=("Arial", 11),
        bg="white",
        fg="gray"
    ).pack()

    # ==========================
    # CARD FRAME
    # ==========================
    card = tk.Frame(
        main_content,
        bg="#f8fafc",
        bd=1,
        relief="solid"
    )

    card.pack(
        pady=30,
        ipadx=50, #kanan
        ipady=25 #kiri
    )

    card.grid_columnconfigure(0, weight=1)

    # ==========================
    # MATRIX ID
    # ==========================
    tk.Label(
        card,
        text="Matrix ID",
        font=("Arial", 11, "bold"),
        bg="#f8fafc"
    ).grid(
        row=0,
        column=0,
        padx=30,
        pady=(10, 5)
    )

    global entry_return_matrix

    entry_return_matrix = tk.Entry(
        card,
        width=FIELD_WIDTH,
        font=("Arial", 11)
    )

    entry_return_matrix.grid(
        row=1,
        column=0,
        padx=30
    )

    # ==========================
    # CURRENT DATE
    # ==========================
    tk.Label(
        card,
        text="Current Date",
        font=("Arial", 11, "bold"),
        bg="#f8fafc",
    ).grid(
        row=2,
        column=0,
        pady=(0, 5)
    )

    global return_date

    return_date = DateEntry(
        card,
        width=FIELD_WIDTH,
        date_pattern="yyyy-mm-dd",
        state="readonly"
    )

    return_date.grid(
        row=3,
        column=0,
        padx=30,
        pady=(10, 5)
    )

    # ==========================
    # CHECK STUDENT BUTTON
    # ==========================
    tk.Button(
        card,
        text="Check Student",
        width=18,
        bg="#2563eb",
        fg="white",
        font=("Arial", 10, "bold"),
        command=load_student_data
    ).grid(
        row=4,
        column=0,
        pady=(0, 20)
    )
    # ==========================
    # STUDENT NAME
    # ==========================
    global label_student_name

    label_student_name = tk.Label(
        card,
        text="Student Name: ",
        font=("Arial", 11, "bold"),
        bg="#f8fafc",
        anchor="w"
    )

    label_student_name.grid(
        row=5,
        column=0,
        pady=(0, 10)
    )

    # ==========================
    # BOOK COMBOBOX
    # ==========================
    global combo_books

    combo_books = ttk.Combobox(
        card,
        width=FIELD_WIDTH,
        state="readonly"
    )

    combo_books.grid(
        row=6,
        column=0,
        padx=30,
        pady=(0, 25)
    )

    # ==========================
    # RETURN BUTTON
    # ==========================
    tk.Button(
        card,
        text="Return Book",
        width=20,
        bg="#16a34a",
        fg="white",
        font=("Arial", 10, "bold"),
        command=return_book
    ).grid(
        row=7,
        column=0,
        pady=(5, 10)
    )

# =========================================================
# CHECK STATUS PAGE
# =========================================================

def check_status():

    clear_page()

    tk.Label(
        main_content,
        text="Check Book Status",
        font=("Arial", 24, "bold"),
        bg="white"
    ).pack(pady=20)

    tk.Label(
        main_content,
        text="Select Book",
        bg="white"
    ).pack()

    global entry_check_book

    entry_check_book = ttk.Combobox(
        main_content,
        width=40,
        state="readonly"
    )

    entry_check_book["values"] = list(books.keys())

    entry_check_book.pack(pady=10)

    tk.Button(
        main_content,
        text="Check Status",
        command=display_status
    ).pack(pady=20)

    global status_output_frame

    status_output_frame = tk.Frame(
    main_content,
    bg="white"
)

    status_output_frame.pack(

    fill="x",
    padx=20,
    pady=20
)

# =========================================================
# DISPLAY STATUS
# =========================================================

def display_status():

    book_name = entry_check_book.get()

    if book_name not in books:

        messagebox.showerror(
            "Error",
            "Book does not exist in library"
        )

        return

    info = books[book_name]

    for widget in status_output_frame.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(
        status_output_frame,
        columns=(
            "Book",
            "Author",
            "Status",
            "Borrower",
            "Matrix",
            "Borrow Date",
            "Due Date"
        ),
        show="headings",
        height=1
    )

    tree.heading("Book", text="Book Name")
    tree.heading("Author", text="Author")
    tree.heading("Status", text="Status")
    tree.heading("Borrower", text="Borrowed By")
    tree.heading("Matrix", text="Matrix ID")
    tree.heading("Borrow Date", text="Borrow Date")
    tree.heading("Due Date", text="Due Date")

    tree.column("Book", width=180)
    tree.column("Author", width=120)
    tree.column("Status", width=100)
    tree.column("Borrower", width=150)
    tree.column("Matrix", width=100)
    tree.column("Borrow Date", width=100)
    tree.column("Due Date", width=100)

    tree.pack(fill="x")

    tree.insert(
        "",
        tk.END,
        values=(
            book_name,
            info["author"],
            info["status"],
            info.get("borrower_name", "-"),
            info.get("matrix_id", "-"),
            info["borrow_date"],
            info["due_date"]
        )
    )

def show_all_books():

    clear_page()

    tk.Label(
        main_content,
        text="Library Book List",
        font=("Arial", 24, "bold"),
        bg="white"
    ).pack(pady=20)

    tree = ttk.Treeview(
        main_content,
        columns=("Book", "Author", "Status"),
        show="headings"
    )

    tree.heading("Book", text="Book Name")
    tree.heading("Author", text="Author")
    tree.heading("Status", text="Status")

    tree.column("Book", width=350)
    tree.column("Author", width=250)
    tree.column("Status", width=150)

    tree.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=20
    )

    for book, details in books.items():

        tree.insert(
            "",
            tk.END,
            values=(
                book,
                details["author"],
                details["status"]
            )
        )

# =========================================================
# SIDEBAR TITLE
# =========================================================

tk.Label(
    sidebar,
    text="📚 MENU",
    bg="#2c3e50",
    fg="white",
    font=("Arial", 18, "bold")
).pack(pady=30)

# =========================================================
# SIDEBAR BUTTONS
# =========================================================

tk.Button(
    sidebar,
    text="🏠 Home",
    width=20,
    command=show_home
).pack(pady=10)

tk.Button(
    sidebar,
    text="📚 Add Book",
    width=20,
    command=show_add_book
).pack(pady=10)

tk.Button(
    sidebar,
    text="📖 Borrow Book",
    width=20,
    command=show_borrow
).pack(pady=10)

tk.Button(
    sidebar,
    text="↩ Return Book",
    width=20,
    command=show_return
).pack(pady=10)

tk.Button(
    sidebar,
    text="📋 Check Status",
    width=20,
    command=check_status
).pack(pady=10)

tk.Button(
    sidebar,
    text="📕 View Books",
    width=20,
    command=show_all_books
).pack(pady=10)

# =========================================================
# DEFAULT PAGE
# =========================================================

show_home()

# =========================================================
# CLOSE WINDOW EVENT
# =========================================================

root.protocol("WM_DELETE_WINDOW", close_program)

login_window.protocol("WM_DELETE_WINDOW", close_program)

# =========================================================
# RUN PROGRAM
# =========================================================

root.mainloop()
