import tkinter as tk
from tkinter import messagebox, Listbox, filedialog
import pyodbc
from collections import Counter

# تابع اتصال به پایگاه داده Access و دریافت اطلاعات از ستون خاص
def fetch_column_from_access(db_path, table_name, column_name):
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + db_path + ';'
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        query = f"SELECT [{column_name}] FROM [{table_name}]"
        cursor.execute(query)

        column_data = [row[0] for row in cursor.fetchall()]
    except pyodbc.Error as e:
        print("Error:", e)
        column_data = []
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

    return column_data

# تابع شمارش تکرار مقاطع
def count_duplicates(input_list):
    return list(Counter(input_list).items())

# تابع پردازش داده‌ها
def process_data():
    db_path = entry_db_path.get()

    # دریافت داده‌ها از پایگاه داده Access
    try:
        b_section = fetch_column_from_access(db_path, "Frame Assignments - Section Properties", "Section Property")
        b_uniqname = fetch_column_from_access(db_path, "Frame Assignments - Section Properties", "UniqueName")
        c_end_point = fetch_column_from_access(db_path, "Column Object Connectivity", "UniquePtJ")
        b_uniqname_conectivity = fetch_column_from_access(db_path, "Beam Object Connectivity", "Unique Name")
        b_strat_point_conectivity = fetch_column_from_access(db_path, "Beam Object Connectivity", "UniquePtI")
        b_end_point_conectivity = fetch_column_from_access(db_path, "Beam Object Connectivity", "UniquePtJ")
        b_realese_summary_name = fetch_column_from_access(db_path, "Frame Assignments - Summary", "UniqueName")
        b_realese_summary_releses = fetch_column_from_access(db_path, "Frame Assignments - Summary", "Releases")
    except Exception as e:
        messagebox.showerror("Error", f"Error fetching data: {e}")
        return

    # دسته‌بندی تیرها
    b_c_list = []
    b_cantiliver = []
    b_cantiliver_joint = []
    for i in range(len(b_uniqname_conectivity)):
        st_point = b_strat_point_conectivity[i]
        end_point = b_end_point_conectivity[i]
        if st_point in c_end_point and end_point in c_end_point:
            b_c_list.append(b_uniqname_conectivity[i])
        elif st_point in c_end_point or end_point in c_end_point:
            b_cantiliver.append(b_uniqname_conectivity[i])

    # تیر به تیر
    b_to_b_list = [item for item in b_uniqname_conectivity if item not in b_cantiliver and item not in b_c_list]

    # اتصالات گیردار
    b_to_c_joint = []
    for i in b_c_list:
        x = b_realese_summary_name.index(i) if i in b_realese_summary_name else -1
        if x != -1 and b_realese_summary_releses[x] == "Yes":
            b_to_c_joint.append(i)

    b_to_c_fixed = [item for item in b_c_list if item not in b_to_c_joint]
    for i in b_cantiliver:
        x = b_realese_summary_name.index(i) if i in b_realese_summary_name else -1
        if x != -1 and b_realese_summary_releses[x] == "Yes":
            b_cantiliver_joint.append(i)

    # دسته‌بندی مقاطع
    def get_sections(beam_list):
        return [b_section[b_uniqname.index(i)] for i in beam_list if i in b_uniqname]
    
    b_cantiliver = [item for item in b_cantiliver if item not in b_cantiliver_joint]

    b_cantiliver_section = get_sections(b_cantiliver)
    b_c_fixed_section = get_sections(b_to_c_fixed)
    b_c_joint_section = get_sections(b_to_c_joint)
    b_b_section = get_sections(b_to_b_list)
    b_cantiliver_joint_section = get_sections(b_cantiliver_joint)


    # نمایش نتایج
    display_results("--- مقاطع تیرهای طره‌ای ---", b_cantiliver_section)
    display_results("--- مقاطع تیرهای گیردار به ستون ---", b_c_fixed_section)
    display_results("--- مقاطع تیرهای مفصلی به ستون ---", b_c_joint_section)
    display_results("--- مقاطع تیر به تیر ---", b_b_section)
    display_results("--- تیر طره مفصلی که با یک تیر دیگر پایدار شده ---", b_cantiliver_joint_section)

# تابع نمایش نتایج در Listbox
def display_results(title, sections):
    listbox_results.insert(tk.END, title)
    if sections:
        section_counts = count_duplicates(sections)
        for section, count in section_counts:
            listbox_results.insert(tk.END, f"مقطع: {section}, تعداد: {count}")
    else:
        listbox_results.insert(tk.END, "هیچ مقطعی یافت نشد.")
    listbox_results.insert(tk.END, "")  # خط خالی برای جداسازی

# تابع برای انتخاب فایل Access
def select_file():
    file_path = filedialog.askopenfilename(title="انتخاب فایل Access", filetypes=[("Access files", "*.accdb;*.mdb")])
    if file_path:
        entry_db_path.delete(0, tk.END)  # پاک کردن ورودی قبلی
        entry_db_path.insert(0, file_path)  # قرار دادن مسیر فایل انتخاب شده

# تابع پاک‌سازی نتایج
def clear_results():
    listbox_results.delete(0, tk.END)  # حذف تمام آیتم‌های Listbox

# ساخت رابط کاربری
root = tk.Tk()
root.title("شمارش اتصالات و مقاطع تیرها")

# بخش وارد کردن مسیر دیتابیس
label_db_path = tk.Label(root, text="مسیر دیتابیس Access:")
label_db_path.pack(pady=5)

entry_db_path = tk.Entry(root, width=50)
entry_db_path.pack(pady=5)

# دکمه برای انتخاب فایل
select_button = tk.Button(root, text="انتخاب فایل", command=select_file)
select_button.pack(pady=5)

# دکمه پردازش
process_button = tk.Button(root, text="پردازش", command=process_data)
process_button.pack(pady=10)

# دکمه پاک‌سازی خروجی
clear_button = tk.Button(root, text="پاک‌سازی خروجی", command=clear_results)
clear_button.pack(pady=5)

# لیست‌نمایش نتایج
listbox_results = Listbox(root, width=80, height=20)
listbox_results.pack(pady=20)

# اجرای برنامه
root.mainloop()
