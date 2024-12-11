from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
import os

window = Tk()
window.title("Thông tin nhân viên")
window.geometry("850x400")

file_name = "dulieunhanvien.csv"
data_columns = ["Mã", "Tên", "Ngày sinh", "Giới tính", "Đơn vị", "Số CMND", "Ngày cấp", "Chức danh", "Nơi cấp"]

# Define functions
def save_to_csv():
    employee_data = {
        "Mã": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Ngày sinh": date_entry.get(),
        "Giới tính": "Nam" if gender.get() == 1 else "Nữ",
        "Đơn vị": combobox.get(),
        "Số CMND": so_entry.get(),
        "Ngày cấp": dat_entry.get(),
        "Chức danh": T_entry.get(),
        "Nơi cấp": S_entry.get(),
    }

    if not os.path.exists(file_name):
        pd.DataFrame(columns=data_columns).to_csv(file_name, index=False)

    df = pd.read_csv(file_name)
    df = pd.concat([df, pd.DataFrame([employee_data])], ignore_index=True)
    df.to_csv(file_name, index=False)
    messagebox.showinfo("Thông báo", "Lưu thông tin thành công!")
    clear_form()

def show_today_birthdays():
    if not os.path.exists(file_name):
        messagebox.showwarning("Cảnh báo", "Không có dữ liệu!")
        return

    today = datetime.now().strftime("%d/%m/%Y")
    df = pd.read_csv(file_name)
    birthdays = df[df["Ngày sinh"] == today]

    if birthdays.empty:
        messagebox.showinfo("Thông báo", "Hôm nay không có sinh nhật nhân viên nào.")
    else:
        result = "Danh sách nhân viên sinh nhật hôm nay:\n" + "\n".join(birthdays["Tên"].astype(str).tolist())
        messagebox.showinfo("Sinh nhật hôm nay", result)


def export_to_excel():
    if not os.path.exists(file_name):
        messagebox.showwarning("Cảnh báo", "Không có dữ liệu!")
        return

    df = pd.read_csv(file_name)
    df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
    df.sort_values(by="Ngày sinh", ascending=True, inplace=True)

    output_file = "danhsachnhanvien.xlsx"
    df.to_excel(output_file, index=False, sheet_name="Danh sách nhân viên")
    messagebox.showinfo("Thông báo", f"Xuất file Excel thành công!\nFile: {output_file}")

def clear_form():
    entry_ma.delete(0, END)
    entry_ten.delete(0, END)
    date_entry.set_date(datetime.now())
    gender.set(0)
    combobox.set("")
    so_entry.delete(0, END)
    dat_entry.set_date(datetime.now())
    T_entry.delete(0, END)
    S_entry.delete(0, END)

# UI components
lbl = Label(window, text="Thông tin nhân viên", fg="black", font=("Arial", 15))
lbl.grid(column=0, row=0, columnspan=2,sticky="W")
lakh= Checkbutton(window,text="Là khách hàng")
lakh.grid(column=1,row=0,sticky="w")
lanv= Checkbutton(window,text="Là nhân viên")
lanv.grid(column=2,row=0)
ma = Label(window, text="Mã", fg="black", font=("Arial", 10))
ma.grid(column=0, row=1, sticky="W")
entry_ma = Entry(window, width=30)
entry_ma.grid(column=0, row=2, padx=5, pady=5, sticky="w")

ten = Label(window, text="Tên", fg="black", font=("Arial", 10))
ten.grid(column=1, row=1, sticky="W")
entry_ten = Entry(window, width=30)
entry_ten.grid(column=1, row=2, padx=5, pady=5)

ten = Label(window, text="Ngày sinh", fg="black", font=("Arial", 10))
ten.grid(column=2, row=1, sticky="W")
date_entry = DateEntry(window, width=20, foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
date_entry.grid(column=2, row=2, sticky="W")

gt = Label(window, text="Giới tính", fg="black", font=("Arial", 10))
gt.grid(column=3, row=1, sticky="W")
gender = IntVar()
chk3 = Radiobutton(window, text="Nam", variable=gender, value=1)
chk3.grid(row=2, column=3, padx=10, pady=5, sticky="W")
chk4 = Radiobutton(window, text="Nữ", variable=gender, value=2)
chk4.grid(row=2, column=4, padx=10, pady=5, sticky="W")

donvi = Label(window, text="Đơn vị", fg="black", font=("Arial", 10))
donvi.grid(column=0, row=3, sticky="W")
donv = StringVar()
don = ["D24CQCC01-B", "D24CQCC02-B", "D24CQCC03-B", "D24CQCC04-B"]
combobox = ttk.Combobox(window, textvariable=donv, values=don, width=27, font=("Arial", 12), state="readonly")
combobox.grid(row=4, column=0, padx=5, pady=5, sticky="W")

CM = Label(window, text="Số CMND", fg="black", font=("Arial", 10))
CM.grid(column=1, row=3, sticky="W")
so_entry = Entry(window, width=30)
so_entry.grid(column=1, row=4, sticky="W")

CM = Label(window, text="Ngày cấp", fg="black", font=("Arial", 10))
CM.grid(column=2, row=3, sticky="W")
dat_entry = DateEntry(window, width=20, foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
dat_entry.grid(column=2, row=4, sticky="W")

CD = Label(window, text="Chức danh", fg="black", font=("Arial", 10))
CD.grid(column=0, row=5, sticky="W")
T_entry = Entry(window, width=40)
T_entry.grid(column=0, row=6, sticky="W")

NC = Label(window, text="Nơi cấp", fg="black", font=("Arial", 10))
NC.grid(column=1, row=5, sticky="W")
S_entry = Entry(window, width=40)
S_entry.grid(column=1, row=6, sticky="W")

# Buttons
btn_save = Button(window, text="Lưu", command=save_to_csv, width=10, height=2)
btn_save.grid(row=7, column=0, padx=10, pady=20)

btn_today_birthdays = Button(window, text="Sinh nhật hôm nay", command=show_today_birthdays, width=15, height=2)
btn_today_birthdays.grid(row=7, column=1, padx=10, pady=20)

btn_export = Button(window, text="Xuất danh sách", command=export_to_excel, width=15, height=2)
btn_export.grid(row=7, column=2, padx=10, pady=20)

window.mainloop()

