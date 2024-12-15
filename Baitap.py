import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import pandas as pd
import os
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("700x400")
FILE_CSV = "nhanvien.csv"
if not os.path.exists(FILE_CSV):
    with open(FILE_CSV, mode="w", newline='', encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp", "Nơi cấp"])
def save_info():
    data = [
        entry_ma.get(),
        entry_ten.get(),
        entry_don_vi.get(),
        entry_chuc_danh.get(),
        entry_ngay_sinh.get(),
        "Nam" if gender_var.get() == 1 else "Nữ",
        entry_cmnd.get(),
        entry_ngay_cap.get(),
        entry_noi_cap.get()
    ]
    with open(FILE_CSV, mode="a", newline='', encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)
    messagebox.showinfo("Thông báo", "Đã lưu thông tin nhân viên!")
    clear_entries()
def clear_entries():
    for entry in entries:
        entry.delete(0, tk.END)
def birthday_today():
    today = datetime.now().strftime("%d/%m")
    results = []
    with open(FILE_CSV, mode="r", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if today in row["Ngày sinh"]:
                results.append(row["Tên"])
    if results:
        messagebox.showinfo("Sinh nhật hôm nay", "\n".join(results))
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")
def export_to_excel():
    df = pd.read_csv(FILE_CSV, encoding="utf-8")
    df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y", errors='coerce')
    df = df.sort_values(by="Ngày sinh", ascending=True)
    df.to_excel("danhsach_nhanvien.xlsx", index=False, encoding="utf-8")
    messagebox.showinfo("Xuất file", "Đã xuất danh sách nhân viên ra file Excel!")
label_ma = tk.Label(root, text="Mã:")
label_ma.grid(row=0, column=0)
entry_ma = tk.Entry(root)
entry_ma.grid(row=0, column=1)
label_ten = tk.Label(root, text="Tên:")
label_ten.grid(row=0, column=2)
entry_ten = tk.Entry(root)
entry_ten.grid(row=0, column=3)
label_don_vi = tk.Label(root, text="Đơn vị:")
label_don_vi.grid(row=1, column=0)
entry_don_vi = tk.Entry(root)
entry_don_vi.grid(row=1, column=1)
label_chuc_danh = tk.Label(root, text="Chức danh:")
label_chuc_danh.grid(row=1, column=2)
entry_chuc_danh = tk.Entry(root)
entry_chuc_danh.grid(row=1, column=3)
label_ngay_sinh = tk.Label(root, text="Ngày sinh (DD/MM/YYYY):")
label_ngay_sinh.grid(row=2, column=0)
entry_ngay_sinh = tk.Entry(root)
entry_ngay_sinh.grid(row=2, column=1)
label_gioi_tinh = tk.Label(root, text="Giới tính:")
label_gioi_tinh.grid(row=2, column=2)
gender_var = tk.IntVar(value=1)
radio_nam = tk.Radiobutton(root, text="Nam", variable=gender_var, value=1)
radio_nu = tk.Radiobutton(root, text="Nữ", variable=gender_var, value=2)
radio_nam.grid(row=2, column=3, sticky="w")
radio_nu.grid(row=2, column=3, sticky="e")
label_cmnd = tk.Label(root, text="Số CMND:")
label_cmnd.grid(row=3, column=0)
entry_cmnd = tk.Entry(root)
entry_cmnd.grid(row=3, column=1)
label_ngay_cap = tk.Label(root, text="Ngày cấp:")
label_ngay_cap.grid(row=3, column=2)
entry_ngay_cap = tk.Entry(root)
entry_ngay_cap.grid(row=3, column=3)
label_noi_cap = tk.Label(root, text="Nơi cấp:")
label_noi_cap.grid(row=4, column=0)
entry_noi_cap = tk.Entry(root)
entry_noi_cap.grid(row=4, column=1)
entries = [entry_ma, entry_ten, entry_don_vi, entry_chuc_danh, entry_ngay_sinh, entry_cmnd, entry_ngay_cap, entry_noi_cap]
btn_save = tk.Button(root, text="Lưu thông tin", command=save_info)
btn_save.grid(row=5, column=0)
btn_birthday = tk.Button(root, text="Sinh nhật ngày hôm nay", command=birthday_today)
btn_birthday.grid(row=5, column=1)
btn_export = tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export.grid(row=5, column=2)
root.mainloop()