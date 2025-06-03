import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.simpledialog import askfloat
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

class FilterWithTimeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lọc dữ liệu theo thời gian")

        self.df = None
        self.filename = None

        tk.Button(root, text="1. Chọn file dữ liệu", command=self.load_file).pack(pady=10)
        tk.Button(root, text="2. Lọc và xuất kết quả", command=self.filter_and_export).pack(pady=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
        if not file_path:
            return

        try:
            with open(file_path, encoding="ISO-8859-1", errors="replace") as f:
                lines = f.readlines()

            data_start = None
            for i, line in enumerate(lines):
                if '***DATA***' in line:
                    data_start = i
                    break

            if data_start is None or data_start + 2 >= len(lines):
                raise ValueError("Không tìm thấy phần dữ liệu!")

            header_line = lines[data_start + 1].strip()
            data_lines = lines[data_start + 2:]

            columns = header_line.split('\t')
            data = [line.strip().split('\t') for line in data_lines if line.strip()]
            self.df = pd.DataFrame(data, columns=columns).apply(pd.to_numeric, errors='coerce')
            self.filename = file_path

            messagebox.showinfo("Thành công", "Đã nạp dữ liệu!")

        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def filter_and_export(self):
        if self.df is None:
            messagebox.showwarning("Chưa có dữ liệu", "Hãy nạp dữ liệu trước.")
            return

        time_col = "Time  (Sec)"
        hour_col = "Time (Hour)"
        col_filter = "Cell Current (A)"
        target_columns = [
            "Cell Power(W)",                   
            "Anode O2 MFM Flow(sccm)",         
            "Cathode H2 MFM Flow(sccm)",       
            "Anode H2 Sensor(%)",              
            "Cathode O2 Sensor(%)"             
        ]

        columns_needed = [time_col, col_filter] + target_columns
        missing = [col for col in columns_needed if col not in self.df.columns]
        if missing:
            messagebox.showerror("Thiếu cột", f"Không tìm thấy các cột: {', '.join(missing)}")
            return

        try:
            low = askfloat("Nhập giới hạn", f"Giá trị thấp nhất cho '{col_filter}':")
            high = askfloat("Nhập giới hạn", f"Giá trị cao nhất cho '{col_filter}':")
            if low is None or high is None:
                return

            filtered_df = self.df[(self.df[col_filter] >= low) & (self.df[col_filter] <= high)]

            # Tạo cột giờ tính từ 0 dựa trên cột giây
            filtered_df[hour_col] = (filtered_df[time_col] - filtered_df[time_col].iloc[0]) / 3600

            selected_df = filtered_df[[time_col, hour_col] + target_columns]

            mean_row = selected_df.mean(numeric_only=True)
            mean_row[time_col] = None
            mean_row[hour_col] = None

            result_df = pd.concat([pd.DataFrame([mean_row]), selected_df], ignore_index=True)

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Lưu kết quả")
            if save_path:
                result_df.to_excel(save_path, index=False)

                # Tô màu dòng trung bình
                wb = load_workbook(save_path)
                ws = wb.active
                fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")
                for cell in ws[2]:
                    cell.fill = fill
                ws.cell(row=2, column=1).comment = Comment("Average value", "System")

                wb.save(save_path)
                messagebox.showinfo("Hoàn tất", f"Đã lưu file tại:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Lỗi khi xử lý", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FilterWithTimeApp(root)
    root.mainloop()
