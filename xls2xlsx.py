import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import win32com.client as win32
import pythoncom
from datetime import datetime

class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool chuyển đổi XLS -> XLSX")
        self.root.geometry("550x400")
        
        # Biến lưu trữ
        self.selected_folder = tk.StringVar()
        self.delete_original = tk.BooleanVar()
        self.is_running = False

        # --- Giao diện ---
        
        # 1. Chọn thư mục
        lbl_folder = tk.Label(root, text="Thư mục chứa file Excel:", font=("Arial", 10, "bold"))
        lbl_folder.pack(pady=(15, 5), anchor="w", padx=10)

        frame_folder = tk.Frame(root)
        frame_folder.pack(fill="x", padx=10)
        
        entry_folder = tk.Entry(frame_folder, textvariable=self.selected_folder)
        entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_browse = tk.Button(frame_folder, text="Chọn Thư Mục", command=self.browse_folder)
        btn_browse.pack(side="right")

        # 2. Tùy chọn xóa file
        chk_delete = tk.Checkbutton(root, text="Xóa file .xls cũ sau khi chuyển đổi thành công", 
                                    variable=self.delete_original, fg="red")
        chk_delete.pack(pady=15, anchor="w", padx=10)

        # 3. Nút chạy
        self.btn_convert = tk.Button(root, text="BẮT ĐẦU CHUYỂN ĐỔI", 
                                     command=self.start_conversion_thread, 
                                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
        self.btn_convert.pack(fill="x", padx=20, pady=5)

        # 4. Log hiển thị trên App
        self.log_text = tk.Text(root, height=12, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

    def log_gui(self, message):
        """Ghi log lên giao diện phần mềm"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder.set(folder)

    def start_conversion_thread(self):
        if self.is_running: return
        
        target_dir = self.selected_folder.get()
        if not target_dir or not os.path.exists(target_dir):
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục hợp lệ!")
            return

        self.is_running = True
        self.btn_convert.config(state="disabled", text="Đang xử lý...")
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")
        
        thread = threading.Thread(target=self.process_files, args=(target_dir,))
        thread.start()

    def process_files(self, target_dir):
        pythoncom.CoInitialize()
        excel = None
        
        # Danh sách để ghi log file txt
        list_success = []
        list_fail = []

        try:
            self.log_gui("Đang khởi động Excel Application...")
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False 

            self.log_gui(f"Bắt đầu quét tại: {target_dir}")
            
            for root_path, dirs, files in os.walk(target_dir):
                for filename in files:
                    if filename.lower().endswith(".xls"):
                        file_path = os.path.join(root_path, filename)
                        abs_file_path = os.path.abspath(file_path)
                        abs_output_path = abs_file_path + "x"

                        self.log_gui(f"Đang xử lý: {filename}...")
                        
                        wb = None
                        try:
                            # Mở và Convert
                            wb = excel.Workbooks.Open(abs_file_path)
                            wb.SaveAs(abs_output_path, FileFormat=51)
                            wb.Close()
                            
                            # Xử lý xóa file cũ
                            msg_extra = ""
                            if self.delete_original.get():
                                try:
                                    os.remove(abs_file_path)
                                    msg_extra = " (Đã xóa gốc)"
                                except:
                                    msg_extra = " (Lỗi xóa gốc)"
                            
                            self.log_gui(f" -> Thành công{msg_extra}")
                            
                            # Thêm vào danh sách thành công
                            list_success.append(f"- {abs_file_path} -> {abs_output_path}")

                        except Exception as e:
                            # Thêm vào danh sách thất bại
                            err_msg = str(e).replace('\n', ' ')
                            list_fail.append(f"- {abs_file_path} | Lỗi: {err_msg}")
                            self.log_gui(f" -> Lỗi: {err_msg}")
                            if wb: 
                                try: wb.Close(SaveChanges=False) 
                                except: pass

        except Exception as e:
            self.log_gui(f"Lỗi khởi tạo Excel: {e}")
        finally:
            if excel:
                try: excel.Quit()
                except: pass
            pythoncom.CoUninitialize()
            
            # --- GHI LOG RA FILE TXT ---
            self.write_log_file(target_dir, list_success, list_fail)
            
            self.is_running = False
            self.root.after(0, lambda: self.btn_convert.config(state="normal", text="BẮT ĐẦU CHUYỂN ĐỔI"))
            self.root.after(0, lambda: messagebox.showinfo("Hoàn tất", f"Đã xong!\nThành công: {len(list_success)}\nThất bại: {len(list_fail)}\nLog đã lưu tại thư mục quét."))

    def write_log_file(self, target_dir, successes, fails):
        """Hàm ghi file log.txt theo format yêu cầu"""
        log_path = os.path.join(target_dir, "conversion_log.txt")
        
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"BÁO CÁO CHUYỂN ĐỔI EXCEL\n")
                f.write(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Thư mục quét: {target_dir}\n")
                f.write("="*50 + "\n\n")

                # Ghi phần thành công
                f.write(f"THÀNH CÔNG: {len(successes)}\n")
                if successes:
                    for item in successes:
                        f.write(f"{item}\n")
                else:
                    f.write("- Không có file nào thành công.\n")

                f.write("\n" + "-"*30 + "\n\n")

                # Ghi phần thất bại
                f.write(f"THẤT BẠI: {len(fails)}\n")
                if fails:
                    for item in fails:
                        f.write(f"{item}\n")
                else:
                    f.write("- Không có file nào thất bại.\n")
            
            self.log_gui("-" * 30)
            self.log_gui(f"Đã lưu file log chi tiết tại:\n{log_path}")
            
        except Exception as e:
            self.log_gui(f"Lỗi khi ghi file log: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()