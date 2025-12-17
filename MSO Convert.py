import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import win32com.client as win32
import pythoncom
from datetime import datetime

def resource_path(relative_path):
    """ Lấy đường dẫn tuyệt đối đến tài nguyên, hoạt động cho dev và cho PyInstaller """
    try:
        # PyInstaller tạo một thư mục tạm thời và lưu đường dẫn trong _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class OfficeConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Chuyển đổi định dạng Office")
        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except:
            # Bỏ qua nếu không tìm thấy icon hoặc lỗi
            pass
        self.root.geometry("550x420")
        self.root.resizable(False, False)
        
        # Biến lưu trữ
        self.selected_folder = tk.StringVar()
        self.delete_original = tk.BooleanVar()
        self.scan_subdirectories = tk.BooleanVar(value=True) # Mặc định là quét thư mục con
        self.is_running = False

        # Các biến cho tùy chọn định dạng
        self.convert_xls = tk.BooleanVar(value=True)
        self.convert_doc = tk.BooleanVar(value=True)
        self.convert_ppt = tk.BooleanVar(value=True)

        # --- Giao diện ---
        
        # 1. Chọn thư mục
        lbl_folder = tk.Label(root, text="Thư mục chứa file Office:", font=("Arial", 10, "bold"))
        lbl_folder.pack(pady=(15, 5), anchor="w", padx=10)

        frame_folder = tk.Frame(root)
        frame_folder.pack(fill="x", padx=10)
        
        entry_folder = tk.Entry(frame_folder, textvariable=self.selected_folder)
        entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_browse = tk.Button(frame_folder, text="Chọn Thư Mục", command=self.browse_folder)
        btn_browse.pack(side="right")

        # 2. Tùy chọn chung
        frame_options = tk.Frame(root)
        frame_options.pack(fill="x", padx=5, pady=5)
        
        chk_subdirs = tk.Checkbutton(frame_options, text="Bao gồm các thư mục con", 
                                     variable=self.scan_subdirectories)
        chk_subdirs.pack(anchor="w")

        chk_delete = tk.Checkbutton(frame_options, text="Xóa file cũ sau khi chuyển đổi thành công", 
                                    variable=self.delete_original, fg="red")
        chk_delete.pack(anchor="w")

        # 3. Tùy chọn định dạng
        lbl_formats = tk.Label(root, text="Chọn định dạng để chuyển đổi:", font=("Arial", 10, "bold"))
        lbl_formats.pack(pady=(10, 5), anchor="w", padx=10)

        frame_formats = tk.Frame(root)
        frame_formats.pack(fill="x", padx=10)

        # Sắp xếp lại các checkbox theo hàng ngang
        tk.Checkbutton(frame_formats, text="Excel", variable=self.convert_xls).pack(side="left", padx=5)
        tk.Checkbutton(frame_formats, text="Word", variable=self.convert_doc).pack(side="left", padx=5)
        tk.Checkbutton(frame_formats, text="PowerPoint", variable=self.convert_ppt).pack(side="left", padx=5)
        
        # 4. Nút chạy
        self.btn_convert = tk.Button(root, text="BẮT ĐẦU CHUYỂN ĐỔI", 
                                     command=self.start_conversion_thread, 
                                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
        self.btn_convert.pack(fill="x", padx=20, pady=5)

        # 5. Log hiển thị trên App
        self.log_text = tk.Text(root, height=10, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

    def log_gui(self, message):
        """Ghi log lên giao diện phần mềm"""
        self.root.after(0, self._log_gui_thread_safe, message)

    def _log_gui_thread_safe(self, message):
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

        # Kiểm tra xem ít nhất một định dạng đã được chọn để chuyển đổi chưa
        if not (self.convert_xls.get() or self.convert_doc.get() or
                self.convert_ppt.get()):
            messagebox.showerror("Lỗi", "Vui lòng chọn ít nhất một định dạng để chuyển đổi!")
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
        excel, word, powerpoint = None, None, None
        
        list_success = []
        list_fail = []

        try:
            self.log_gui("Đang khởi động các ứng dụng Office...")
            excel = win32.Dispatch("Excel.Application")
            word = win32.Dispatch("Word.Application")
            powerpoint = win32.Dispatch("PowerPoint.Application")
            
            excel.Visible, word.Visible = False, False
            # powerpoint.Visible = False # Gây lỗi trên một số máy, không ẩn cửa sổ PowerPoint

            excel.DisplayAlerts, word.DisplayAlerts = False, False

            self.log_gui(f"Bắt đầu quét tại: {target_dir}")
            
            for root_path, _, files in os.walk(target_dir):
                for filename in files:
                    file_ext = os.path.splitext(filename)[1].lower()

                    should_process = False
                    if file_ext == ".xls" and self.convert_xls.get(): should_process = True
                    elif file_ext == ".doc" and self.convert_doc.get(): should_process = True
                    elif file_ext == ".ppt" and self.convert_ppt.get(): should_process = True

                    if not should_process or filename.startswith('~'):
                        continue

                    file_path = os.path.join(root_path, filename)
                    abs_file_path = os.path.abspath(file_path)
                    
                    abs_output_path = abs_file_path + "x"

                    self.log_gui(f"Đang xử lý: {filename}...")

                    if os.path.exists(abs_output_path):
                        self.log_gui(f" -> File đích '{os.path.basename(abs_output_path)}' đã tồn tại. Thử tạo file mới...")
                        
                        base_name, ext = os.path.splitext(abs_output_path)
                        abs_output_path = f"{base_name}_{ext}"

                        if os.path.exists(abs_output_path):
                            self.log_gui(f" -> Bỏ qua, file '{os.path.basename(abs_output_path)}' cũng đã tồn tại.")
                            continue
                    
                    document = None
                    try:
                        if file_ext == ".xls":
                            document = excel.Workbooks.Open(abs_file_path)
                            document.SaveAs(abs_output_path, FileFormat=51) # xlOpenXMLWorkbook
                        elif file_ext == ".doc":
                            document = word.Documents.Open(abs_file_path)
                            document.SaveAs(abs_output_path, FileFormat=16) # wdFormatXMLDocument
                        elif file_ext == ".ppt":
                            document = powerpoint.Presentations.Open(abs_file_path, WithWindow=False)
                            document.SaveAs(abs_output_path, FileFormat=24) # ppSaveAsOpenXMLPresentation

                        if document:
                            document.Close()
                        
                        msg_extra = ""
                        if self.delete_original.get():
                            try:
                                os.remove(abs_file_path)
                                msg_extra = " (Đã xóa gốc)"
                            except Exception as del_e:
                                msg_extra = f" (Lỗi xóa gốc: {del_e})"
                        
                        self.log_gui(f" -> Chuyển đổi thành công sang '{os.path.basename(abs_output_path)}'{msg_extra}")
                        list_success.append(f"- {abs_file_path} -> {abs_output_path}{msg_extra}")

                    except Exception as e:
                        err_msg = str(e).replace('\n', ' ').replace('\r', '')
                        list_fail.append(f"- {abs_file_path} | Lỗi: {err_msg}")
                        self.log_gui(f" -> Lỗi: {err_msg}")
                        if document: 
                            try: document.Close(SaveChanges=False) 
                            except: pass
                
                if not self.scan_subdirectories.get():
                    break

        except Exception as e:
            self.log_gui(f"Lỗi nghiêm trọng khi khởi tạo hoặc xử lý: {e}")
        finally:
            if excel:
                try: excel.Quit()
                except: pass
            if word:
                try: word.Quit()
                except: pass
            if powerpoint:
                try: powerpoint.Quit()
                except: pass
            pythoncom.CoUninitialize()
            
            self.write_log_file(target_dir, list_success, list_fail)
            
            self.is_running = False
            self.root.after(0, lambda: self.btn_convert.config(state="normal", text="BẮT ĐẦU CHUYỂN ĐỔI"))
            self.root.after(0, lambda: messagebox.showinfo("Hoàn tất", f"Đã xong!\nThành công: {len(list_success)}\nThất bại: {len(list_fail)}\nLog đã được lưu tại thư mục quét."))

    def write_log_file(self, target_dir, successes, fails):
        """Hàm ghi file log.txt theo format yêu cầu"""
        log_path = os.path.join(target_dir, "conversion_log.txt")
        
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"BÁO CÁO CHUYỂN ĐỔI FILE OFFICE\n")
                f.write(f"Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Thư mục quét: {target_dir}\n")
                f.write("="*50 + "\n\n")

                f.write(f"THÀNH CÔNG: {len(successes)}\n")
                if successes:
                    f.writelines(f"{item}\n" for item in successes)
                else:
                    f.write("- Không có file nào được chuyển đổi thành công.\n")

                f.write("\n" + "-"*30 + "\n\n")

                f.write(f"THẤT BẠI: {len(fails)}\n")
                if fails:
                    f.writelines(f"{item}\n" for item in fails)
                else:
                    f.write("- Không có file nào thất bại.\n")
            
            self.log_gui("-" * 30)
            self.log_gui(f"Đã lưu file log chi tiết tại:\n{log_path}")
            
        except Exception as e:
            self.log_gui(f"Lỗi khi ghi file log: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeConverterApp(root)
    root.mainloop()