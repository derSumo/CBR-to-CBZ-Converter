import os
import threading
import zipfile
import rarfile
import shutil
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import END

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class CBRtoCBZConverter(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CBR zu CBZ Konverter")
        self.geometry("850x550")
        self.resizable(False, False)
        self.files = []
        self.stop_flag = False
        self.output_dir = ""

        self.init_ui()

    def init_ui(self):
        # Eingabeordner auswählen
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(padx=20, pady=(20, 10), fill="x")

        self.path_entry = ctk.CTkEntry(input_frame, placeholder_text="Ordner mit CBR-Dateien auswählen...")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)

        browse_btn = ctk.CTkButton(input_frame, text="Durchsuchen", command=self.select_folder)
        browse_btn.pack(side="right", padx=(5, 10))

        # Speicherort auswählen
        output_frame = ctk.CTkFrame(self)
        output_frame.pack(padx=20, pady=(0, 10), fill="x")

        self.output_entry = ctk.CTkEntry(output_frame, placeholder_text="Speicherort für CBZ-Dateien...")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)

        output_btn = ctk.CTkButton(output_frame, text="Speicherort", command=self.select_output_folder)
        output_btn.pack(side="right", padx=(5, 10))

        # Dateiliste
        self.file_listbox = ctk.CTkTextbox(self, height=200)
        self.file_listbox.pack(padx=20, pady=(10, 5), fill="both", expand=False)
        self.file_listbox.configure(state="disabled")
        self.file_listbox.bind("<Delete>", self.remove_selected_line)

        # Fortschritt
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(padx=20, pady=10, fill="x")
        self.progress_bar.set(0)

        # Buttons
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(padx=20, pady=10)

        self.start_btn = ctk.CTkButton(btn_frame, text="Start", command=self.start_conversion)
        self.start_btn.grid(row=0, column=0, padx=10)

        self.stop_btn = ctk.CTkButton(btn_frame, text="Stopp", command=self.stop_conversion, fg_color="red")
        self.stop_btn.grid(row=0, column=1, padx=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.path_entry.delete(0, END)
            self.path_entry.insert(0, folder)
            self.scan_folder(folder)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, folder)
            self.output_dir = folder

    def scan_folder(self, folder):
        self.files.clear()
        self.file_listbox.configure(state="normal")
        self.file_listbox.delete("1.0", END)

        for file in os.listdir(folder):
            if file.lower().endswith(".cbr"):
                full_path = os.path.join(folder, file)
                size_mb = os.path.getsize(full_path) / (1024 * 1024)
                self.files.append(full_path)
                self.file_listbox.insert(END, f"{file} | {size_mb:.2f} MB\n")

        self.file_listbox.configure(state="disabled")

    def remove_selected_line(self, event=None):
        try:
            index = int(self.file_listbox.index("insert").split('.')[0]) - 1
            if 0 <= index < len(self.files):
                del self.files[index]
                self.scan_folder(self.path_entry.get())
        except:
            pass

    def convert_cbr_to_cbz(self, cbr_path, cbz_path):
        temp_dir = os.path.join(os.path.dirname(cbr_path), "__temp_extract__")
        os.makedirs(temp_dir, exist_ok=True)

        try:
            with rarfile.RarFile(cbr_path) as rf:
                rf.extractall(temp_dir)

            with zipfile.ZipFile(cbz_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        rel_path = os.path.relpath(file_path, temp_dir)
                        zf.write(file_path, arcname=rel_path)
        except rarfile.NeedFirstVolume:
            raise Exception("CBR ist mehrteilig und benötigt das erste Volume.")
        except rarfile.RarCannotExec:
            raise Exception("Kein funktionierendes UNRAR-Tool gefunden.\nBitte `unrar` installieren oder PATH prüfen.")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def run_conversion(self):
        total = len(self.files)
        if total == 0:
            messagebox.showinfo("Hinweis", "Keine CBR-Dateien gefunden.")
            return

        for idx, cbr_file in enumerate(self.files):
            if self.stop_flag:
                break

            filename = os.path.basename(cbr_file)
            cbz_file = os.path.join(self.output_dir, filename[:-4] + ".cbz")

            try:
                self.convert_cbr_to_cbz(cbr_file, cbz_file)
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler bei Datei: {filename}\n{e}")
                continue

            self.progress_bar.set((idx + 1) / total)

        messagebox.showinfo("Fertig", "Konvertierung abgeschlossen!")
        self.progress_bar.set(0)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("Achtung", "Bitte zuerst einen Ordner mit CBR-Dateien auswählen.")
            return

        self.output_dir = self.output_entry.get()
        if not self.output_dir or not os.path.isdir(self.output_dir):
            messagebox.showwarning("Achtung", "Bitte gültigen Speicherort auswählen.")
            return

        self.stop_flag = False
        threading.Thread(target=self.run_conversion).start()

    def stop_conversion(self):
        self.stop_flag = True
        messagebox.showinfo("Stopp", "Konvertierung wird abgebrochen...")

if __name__ == "__main__":
    app = CBRtoCBZConverter()
    app.mainloop()
