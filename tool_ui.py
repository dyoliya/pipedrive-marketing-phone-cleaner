import os
import sys
import threading
import subprocess
import io
import customtkinter as ctk
import tkinter as tk 
from tkinter import messagebox
from pd_marketing_cleaning_tool import main as cleaning_main

ctk.set_appearance_mode("dark")  # "dark" or "light"
ctk.set_default_color_theme("dark-blue")  # optional theme

class MinimalToolUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Pipedrive Marketing Cleaning Tool")
        self.geometry("620x480")
        self.configure(fg_color="#273946")  # dark charcoal background

        self.dots_running = False
        self.dots_count = 0

        self.input_folder = "for_processing"

        # Title label
        self.title_label = ctk.CTkLabel(self,
                                        text="Pipedrive Marketing Cleaning Tool",
                                        text_color="#fff6de",
                                        font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"))
        self.title_label.pack(pady=(20, 5))

        # Input folder label
        self.input_folder_label = ctk.CTkLabel(self,
                                              text=f"Input folder: {self.input_folder}",
                                              text_color="#BBB8A6",
                                              font=ctk.CTkFont(family="Segoe UI", size=12))
        self.input_folder_label.pack(pady=(0, 15))

        # Button frame
        btn_frame = ctk.CTkFrame(self, fg_color="#273946", corner_radius=0)
        btn_frame.pack(pady=5)

        # Buttons with rounded corners
        self.open_btn = ctk.CTkButton(btn_frame, text="Open Folder",
                                      fg_color="#CB1F47",
                                      hover_color="#ffab4c",
                                      command=self.open_input_folder)
        self.open_btn.pack(side="left", padx=5)

        self.refresh_btn = ctk.CTkButton(btn_frame, text="Refresh",
                                         fg_color="#CB1F47",
                                         hover_color="#ffab4c",
                                         command=self.load_file_list)
        self.refresh_btn.pack(side="left", padx=5)


        self.list_container = ctk.CTkFrame(self, fg_color="#273946", corner_radius=0)
        self.list_container.pack(padx=20, pady=10, fill="x")

        self.file_listbox = tk.Listbox(self.list_container,
                                    height=6,
                                    bg="#fff6de",
                                    fg="#273946",
                                    font=("Segoe UI", 11),
                                    highlightthickness=0,
                                    relief="flat",
                                    selectbackground="#CB1F47",
                                    selectforeground="white")
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)

        scrollbar = tk.Scrollbar(self.list_container, command=self.file_listbox.yview)
        scrollbar.pack(side="right", fill="y", padx=(0,5), pady=5)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # Instruction label
        self.instruction_label = ctk.CTkLabel(self,
                                             text="",
                                             text_color="#BBB8A6",
                                             wraplength=550,
                                             justify="center",
                                             font=ctk.CTkFont(family="Segoe UI", size=12))
        self.instruction_label.pack(pady=5)

        # Message label
        self.message_label = ctk.CTkLabel(self,
                                          text="Waiting to start...",
                                          text_color="#BBB8A6",
                                          font=ctk.CTkFont(family="Segoe UI", size=12))
        self.message_label.pack(pady=5)

        # Progress bar
        self.progress = ctk.CTkProgressBar(self, width=500)
        self.progress.set(0)
        self.progress.pack(pady=10)

        # Run button
        self.run_btn = ctk.CTkButton(self, text="RUN TOOL",
                                     width=120,
                                     fg_color="#CB1F47",
                                     hover_color="#ffab4c",
                                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                                     command=self.run_tool)
        self.run_btn.pack(pady=15)

        self.load_file_list()

    def open_input_folder(self):
        folder = self.input_folder
        if os.path.isdir(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])
        else:
            messagebox.showerror("Error", f"Input folder not found:\n{folder}")

    def load_file_list(self):
        folder = self.input_folder
        self.file_listbox.delete(0, "end")
        if os.path.isdir(folder):
            files = sorted(f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f)))
            if files:
                for f in files:
                    self.file_listbox.insert("end", f)
                self.update_message(f"{len(files)} file(s) found in '{folder}'")
                self.instruction_label.configure(
                    text="If correct, click Run Tool. Otherwise, update files in Input folder."
                )
            else:
                self.update_message(f"No files found in '{folder}'")
                self.instruction_label.configure(
                    text="Add the correct files to the Input folder, then click Refresh."
                )
        else:
            self.update_message(f"Input folder not found: {folder}")
            self.instruction_label.configure(text="")

    def open_folder(self, folder):
        if os.path.isdir(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])
        else:
            messagebox.showerror("Error", f"Folder not found:\n{folder}")

    def run_tool(self):
        folder = self.input_folder
        if not os.path.exists(folder):
            messagebox.showerror("Error", f"Input folder not found:\n{folder}")
            return
        self.message_label.configure(text=f"Running on folder: {folder} ...")
        self.progress.set(0)
        self.dots_running = True
        self.dots_count = 0
        self.run_btn.configure(state="disabled")
        self.animate_dots()
        self.show_wait_popup()
        threading.Thread(target=self.run_main_process, daemon=True).start()

    def show_wait_popup(self):
        # Create a top-level window
        self.wait_popup = ctk.CTkToplevel(self)
        self.wait_popup.title("Please Wait")
        self.wait_popup.geometry("300x100")
        self.wait_popup.resizable(False, False)
        self.wait_popup.transient(self)  # stay on top of parent
        self.wait_popup.grab_set()       # block interaction with main window

        # Center text label
        self.wait_label = ctk.CTkLabel(self.wait_popup,
                                       text="Processing",
                                       font=ctk.CTkFont(family="Segoe UI", size=14))
        self.wait_label.pack(expand=True, pady=20)

        # Animation control
        self.wait_dots_running = True
        self.wait_dots_count = 0
        self.animate_wait_popup()

    def animate_wait_popup(self):
        if not getattr(self, "wait_dots_running", False):
            return
        self.wait_dots_count = (self.wait_dots_count + 1) % 4
        dots = "." * self.wait_dots_count
        self.wait_label.configure(text=f"Processing{dots}")
        self.wait_label.after(500, self.animate_wait_popup)

    def close_wait_popup(self):
        if hasattr(self, "wait_popup"):
            self.wait_dots_running = False
            self.wait_popup.destroy()

    def animate_dots(self):
            if not self.dots_running:
                return
            self.dots_count = (self.dots_count + 1) % 4
            dots = "." * self.dots_count
            base_text = f"Running on folder: {self.input_folder}"
            self.message_label.configure(text=base_text + dots)
            self.message_label.after(500, self.animate_dots)

    def run_main_process(self):
        try:
            if sys.stdout is None:
                sys.stdout = io.StringIO()
            if sys.stderr is None:
                sys.stderr = io.StringIO()

            cleaning_main()
            self.dots_running = False
            self.close_wait_popup() 
            self.run_btn.configure(state="normal")
            self.update_message("Processing finished successfully!")
            self.progress.set(1.0)

            def ask_open_folder():
                if messagebox.askyesno("Done", "Processing finished!\nOpen output folder?"):
                    output_folder = os.path.abspath("output")
                    if not os.path.exists(output_folder):
                        os.makedirs(output_folder)
                    self.open_folder(output_folder)
            self.message_label.after(0, ask_open_folder)

        except Exception as e:
            self.close_wait_popup()
            self.run_btn.configure(state="normal")
            self.update_message(f"Failed to run tool:\n{e}")

    def update_message(self, text):
        self.message_label.after(0, lambda: self.message_label.configure(text=text))


if __name__ == "__main__":
    app = MinimalToolUI()
    app.mainloop()
