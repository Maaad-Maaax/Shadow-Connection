import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import time
import ctypes
import sys

def hide_console():
    if sys.platform == "win32":
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

hide_console()

class RDPConnector:
    def __init__(self, root):
        self.root = root
        self.root.title("Теневое подключение")
        self.root.iconbitmap("favico.ico") 
        self.root.overrideredirect(False)
        self.setup_ui()
        self.center_window(self.root)
        self.root.state('zoomed')
        self.root.bind("<F11>", self.toggle_fullscreen)
        self.root.bind("<Escape>", self.exit_fullscreen)
        self.fullscreen = True
        self.root.attributes("-alpha", 0.95)
        self.polling_thread = None
        self.stop_polling = False
        self.start_device_polling()

    def toggle_fullscreen(self, event=None):
        self.fullscreen = not self.fullscreen
        self.root.attributes("-fullscreen", self.fullscreen)

    def exit_fullscreen(self, event=None):
        self.fullscreen = False
        self.root.attributes("-fullscreen", False)

    def center_window(self, window):
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_ui(self):
        style = ttk.Style()
        style.configure("Bold.TLabelframe.Label", font=('Arial', 12, 'bold'), foreground='white', background='black')
        style.configure("Bold.TLabelframe", background='black', borderwidth=2, relief='solid')
        style.configure("TScrollbar", background="black", troughcolor="black", bordercolor="black", arrowcolor="white")
        style.map("TScrollbar", background=[("active", "gray")])

        self.canvas = tk.Canvas(self.root, bg='black', highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.canvas.yview, style="TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.content_frame = ttk.Frame(self.canvas, style="Bold.TLabelframe")
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        departments = {}
        for comp in computers:
            dept = comp['department']
            departments.setdefault(dept, []).append(comp)

        sorted_depts = sorted(departments.keys())

        self.rows = []
        current_row = ttk.Frame(self.content_frame, style="Bold.TLabelframe")
        current_row.pack(fill=tk.X, padx=5, pady=5)
        self.rows.append(current_row)

        for idx, dept in enumerate(sorted_depts):
            if idx > 0 and idx % 5 == 0:
                current_row = ttk.Frame(self.content_frame, style="Bold.TLabelframe")
                current_row.pack(fill=tk.X, padx=5, pady=5)
                self.rows.append(current_row)
            
            self.add_department(current_row, dept, departments[dept])

        self.root.update_idletasks()
        self.center_window(self.root)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_department(self, parent_frame, dept_name, comp_list):
        frame = ttk.LabelFrame(
            parent_frame,
            text=dept_name,
            style="Bold.TLabelframe"
        )
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        for comp in comp_list:
            button_text = f"{comp['name']}\n----------------------\n{comp['fio']}"

            btn = tk.Button(
                frame,
                text=button_text,
                command=lambda c=comp: self.connect(c['name']),
                justify=tk.CENTER,
                anchor='center',
                width=42,
                wraplength=300,
                cursor="hand2",
                bg='black',
                fg='white',
                highlightbackground='white',
                highlightthickness=2,
                font=('Arial', 10, 'bold')
            )
            btn.pack(padx=7, pady=10, fill=tk.X)
            comp['button'] = btn

    def start_device_polling(self):
        def poll_devices():
            while not self.stop_polling:
                threads = []
                for comp in computers:
                    thread = threading.Thread(target=self.check_device_availability, args=(comp,))
                    threads.append(thread)
                    thread.start()

                for thread in threads:
                    thread.join()

                time.sleep(10)  # Опрос каждые 10 секунд

        self.polling_thread = threading.Thread(target=poll_devices, daemon=True)
        self.polling_thread.start()

    def check_device_availability(self, comp):
        try:
            result = subprocess.run(
                ['quser.exe', f'/server:{comp["name"]}'],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=3
            )
            
            if result.returncode == 0:
                sessions = self.parse_sessions(result.stdout)
                is_available = len(sessions) > 0
            else:
                is_available = False
        except subprocess.TimeoutExpired:
            is_available = False
        except Exception as e:
            print(f"Ошибка при проверке сессий на устройстве {comp['name']}: {e}")
            is_available = False

        self.root.after(0, self.update_button_color, comp['button'], is_available)

    def update_button_color(self, button, is_available):
        color = 'green' if is_available else 'red'
        button.config(bg=color)

    def connect(self, pc_name):
        try:
            result = subprocess.run(
                ['quser.exe', f'/server:{pc_name}'],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=5
            )
            
            if result.returncode == 0:
                sessions = self.parse_sessions(result.stdout)
                if sessions:
                    session_id = sessions[0]['session_id']
                    self.show_mode_dialog(pc_name, session_id)
                else:
                    messagebox.showinfo("Информация", "Отсутствует активная сессия пользователя, либо ПК не включен!")
            else:
                messagebox.showinfo("Информация", "Отсутствует активная сессия пользователя, либо ПК не включен!")
                
        except subprocess.TimeoutExpired:
            messagebox.showinfo("Информация", "Отсутствует активная сессия пользователя, либо ПК не включен!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться:\n{e}")
    
    def parse_sessions(self, output):
        lines = output.split('\n')
        if len(lines) < 2:
            return []
        
        sessions = []
        for line in lines[1:]:
            if line.strip():
                parts = line.split()
                if len(parts) >= 3:
                    sessions.append({
                        'username': parts[0],
                        'session_id': parts[2],
                        'state': parts[3]
                    })
        return sessions
    
    def show_mode_dialog(self, pc_name, session_id):
        dialog = tk.Toplevel(self.root)
        dialog.title("Выбор режима подключения")
        dialog.geometry("400x150")
        self.center_window(dialog)
        dialog.attributes("-alpha", 0.95)
        dialog.configure(bg="black")

        style = ttk.Style()
        style.configure("Dialog.TLabel", background="black", foreground="white", font=('Arial', 12))
        style.configure("Dialog.TRadiobutton", background="black", foreground="white", font=('Arial', 12))
        style.configure("Dialog.TButton", cursor="hand2", background="black", foreground="black", font=('Arial', 12))

        ttk.Label(dialog, text=f"Компьютер: {pc_name}", style="Dialog.TLabel").pack(pady=5)
        ttk.Label(dialog, text="Выберите режим подключения:", style="Dialog.TLabel").pack()
        
        mode_var = tk.IntVar(value=1)
        
        ttk.Radiobutton(dialog, text="Только просмотр", variable=mode_var, value=1, style="Dialog.TRadiobutton").pack()
        ttk.Radiobutton(dialog, text="Полный доступ", variable=mode_var, value=2, style="Dialog.TRadiobutton").pack()
        
        ttk.Button(
            dialog,
            text="Подключиться",
            command=lambda: self.start_rdp(pc_name, session_id, mode_var.get(), dialog),
            cursor="hand2",
            style="Dialog.TButton"
        ).pack(pady=10)
    
    def start_rdp(self, pc_name, session_id, mode, dialog):
        control = "/control" if mode == 2 else ""
        command = [
            'mstsc.exe',
            f'/shadow:{session_id}',
            f'/v:{pc_name}',
            control,
            '/noConsentPrompt'
        ]
        
        try:
            subprocess.Popen([c for c in command if c])
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка запуска RDP:\n{e}")

if __name__ == "__main__":
    computers = []

    with open('data.txt', 'r', encoding='utf-8') as file:
        for line in file:
            fio, name, department = line.strip().split(',')
            computers.append({'fio': fio, 'name': name, 'department': department})
    root = tk.Tk()
    app = RDPConnector(root)
    root.mainloop()