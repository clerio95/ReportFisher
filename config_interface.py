import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
import threading # Import threading
from robozamReports import start_bot_logic # Import the new function
import ctypes
import time
import pystray
from PIL import Image
import sys

CONFIG_FILE = "bot_config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            if "autosystem_path" not in config:
                config["autosystem_path"] = "C:\\autosystem\\main.exe"
            if "report_source_file" not in config:
                config["report_source_file"] = "C:\\autosystem\\relatorios\\relatorio.txt"
            if "report_dest_folder" not in config:
                config["report_dest_folder"] = "C:\\destino\\relatorios"
            return config
    return {"execution_frequency_minutes": 2, "autosystem_path": "C:\\autosystem\\main.exe", "report_source_file": "C:\\autosystem\\relatorios\\relatorio.txt", "report_dest_folder": "C:\\destino\\relatorios"}

def save_config(frequency, autosystem_path, report_source_file, report_dest_folder):
    config = {
        "execution_frequency_minutes": frequency,
        "autosystem_path": autosystem_path,
        "report_source_file": report_source_file,
        "report_dest_folder": report_dest_folder
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)
    messagebox.showinfo("Configuration Saved", f"Bot will now run every {frequency} minutes, use {autosystem_path} as AutoSystem executable, copy {report_source_file} to {report_dest_folder}.")

def prevent_sleep():
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_AWAYMODE_REQUIRED = 0x00000040  # Opcional, para Windows Media Center
    while True:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
        )
        time.sleep(30)  # Chama a cada 30 segundos

def hide_window(root):
    root.withdraw()

def show_window(root):
    root.deiconify()
    root.after(0, root.lift)

def on_tray_restore(icon, root):
    show_window(root)
    icon.stop()

def on_tray_exit(icon, root):
    icon.stop()
    root.quit()

def create_tray_icon(root):
    # Usa o mesmo ícone do executável
    try:
        image = Image.open('robot.ico')
    except Exception:
        # fallback: quadrado preto
        image = Image.new('RGB', (64, 64), color='black')
    menu = pystray.Menu(
        pystray.MenuItem('Restaurar', lambda: on_tray_restore(icon, root)),
        pystray.MenuItem('Sair', lambda: on_tray_exit(icon, root))
    )
    icon = pystray.Icon('Robozam', image, 'Robozam', menu)
    # pystray roda em thread separada
    threading.Thread(target=icon.run, daemon=True).start()
    return icon

def start_bot():
    current_config = load_config()
    try:
        # Thread para impedir o repouso
        sleep_thread = threading.Thread(target=prevent_sleep, daemon=True)
        sleep_thread.start()
        # Start the bot logic in a new thread to keep the GUI responsive
        bot_thread = threading.Thread(target=start_bot_logic, args=(current_config["execution_frequency_minutes"], current_config["autosystem_path"]), daemon=True)
        bot_thread.start()
        # Esconde a janela principal e mostra o tray
        hide_window(root)
        create_tray_icon(root)
        messagebox.showinfo("Bot Started", "O bot foi iniciado em background. O ícone está na bandeja do sistema.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start bot: {e}")

def create_interface():
    global root
    root = tk.Tk()
    root.title("Bot Configuration")

    current_config = load_config()
    
    tk.Label(root, text="Execution Frequency (minutes):").pack(pady=10)
    frequency_entry = tk.Entry(root)
    frequency_entry.insert(0, str(current_config["execution_frequency_minutes"]))
    frequency_entry.pack(pady=5)

    tk.Label(root, text="AutoSystem Executable Path:").pack(pady=10)
    autosystem_path_entry = tk.Entry(root, width=50)
    autosystem_path_entry.insert(0, current_config["autosystem_path"])
    autosystem_path_entry.pack(pady=5)

    def browse_autosystem_path():
        filepath = filedialog.askopenfilename(
            title="Select AutoSystem Executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filepath:
            autosystem_path_entry.delete(0, tk.END)
            autosystem_path_entry.insert(0, filepath)

    browse_button = tk.Button(root, text="Browse", command=browse_autosystem_path)
    browse_button.pack(pady=5)

    # Arquivo de origem
    tk.Label(root, text="Arquivo de origem do relatório (X):").pack(pady=10)
    report_source_file_entry = tk.Entry(root, width=50)
    report_source_file_entry.insert(0, current_config["report_source_file"])
    report_source_file_entry.pack(pady=5)
    def browse_report_source_file():
        file = filedialog.askopenfilename(title="Selecione o arquivo de origem do relatório (X)", filetypes=[("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")])
        if file:
            report_source_file_entry.delete(0, tk.END)
            report_source_file_entry.insert(0, file)
    browse_source_file_button = tk.Button(root, text="Selecionar Arquivo X", command=browse_report_source_file)
    browse_source_file_button.pack(pady=5)

    # Pasta de destino
    tk.Label(root, text="Pasta de destino dos relatórios (Y):").pack(pady=10)
    report_dest_entry = tk.Entry(root, width=50)
    report_dest_entry.insert(0, current_config["report_dest_folder"])
    report_dest_entry.pack(pady=5)
    def browse_report_dest():
        folder = filedialog.askdirectory(title="Selecione a pasta de destino dos relatórios (Y)")
        if folder:
            report_dest_entry.delete(0, tk.END)
            report_dest_entry.insert(0, folder)
    browse_dest_button = tk.Button(root, text="Selecionar Pasta Y", command=browse_report_dest)
    browse_dest_button.pack(pady=5)

    def on_save():
        try:
            frequency = int(frequency_entry.get())
            if frequency <= 0:
                raise ValueError("Frequency must be a positive integer.")
            autosystem_path = autosystem_path_entry.get()
            report_source_file = report_source_file_entry.get()
            report_dest_folder = report_dest_entry.get()
            save_config(frequency, autosystem_path, report_source_file, report_dest_folder)
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))

    save_button = tk.Button(root, text="Save Configuration", command=on_save)
    save_button.pack(pady=10)

    start_button = tk.Button(root, text="Start Bot", command=start_bot)
    start_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_interface() 