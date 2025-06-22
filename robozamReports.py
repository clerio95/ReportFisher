import datetime
import time
import subprocess
from pywinauto import Application, Desktop, keyboard
from pywinauto.findwindows import ElementNotFoundError
import io
import sys
import os
import psutil
import logging
import re
import schedule
import threading
import pystray
from PIL import Image
import json
import glob
import unicodedata
import shutil

# Configura logging (sem arquivo)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Estado do bot
bot_running = True

# Evento para sinalizar o encerramento da thread de monitoramento do Bloco de Notas
stop_notepad_monitor_event = threading.Event()

# Parâmetros centralizados para fácil ajuste
MONITOR_START_DELAY = 8  # segundos após início do main.exe
MONITOR_TIME = 30        # duração do monitoramento em segundos
MONITOR_POLL_INTERVAL = 0.1  # intervalo de polling em segundos

def notepad_monitor_thread():
    logging.info("Iniciando thread de monitoramento do Bloco de Notas.")
    while not stop_notepad_monitor_event.is_set():
        close_notepad()
        time.sleep(0.5) # Verifica a cada 0.5 segundos
    logging.info("Thread de monitoramento do Bloco de Notas encerrada.")

def close_notepad():
    logging.info("Verificando janelas do Bloco de Notas...")
    notepad_timeout = 15
    notepad_start_time = time.time()
    app = Application(backend="win32")
    notepad_closed = False

    while time.time() - notepad_start_time < notepad_timeout:
        if stop_notepad_monitor_event.is_set():
            logging.info("Monitoramento de Notepad interrompido após sequência de fechamento.")
            break
        try:
            desktop = Desktop(backend="win32")
            windows = [w for w in desktop.windows() if re.search(r"relat[oó]rio.*Bloco de Notas", w.window_text(), re.IGNORECASE) and w.is_visible()]
            if windows:
                for notepad_window in windows:
                    if stop_notepad_monitor_event.is_set():
                        logging.info("Monitoramento de Notepad interrompido após sequência de fechamento.")
                        break
                    logging.info(f"Janela do Bloco de Notas encontrada: Título: {notepad_window.window_text()}, Handle: {notepad_window.handle}")
                    try:
                        app.connect(handle=notepad_window.handle)
                        notepad_dialog = app.window(handle=notepad_window.handle)
                        notepad_dialog.wait('enabled ready visible', timeout=3)
                        notepad_dialog.set_focus()
                        time.sleep(0.2)
                        logging.info("Enviando ALT+F4 para fechar a janela...")
                        keyboard.send_keys("%{F4}", pause=0.1)
                        time.sleep(0.2)
                        confirm_timeout = 5
                        confirm_start_time = time.time()
                        while time.time() - confirm_start_time < confirm_timeout:
                            if stop_notepad_monitor_event.is_set():
                                logging.info("Monitoramento de Notepad interrompido após sequência de fechamento.")
                                break
                            desktop = Desktop(backend="win32")
                            confirm_windows = [w for w in desktop.windows() if "Bloco de Notas" in w.window_text() and any(s in w.window_text() for s in ["Salvar", "Salvar alterações", "Save", "Untitled - Notepad"]) and w.is_visible()]
                            if confirm_windows:
                                confirm_window = confirm_windows[0]
                                logging.info(f"Pop-up de confirmação encontrado: Título: {confirm_window.window_text()}, Handle: {confirm_window.handle}")
                                try:
                                    app.connect(handle=confirm_window.handle)
                                    confirm_dialog = app.window(handle=confirm_window.handle)
                                    confirm_dialog.wait('enabled ready visible', timeout=3)
                                    confirm_dialog.set_focus()
                                    time.sleep(0.2)
                                    logging.info("Listando controles do pop-up de confirmação do Bloco de Notas:")
                                    for i, control in enumerate(confirm_dialog.children()):
                                        logging.info(f"  Control {i}: Title={repr(control.window_text())}, ControlType={repr(control.control_type)}, ClassName={repr(control.class_name())}, AutomationId={repr(getattr(control, 'automation_id', 'N/A'))}")
                                    button_clicked = False
                                    button_titles = ["Não Salvar", "Don't Save", "Do not save", "Não", "No", "Cancel"]
                                    for btn_title in button_titles:
                                        try:
                                            if confirm_dialog.child_window(title=btn_title, control_type="Button").exists():
                                                confirm_dialog[btn_title].click()
                                                logging.info(f"Clicado '{btn_title}' no pop-up de confirmação.")
                                                button_clicked = True
                                                break
                                            else:
                                                logging.info(f"Botão '{btn_title}' não encontrado.")
                                        except ElementNotFoundError:
                                            logging.info(f"ElementNotFoundError for button '{btn_title}'.")
                                        except Exception as e_btn:
                                            logging.error(f"Erro ao tentar clicar no botão '{btn_title}': {e_btn}")
                                    if not button_clicked:
                                        logging.info("Nenhum botão de 'Não Salvar' ou 'Don\'t Save' encontrado, enviando ALT+N como fallback.")
                                        keyboard.send_keys("%N", pause=0.2)
                                        time.sleep(0.2)
                                        if confirm_window.exists():
                                            logging.warning("Pop-up de confirmação ainda visível após ALT+N, tentando ESC.")
                                            keyboard.send_keys("{ESC}", pause=0.2)
                                            time.sleep(0.2)
                                            if confirm_window.exists():
                                                logging.warning("Pop-up de confirmação ainda visível após ESC, tentando ALT+F4.")
                                                keyboard.send_keys("%{F4}", pause=0.2)
                                                time.sleep(0.2)
                                    dialog_closed_check_start_time = time.time()
                                    while time.time() - dialog_closed_check_start_time < 2:
                                        if stop_notepad_monitor_event.is_set():
                                            logging.info("Monitoramento de Notepad interrompido após sequência de fechamento.")
                                            break
                                        if not confirm_window.exists():
                                            logging.info("Pop-up de confirmação fechado com sucesso!")
                                            break
                                        time.sleep(0.1)
                                    else:
                                        logging.warning("Pop-up de confirmação ainda visível após tentativa de fechamento.")
                                except Exception as ex:
                                    logging.error(f"Erro ao interagir com pop-up de confirmação do Bloco de Notas: {ex}")
                            time.sleep(0.1)
                        windows = [w for w in desktop.windows() if w.handle == notepad_window.handle and w.is_visible()]
                        if not windows:
                            logging.info("Janela do Bloco de Notas fechada com sucesso!")
                            notepad_closed = True
                        else:
                            logging.warning("Janela do Bloco de Notas ainda aberta, tentando ESC para fechar...")
                            keyboard.send_keys("{ESC}", pause=0.2)
                            time.sleep(0.2)
                            windows = [w for w in desktop.windows() if w.handle == notepad_window.handle and w.is_visible()]
                            if not windows:
                                logging.info("Janela do Bloco de Notas fechada com sucesso após ESC!")
                                notepad_closed = True
                            else:
                                logging.warning("Janela do Bloco de Notas ainda aberta após ESC.")
                                keyboard.send_keys("%{F4}", pause=0.2)
                                time.sleep(0.2)
                    except Exception as e:
                        logging.error(f"Erro ao fechar janela do Bloco de Notas: {e}")
                if notepad_closed:
                    break
            else:
                logging.info("Nenhuma janela do Bloco de Notas encontrada, prosseguindo...")
                notepad_closed = True
                break
            time.sleep(0.1)
        except Exception as e:
            logging.error(f"Erro ao verificar janelas do Bloco de Notas: {e}")
            time.sleep(0.1)
    if not notepad_closed:
        logging.warning("Janela do Bloco de Notas não foi fechada. Forçando encerramento de notepad.exe...")
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == "notepad.exe":
                proc.terminate()
                proc.wait(timeout=2)
                logging.info(f"Processo notepad.exe (PID: {proc.pid}) encerrado.")
        except psutil.NoSuchProcess:
            pass
        except Exception as e:
            logging.error(f"Erro ao encerrar notepad.exe (PID: {proc.pid}): {e}")
    return notepad_closed

def terminate_autosystem(process):
    logging.info("Encerrando AutoSystem...")
    try:
        parent = psutil.Process(process.pid)
        children = parent.children(recursive=True)
        for child in children:
            try:
                child.terminate()
                child.wait(timeout=2)
                logging.info(f"Processo filho (PID: {child.pid}) encerrado.")
            except psutil.NoSuchProcess:
                pass
            except Exception as e:
                logging.error(f"Erro ao encerrar processo filho (PID: {child.pid}): {e}")
        parent.terminate()
        parent.wait(timeout=2)
        logging.info(f"Processo AutoSystem (PID: {process.pid}) encerrado.")
    except psutil.NoSuchProcess:
        logging.info("Processo AutoSystem já encerrado.")
    except Exception as e:
        logging.error(f"Erro ao encerrar AutoSystem: {e}")

def send_close_sequence_when_report_appears(timeout=40):
    def _send():
        logging.info(f"Aguardando até {timeout}s pela janela 'Visualizar Impressão' para fechar o relatório...")
        start = time.time()
        while time.time() - start < timeout:
            try:
                desktop = Desktop(backend="win32")
                windows = [w for w in desktop.windows() if "Visualizar Impressão" in w.window_text() and w.is_visible()]
                if windows:
                    logging.info("Janela 'Visualizar Impressão' detectada! Enviando sequência ESC ESC ESC ENTER para fechar o relatório...")
                    keyboard.send_keys("{ESC}", pause=0.1)
                    keyboard.send_keys("{ESC}", pause=0.1)
                    keyboard.send_keys("{ESC}", pause=0.1)
                    keyboard.send_keys("{ENTER}", pause=0.1)
                    return
            except Exception:
                pass
            time.sleep(0.2)
        logging.error("Janela 'Visualizar Impressão' não apareceu dentro do tempo limite.")
    threading.Thread(target=_send, daemon=True).start()

def aggressive_close_notepad_on_txt(report_path=r"C:\autosystem\relatorios", poll_interval=0.05, timeout=22, close_window_time=10):
    def _monitor():
        logging.info(f"Monitorando surgimento de arquivos .txt em {report_path} por até {timeout}s...")
        start = time.time()
        seen = set()
        if os.path.exists(report_path):
            seen = set(glob.glob(os.path.join(report_path, '*.txt')))
        while time.time() - start < timeout:
            try:
                if not os.path.exists(report_path):
                    time.sleep(poll_interval)
                    continue
                current = set(glob.glob(os.path.join(report_path, '*.txt')))
                new_files = current - seen
                if new_files:
                    for f in new_files:
                        logging.info(f"Arquivo txt detectado: {f}. Tentando fechar o Notepad agressivamente por {close_window_time}s...")
                        close_start = time.time()
                        while time.time() - close_start < close_window_time:
                            close_notepad()
                            time.sleep(poll_interval)
                    seen = current
            except Exception as e:
                logging.error(f"Erro ao monitorar arquivos txt: {e}")
            time.sleep(poll_interval)
        logging.info("Monitoramento de arquivos txt encerrado.")
    threading.Thread(target=_monitor, daemon=True).start()

def send_delayed_close_sequence(delay, log_prefix=""):
    def _send():
        if log_prefix:
            logging.info(f"{log_prefix}Aguardando {delay} segundos para enviar sequência de fechamento...")
        else:
            logging.info(f"Aguardando {delay} segundos para enviar sequência de fechamento...")
        time.sleep(delay)
        logging.info("Enviando sequência fixa RIGHT ENTER ESC ESC ESC ENTER após delay...")
        keyboard.send_keys("{RIGHT}", pause=0.1)
        keyboard.send_keys("{ENTER}", pause=0.1)
        keyboard.send_keys("{ESC}", pause=0.1)
        keyboard.send_keys("{ESC}", pause=0.1)
        keyboard.send_keys("{ESC}", pause=0.1)
        keyboard.send_keys("{ENTER}", pause=0.1)
        time.sleep(0.3)
        stop_notepad_monitor_event.set()  # Sinaliza para parar o monitoramento
    t = threading.Thread(target=_send, daemon=True)
    t.start()

def aggressive_monitor_and_close_notepad(start_delay=8, monitor_time=30, poll_interval=0.1):
    def _monitor():
        logging.info(f"Aguardando {start_delay}s após início do main.exe para iniciar monitoramento agressivo de Notepad em todo o Windows...")
        time.sleep(start_delay)
        start = time.time()
        while time.time() - start < monitor_time:
            try:
                import psutil
                for proc in psutil.process_iter(['pid', 'name']):
                    if proc.name().lower() == "notepad.exe":
                        proc.terminate()
                        proc.wait(timeout=2)
                        logging.info(f"Processo notepad.exe (PID: {proc.pid}) encerrado.")
            except Exception as e:
                logging.error(f"Erro ao tentar fechar o Notepad: {e}")
            time.sleep(poll_interval)
        logging.info("Monitoramento agressivo de Notepad encerrado.")
    import threading
    threading.Thread(target=_monitor, daemon=True).start()

def smart_monitor_and_close_notepad_with_relatorio(start_delay=MONITOR_START_DELAY, monitor_time=MONITOR_TIME, poll_interval=MONITOR_POLL_INTERVAL):
    import threading
    def normalize(text):
        return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII').lower()
    def _monitor():
        logging.info(f"Aguardando {start_delay}s após início do main.exe para iniciar monitoramento inteligente de Notepad com 'relatorio'...")
        time.sleep(start_delay)
        start = time.time()
        closed_handles = set()
        while time.time() - start < monitor_time:
            if stop_notepad_monitor_event.is_set():
                logging.info("Monitoramento de Notepad interrompido após sequência de fechamento.")
                break
            try:
                desktop = Desktop(backend="win32")
                for w in desktop.windows():
                    try:
                        title = normalize(w.window_text())
                        class_name = w.class_name()
                        handle = w.handle
                        if class_name and "notepad" in class_name.lower() and "relatorio" in title and "bloco de notas" in title:
                            if handle in closed_handles:
                                continue
                            logging.info(f"Fechando Notepad com título: {w.window_text()} (class: {class_name})")
                            try:
                                w.close()
                                time.sleep(0.2)
                                try:
                                    _ = w.exists(timeout=1)
                                except Exception:
                                    pass
                                try:
                                    import psutil
                                    proc = psutil.Process(w.process_id())
                                    proc.terminate()
                                    proc.wait(timeout=2)
                                    logging.info(f"Processo notepad.exe (PID: {proc.pid}) encerrado à força.")
                                except Exception as e:
                                    logging.error(f"Erro ao tentar forçar o término do Notepad: {type(e).__name__}: {e}")
                                closed_handles.add(handle)
                            except Exception as e:
                                logging.error(f"Erro ao tentar fechar a janela do Notepad: {type(e).__name__}: {e}")
                    except Exception as e:
                        logging.error(f"Erro ao processar janela do Notepad: {type(e).__name__}: {e}")
            except Exception as e:
                logging.error(f"Erro ao monitorar janelas do Notepad: {type(e).__name__}: {e}")
            time.sleep(poll_interval)
        logging.info("Monitoramento inteligente de Notepad encerrado.")
    threading.Thread(target=_monitor, daemon=True).start()

def exportar_relatorio(autosystem_path_arg):
    global bot_running
    if not bot_running:
        logging.info("Bot pausado. Pulando execução.")
        return

    # Carregar configuração para obter o arquivo de origem e pasta de destino
    CONFIG_FILE = "bot_config.json"
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            report_source_file = config.get("report_source_file", r"C:\\autosystem\\relatorios\\relatorio.txt")
            report_dest_folder = config.get("report_dest_folder", r"C:\\destino\\relatorios")
    else:
        report_source_file = r"C:\\autosystem\\relatorios\\relatorio.txt"
        report_dest_folder = r"C:\\destino\\relatorios"

    logging.info("Iniciando o aplicativo...")

    # Pre-run cleanup: Ensure no lingering notepad or autosystem processes
    logging.info("Executando limpeza pré-execução...")
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == "notepad.exe":
                proc.terminate()
                proc.wait(timeout=5)
                logging.info(f"Processo notepad.exe (PID: {proc.pid}) encerrado para limpeza prévia.")
            elif proc.name().lower() == "main.exe": # AutoSystem process
                proc.terminate()
                proc.wait(timeout=5)
                logging.info(f"Processo main.exe (AutoSystem) (PID: {proc.pid}) encerrado para limpeza prévia.")
        except psutil.NoSuchProcess:
            pass
        except Exception as e:
            logging.error(f"Erro durante a limpeza pré-execução para PID {proc.pid}: {e}")

    process = None # Initialize process to None
    notepad_thread = None
    try:
        # Use the provided autosystem_path_arg
        process = subprocess.Popen(autosystem_path_arg)
        logging.info(f"Processo iniciado com PID: {process.pid}")

        stop_notepad_monitor_event.clear()  # Limpa o evento no início de cada execução
        # Após 8 segundos do início do main.exe, monitora e fecha Notepad com 'relatorio' no título por 30 segundos
        smart_monitor_and_close_notepad_with_relatorio()

        # Inicia monitoramento agressivo para fechar o relatório assim que um txt surgir
        aggressive_close_notepad_on_txt(timeout=22)

        # Dispara a thread para enviar a sequência após 24 segundos
        send_delayed_close_sequence(18)

        max_startup_timeout = 45
        start_time = time.time()
        app = Application(backend="win32")
        initial_window_found = False

        logging.info("Aguardando o aplicativo abrir (pode levar até 45 segundos)...")
        while time.time() - start_time < max_startup_timeout:
            try:
                desktop = Desktop(backend="win32")
                windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible()]
                if windows:
                    logging.info("Janelas encontradas para o processo:")
                    for w in windows:
                        logging.info(f" - Título: {w.window_text()}, Handle: {w.handle}, PID: {w.process_id()}")
                    initial_window_found = True
                    break
            except Exception as e:
                logging.error(f"Verificando janelas... Erro: {e}")
            time.sleep(1)

        if not initial_window_found:
            logging.error(f"Nenhuma janela encontrada para o processo {process.pid} em {max_startup_timeout} segundos.")
            return False

        logging.info("Tentando conectar à janela de login...")
        while time.time() - start_time < max_startup_timeout:
            try:
                app.connect(title_re=".*Linx Postos Gerencial.*", process=process.pid)
                logging.info("Conexão à janela de login bem-sucedida!")
                break
            except Exception as e:
                logging.error(f"Tentando conectar à janela de login... Erro: {e}")
                time.sleep(1)
        else:
            logging.error("Não foi possível conectar à janela de login dentro do tempo limite.")
            return False

        try:
            login_window = app.window(title_re=".*Linx Postos Gerencial.*")
            login_window.wait('enabled ready visible', timeout=15)
            logging.info("Janela de login encontrada e pronta!")
        except Exception as e:
            logging.error(f"Erro ao acessar a janela de login: {e}")
            return False

        login_window.set_focus()
        time.sleep(0.5)

        try:
            logging.info("Enviando 'cf'...")
            keyboard.send_keys("cf", pause=0.1)
            time.sleep(0.2)
            logging.info("Enviando TAB...")
            keyboard.send_keys("{TAB}", pause=0.1)
            time.sleep(0.2)
            logging.info("Enviando '7418'...")
            keyboard.send_keys("7418", pause=0.1)
            time.sleep(0.2)
            logging.info("Enviando F5...")
            keyboard.send_keys("{F5}", pause=0.1)
        except Exception as e:
            logging.error(f"Erro ao enviar teclas: {e}")
            return False

        logging.info("Verificando a janela 'Questão'...")
        questao_timeout = 30
        questao_start_time = time.time()
        questao_window = None

        while time.time() - questao_start_time < questao_timeout:
            try:
                desktop = Desktop(backend="win32")
                windows = [w for w in desktop.windows() if w.process_id() == process.pid and "Questão" in w.window_text() and w.is_visible()]
                if windows:
                    questao_window = windows[0]
                    logging.info(f"Janela 'Questão' encontrada: Título: {questao_window.window_text()}, Handle: {questao_window.handle}")
                    break
                logging.info("Procurando janela 'Questão'...")
                time.sleep(1)
            except Exception as e:
                logging.error(f"Erro ao procurar janela 'Questão': {e}")
                time.sleep(1)

        if questao_window:
            try:
                app.connect(handle=questao_window.handle)
                questao_dialog = app.window(handle=questao_window.handle)
                questao_dialog.wait('enabled ready visible', timeout=5)
                questao_dialog.set_focus()
                time.sleep(0.2)
                logging.info("Enviando ESC para fechar a janela 'Questão'...")
                keyboard.send_keys("{ESC}", pause=0.1)
                time.sleep(1)
                close_timeout = 5
                close_start_time = time.time()
                while time.time() - close_start_time < close_timeout:
                    desktop = Desktop(backend="win32")
                    windows = [w for w in desktop.windows() if w.process_id() == process.pid and "Questão" in w.window_text() and w.is_visible()]
                    if not windows:
                        logging.info("Janela 'Questão' fechada com sucesso!")
                        break
                    logging.info("Janela 'Questão' ainda visível, aguardando fechamento...")
                    time.sleep(1)
                else:
                    logging.error("Janela 'Questão' não foi fechada. Tentando clicar em botão 'Não'...")
                    try:
                        questao_dialog['Não'].click()
                        time.sleep(1)
                        windows = [w for w in desktop.windows() if w.process_id() == process.pid and "Questão" in w.window_text() and w.is_visible()]
                        if not windows:
                            logging.info("Janela 'Questão' fechada com sucesso após clicar em 'Não'!")
                        else:
                            logging.error("Janela 'Questão' ainda não foi fechada.")
                            return False
                    except Exception as e:
                        logging.error(f"Erro ao tentar clicar no botão 'Não': {e}")
                        return False
            except Exception as e:
                logging.error(f"Erro ao acessar ou fechar a janela 'Questão': {e}")
                return False
        else:
            logging.info("Janela 'Questão' não encontrada dentro do tempo limite de 30 segundos. Prosseguindo...")

        logging.info("Tentando conectar à janela principal...")
        while time.time() - start_time < max_startup_timeout:
            try:
                app.connect(title_re=".*AutoSystem PRO.*", process=process.pid)
                logging.info("Conexão à janela principal bem-sucedida!")
                break
            except Exception as e:
                logging.error(f"Tentando conectar à janela principal... Erro: {e}")
                time.sleep(1)
        else:
            logging.error("Não foi possível conectar à janela principal dentro do tempo limite.")
            return False

        try:
            main_window = app.window(title_re=".*AutoSystem PRO.*")
            main_window.wait('enabled ready visible', timeout=15)
            logging.info("Janela principal encontrada e pronta!")
        except Exception as e:
            logging.error(f"Erro ao acessar a janela principal: {e}")
            return False

        main_window.set_focus()
        time.sleep(0.5)

        logging.info("Abrindo menu 'Financeiro' via ALT + F...")
        keyboard.send_keys("%F", pause=0.1)

        logging.info("Navegando até 'Relatórios' e depois 'Produtividade por Funcionários'...")

        try:
            keyboard.send_keys("{UP}", pause=0.1)
            keyboard.send_keys("{RIGHT}", pause=0.1)
            for _ in range(5):
                keyboard.send_keys("{UP}", pause=0.1)
            keyboard.send_keys("{ENTER}", pause=0.1)
            logging.info("Opção 'Produtividade por Funcionários' acionada.")
            keyboard.send_keys("{1}", pause=0.1)
            for _ in range(5):
                keyboard.send_keys("{DOWN}", pause=0.1)
            keyboard.send_keys("{1}", pause=0.1)
            keyboard.send_keys("{DOWN}", pause=0.1)
            keyboard.send_keys("{F5}", pause=0.1)
        except Exception as e:
            logging.error(f"Erro ao navegar no menu: {e}")
            return False

        logging.info("Verificando a janela de parâmetros do relatório...")
        report_window_timeout = 30
        report_start_time = time.time()
        report_window = None

        while time.time() - report_start_time < report_window_timeout:
            try:
                desktop = Desktop(backend="win32")
                windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible() and "AutoSystem" in w.window_text()]
                if windows:
                    report_window = windows[0]
                    logging.info(f"Janela de relatório encontrada: Título: {report_window.window_text()}, Handle: {report_window.handle}")
                    break
                logging.info("Procurando janela de relatório...")
                time.sleep(1)
            except Exception as e:
                logging.error(f"Erro ao procurar janela de relatório: {e}")
                time.sleep(1)

        if report_window:
            try:
                app.connect(handle=report_window.handle)
                report_dialog = app.window(handle=report_window.handle)
                report_dialog.wait('enabled ready visible', timeout=5)
                report_dialog.set_focus()
                time.sleep(0.2)
            except Exception as e:
                logging.error(f"Erro ao acessar a janela de relatório: {e}")
                return False
        else:
            logging.error("Janela de relatório não encontrada dentro do tempo limite de 30 segundos.")
            return False

        try:
            logging.info("Executando sequência final de teclas...")
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{ENTER}", pause=0.1)
            keyboard.send_keys("{F5}", pause=0.1)
            
            logging.info("Aguardando a janela 'Visualizar Impressão' abrir...")
            report_window_timeout = 30
            report_start_time = time.time()
            report_window_found = False

            while time.time() - report_start_time < report_window_timeout:
                try:
                    desktop = Desktop(backend="win32")
                    windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible() and "Visualizar Impressão" in w.window_text()]
                    if windows:
                        report_window = windows[0]
                        logging.info(f"Janela 'Visualizar Impressão' encontrada: {report_window.window_text()}")
                        report_window_found = True
                        break
                    logging.info("Aguardando janela 'Visualizar Impressão'...")
                    time.sleep(1)
                except Exception as e:
                    logging.error(f"Erro ao procurar janela 'Visualizar Impressão': {e}")
                    time.sleep(1)

            if not report_window_found:
                logging.error("Janela 'Visualizar Impressão' não encontrada dentro do tempo limite.")
                return False

            try:
                app.connect(handle=report_window.handle)
                report_dialog = app.window(handle=report_window.handle)
                report_dialog.wait('enabled ready visible', timeout=5)
                report_dialog.set_focus()
                time.sleep(1)
            except Exception as e:
                logging.error(f"Erro ao focalizar na janela 'Visualizar Impressão': {e}")
                return False

            logging.info("Executando sequência de teclas na janela 'Visualizar Impressão'...")
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{TAB}", pause=0.1)
            keyboard.send_keys("{ENTER}", pause=0.1)
            keyboard.send_keys("{F5}", pause=0.1)
            logging.info("Sequência de teclas concluída.")
            
            logging.info("Verificando janela 'Informação'...")
            info_window_timeout = 1
            info_start_time = time.time()
            
            while time.time() - info_start_time < info_window_timeout:
                try:
                    desktop = Desktop(backend="win32")
                    windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible() and "Informação" in w.window_text()]
                    if windows:
                        info_window = windows[0]
                        logging.info(f"Janela 'Informação' encontrada, fechando...")
                        keyboard.send_keys("{ENTER}", pause=0.1)
                        time.sleep(0.2)
                        # Aguarda 2 segundos e fecha o Bloco de Notas
                        logging.info("Aguardando 2 segundos para fechamento do Bloco de Notas após janela 'Informação'...")
                        time.sleep(2)
                        close_notepad()
                        break
                    time.sleep(0.2)
                except Exception as e:
                    logging.error(f"Erro ao verificar janela 'Informação': {e}")
                    time.sleep(0.2)

            logging.info("Verificando janela 'Erro'...")
            error_window_timeout = 1
            error_start_time = time.time()
            
            while time.time() - error_start_time < error_window_timeout:
                try:
                    desktop = Desktop(backend="win32")
                    windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible() and "Erro" in w.window_text()]
                    if windows:
                        error_window = windows[0]
                        logging.info(f"Janela 'Erro' encontrada, fechando...")
                        keyboard.send_keys("{ENTER}", pause=0.1)
                        time.sleep(0.2)
                        break
                    time.sleep(0.2)
                except Exception as e:
                    logging.error(f"Erro ao verificar janela 'Erro': {e}")
                    time.sleep(0.2)

            logging.info("Verificando janela 'Conferência de Caixa'...")
            caixa_window_timeout = 1
            caixa_start_time = time.time()
            
            while time.time() - caixa_start_time < caixa_window_timeout:
                try:
                    desktop = Desktop(backend="win32")
                    windows = [w for w in desktop.windows() if w.process_id() == process.pid and w.is_visible() and "Conferência de Caixa" in w.window_text()]
                    if windows:
                        caixa_window = windows[0]
                        logging.info(f"Janela 'Conferência de Caixa' encontrada, fechando...")
                        keyboard.send_keys("{ESC}", pause=0.1)
                        time.sleep(0.2)
                        break
                    time.sleep(0.2)
                except Exception as e:
                    logging.error(f"Erro ao verificar janela 'Conferência de Caixa': {e}")
                    time.sleep(0.2)

            notepad_closed = close_notepad()
            if not notepad_closed:
                logging.error("Falha ao fechar a janela do Bloco de Notas. Programa pode não finalizar corretamente (fechamento forçado). Aguarde mais tempo para o sistema processar.")
                time.sleep(0.5)
            else:
                logging.info("Bloco de Notas fechado normalmente.")
                time.sleep(0.2)
            # Focaliza na janela 'Visualizar Impressão' e envia a nova sequência
            try:
                app.connect(handle=report_window.handle)
                report_dialog = app.window(handle=report_window.handle)
                report_dialog.wait('enabled ready visible', timeout=3)
                report_dialog.set_focus()
                time.sleep(0.1)
                logging.info("Enviando sequência ESC ESC ESC ENTER na janela 'Visualizar Impressão'...")
                keyboard.send_keys("{ESC}", pause=0.1)
                keyboard.send_keys("{ESC}", pause=0.1)
                keyboard.send_keys("{ESC}", pause=0.1)
                keyboard.send_keys("{ENTER}", pause=0.1)
                time.sleep(0.3)
            except Exception as e:
                logging.error(f"Erro ao focalizar/enviar teclas para 'Visualizar Impressão': {e}")
            
            logging.info("Aguardando geração do relatório...")
            time.sleep(1)
            
            logging.info("Processo concluído. Verifique o relatório em C:\autosystem\relatorios.")
            # Após a geração do relatório e fechamento do Notepad, aguardar 3 segundos antes de copiar o arquivo de origem
            time.sleep(3)
            try:
                if os.path.exists(report_source_file):
                    if not os.path.exists(report_dest_folder):
                        os.makedirs(report_dest_folder)
                    dest_file = os.path.join(report_dest_folder, os.path.basename(report_source_file))
                    shutil.copy2(report_source_file, dest_file)
                    logging.info(f"Arquivo de relatório copiado de {report_source_file} para {dest_file}")
                else:
                    logging.warning(f"Arquivo de relatório não encontrado: {report_source_file} para copiar.")
            except Exception as e:
                logging.error(f"Erro ao copiar arquivo de relatório: {e}")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao executar sequência final de teclas: {e}")
            return False
    finally:
        if process and process.poll() is None: # Only terminate if the process is still running
            terminate_autosystem(process)
        if notepad_thread:
            stop_notepad_monitor_event.set()
            notepad_thread.join()

def pause_bot(icon, item):
    global bot_running
    bot_running = False
    icon.title = "Bot Relatórios - Pausado"
    logging.info("Bot pausado pelo usuário.")

def resume_bot(icon, item):
    global bot_running
    bot_running = True
    icon.title = "Bot Relatórios - Ativo"
    logging.info("Bot retomado pelo usuário.")

def exit_bot(icon, item):
    logging.info("Encerrando o bot...")
    global bot_running
    bot_running = False
    stop_notepad_monitor_event.set() # Ensure the notepad monitor thread stops
    icon.stop()
    sys.exit(0)

def run_system_tray():
    # Carrega o ícone
    try:
        image = Image.open("robot.ico")
    except Exception as e:
        logging.error(f"Erro ao carregar ícone: {e}. Usando ícone padrão.")
        image = Image.new('RGB', (64, 64), color = 'blue')  # Ícone fallback

    # Cria o menu do ícone
    menu = (
        pystray.MenuItem("Pausar", pause_bot),
        pystray.MenuItem("Retomar", resume_bot),
        pystray.MenuItem("Sair", exit_bot)
    )
    
    # Cria o ícone na bandeja do sistema
    icon = pystray.Icon("Bot Relatórios", image, "Bot Relatórios - Ativo", menu)
    icon.run()

def start_bot_logic(execution_frequency, autosystem_path):
    logging.info(f"Iniciando bot em {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    global bot_running
    bot_running = True # Initialize bot_running flag

    logging.info(f"Bot configured to run every {execution_frequency} minutes.")
    
    # Agenda a execução a cada X minutos, X vindo da configuração
    schedule.every(execution_frequency).minutes.do(exportar_relatorio, autosystem_path_arg=autosystem_path)
    
    # Inicia o ícone na bandeja do sistema em uma thread separada
    tray_thread = threading.Thread(target=run_system_tray, daemon=True)
    tray_thread.start()
    
    # Executa a primeira iteração imediatamente
    exportar_relatorio(autosystem_path_arg=autosystem_path)
    
    # Loop principal para o agendador
    while bot_running:
        schedule.run_pending()
        time.sleep(1)

def main():
    CONFIG_FILE = "bot_config.json"
    def load_config():
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                if "autosystem_path" not in config: 
                    config["autosystem_path"] = "C:\\autosystem\\main.exe"
                return config
        return {"execution_frequency_minutes": 2, "autosystem_path": "C:\\autosystem\\main.exe"} 
    
    config = load_config()
    execution_frequency = config["execution_frequency_minutes"]
    autosystem_path = config["autosystem_path"]

    start_bot_logic(execution_frequency, autosystem_path)

if __name__ == "__main__":
    main()

# Nova função para verificar e fechar Notepad após 8 segundos do início do main.exe
def close_txt_notepad_after_delay(report_path=r"C:\autosystem\relatorios", delay=8):
    def _check_and_close():
        logging.info(f"Aguardando {delay}s após início do main.exe para verificar arquivos .txt...")
        time.sleep(delay)
        if os.path.exists(report_path):
            txts = glob.glob(os.path.join(report_path, '*.txt'))
            if txts:
                logging.info(f"Arquivos .txt detectados: {txts}. Tentando fechar o Notepad imediatamente...")
                try:
                    import psutil
                    for proc in psutil.process_iter(['pid', 'name']):
                        if proc.name().lower() == "notepad.exe":
                            proc.terminate()
                            proc.wait(timeout=2)
                            logging.info(f"Processo notepad.exe (PID: {proc.pid}) encerrado.")
                except Exception as e:
                    logging.error(f"Erro ao tentar fechar o Notepad: {e}")
            else:
                logging.info("Nenhum arquivo .txt encontrado após o delay.")
        else:
            logging.info(f"Pasta de relatórios não encontrada: {report_path}")
    import threading
    threading.Thread(target=_check_and_close, daemon=True).start()