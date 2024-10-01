import sys
import subprocess
import requests
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QTextEdit, QVBoxLayout, QWidget, QLabel, QComboBox, QSystemTrayIcon, QMenu, QAction, QCheckBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QCoreApplication
import win32print
import winreg
import os


class ApiProcess(QThread):
    """Thread para gerenciar o processo da API"""

    def __init__(self):
        super().__init__()
        self.process = None
        self.running = False

    def run(self):
        # Configura as informações de inicialização para ocultar a janela
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        # Executa o script da API sem mostrar a janela do terminal
        self.process = subprocess.Popen(
            ['python', 'app.py'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo)
        self.running = True

        # Aguarda o processo terminar
        stdout, stderr = self.process.communicate()
        if stdout:
            print(stdout.decode())
        if stderr:
            print(stderr.decode())

        self.running = False

    def stop(self):
        if self.process and self.running:
            self.process.terminate()
            self.process.wait()  # Espera o processo ser encerrado
            self.process = None
            self.running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Printer Control")
        self.setGeometry(100, 100, 500, 500)
        self.setWindowIcon(QIcon('icon.png'))

        self.status_label = QLabel("API Status: Unknown")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.print_button = QPushButton("Print Text")
        self.print_button.clicked.connect(self.print_text)

        self.api_status_button = QPushButton("Check API Status")
        self.api_status_button.clicked.connect(self.check_api_status)

        self.start_api_button = QPushButton("Start API")
        self.start_api_button.clicked.connect(self.start_api)

        self.stop_api_button = QPushButton("Stop API")
        self.stop_api_button.clicked.connect(self.stop_api)

        self.exit_button = QPushButton("Exit Application")
        self.exit_button.clicked.connect(self.close_application)

        self.printer_combo = QComboBox()
        self.font_size_combo = QComboBox()
        self.margin_left_combo = QComboBox()
        self.margin_top_combo = QComboBox()
        self.margin_right_combo = QComboBox()
        self.margin_bottom_combo = QComboBox()

        # Adiciona tamanhos de fontes e margens
        self.font_size_combo.addItems(['21', '22', '24', '26', '28', '30'])
        self.margin_left_combo.addItems(['10', '20', '30', '40', '50'])
        self.margin_top_combo.addItems(['10', '20', '30', '40', '50'])
        self.margin_right_combo.addItems(['10', '20', '30', '40', '50'])
        self.margin_bottom_combo.addItems(['10', '20', '30', '40', '50'])

        self.load_printers()

        self.start_with_windows_checkbox = QCheckBox("Start with Windows")
        self.start_with_windows_checkbox.setChecked(
            self.check_startup_status())
        self.start_with_windows_checkbox.stateChanged.connect(
            self.set_startup_status)

        layout = QVBoxLayout()
        layout.addWidget(self.status_label)
        layout.addWidget(self.log_text)
        layout.addWidget(QLabel("Select Printer:"))
        layout.addWidget(self.printer_combo)
        layout.addWidget(QLabel("Select Font Size:"))
        layout.addWidget(self.font_size_combo)
        layout.addWidget(QLabel("Left Margin:"))
        layout.addWidget(self.margin_left_combo)
        layout.addWidget(QLabel("Top Margin:"))
        layout.addWidget(self.margin_top_combo)
        layout.addWidget(QLabel("Right Margin:"))
        layout.addWidget(self.margin_right_combo)
        layout.addWidget(QLabel("Bottom Margin:"))
        layout.addWidget(self.margin_bottom_combo)
        layout.addWidget(self.print_button)
        layout.addWidget(self.api_status_button)
        layout.addWidget(self.start_api_button)
        layout.addWidget(self.stop_api_button)
        layout.addWidget(self.exit_button)
        layout.addWidget(self.start_with_windows_checkbox)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.api_process = ApiProcess()

        # Configura o ícone da bandeja do sistema
        self.tray_icon = QSystemTrayIcon(QIcon('icon.png'), self)
        self.tray_icon.setToolTip('Servidor de Impressão')

        # Cria o menu da bandeja do sistema
        tray_menu = QMenu()
        restore_action = QAction("Restore")
        restore_action.triggered.connect(self.restore_window)
        exit_action = QAction("Exit")
        exit_action.triggered.connect(self.close_application)
        tray_menu.addAction(restore_action)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # Adiciona evento para mostrar a janela ao clicar no ícone da bandeja
        self.tray_icon.activated.connect(self.tray_icon_activated)

    def load_printers(self):
        # Lista as impressoras disponíveis
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
        printer_names = [printer[2] for printer in printers]
        self.printer_combo.addItems(printer_names)

    def print_text(self):
        text = '{ "data": { "font_size": 25, "printer_name": "POS58 10.0.0.6", "type": "text", "text": "Teste" } }'
        printer_name = self.printer_combo.currentText()
        font_size = int(self.font_size_combo.currentText())
        margins = {
            'left': max(int(self.margin_left_combo.currentText()), 10),
            'top': max(int(self.margin_top_combo.currentText()), 10),
            'right': max(int(self.margin_right_combo.currentText()), 10),
            'bottom': max(int(self.margin_bottom_combo.currentText()), 10)
        }
        try:
            response = requests.post('http://127.0.0.1:5000/print', json={
                'text': text, 'printer_name': printer_name, 'font_size': font_size, 'margins': margins})
            self.log_text.append(f"Print Request Status: {response.status_code}")
            self.log_text.append(f"Response: {response.json()}")
        except requests.RequestException as e:
            self.log_text.append(f"Error: {e}")

    def check_api_status(self):
        try:
            response = requests.get('http://127.0.0.1:5000/status')
            if response.status_code == 200:
                self.status_label.setText("API Status: Running")
            else:
                self.status_label.setText("API Status: Not Running")
        except requests.RequestException as e:
            self.status_label.setText(f"API Status: Error - {e}")

    def start_api(self):
        self.api_process.start()
        self.status_label.setText("API Status: Starting")

    def stop_api(self):
        self.api_process.stop()
        self.status_label.setText("API Status: Stopped")

    def restore_window(self):
        self.showNormal()
        self.activateWindow()

    def close_application(self):
        self.api_process.stop()
        QCoreApplication.instance().quit()

    def closeEvent(self, event):
        """Minimiza a janela ao invés de fechá-la"""
        event.ignore()
        self.hide()

    def set_startup_status(self, state):
        if state == Qt.Checked:
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def add_to_startup(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, 'PrinterControlApp', 0, winreg.REG_SZ,
                              sys.executable + ' ' + os.path.abspath(__file__))
            winreg.CloseKey(key)
        except Exception as e:
            self.log_text.append(f"Error setting startup: {e}")

    def remove_from_startup(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, 'PrinterControlApp')
            winreg.CloseKey(key)
        except Exception as e:
            self.log_text.append(f"Error removing startup: {e}")

    def check_startup_status(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, 'PrinterControlApp')
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            self.log_text.append(f"Error checking startup: {e}")
            return False

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.restore_window()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
