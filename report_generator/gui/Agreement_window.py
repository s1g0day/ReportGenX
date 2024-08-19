import os
import sys
import yaml
import warnings
from datetime import datetime
from gui.ui_main_windows import MainWindow
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QMessageBox, QTextEdit, QCheckBox, QFrame, QSizePolicy
warnings.filterwarnings("ignore", category=DeprecationWarning)

class DisclaimerWindow(QMainWindow):
    def __init__(self, push_config):
        super().__init__()
        self.push_config = push_config

        self.setWindowTitle("用户须知")
        self.setFixedSize(500, 300)  # 固定窗口大小

        self.create_widgets()

    def create_widgets(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        disclaimer_label = QLabel("渗透测试报告生成")
        disclaimer_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        layout.addWidget(disclaimer_label)

        self.content_text = QTextEdit()
        self.content_text.setFixedSize(460, 150)
        
        self.content_text.setReadOnly(True)
        self.content_text.insertPlainText(
            "免责声明:\n\n"
            "本工具仅用于安全测试为目的，使用者需遵守当地的法律法规。\n"
            "使用本工具导致的一切后果由使用者承担。谨慎使用，请勿用于非法用途。\n"
            "如果你同意此免责声明，请点击\"同意\"继续使用，否则点击\"拒绝\"退出。"
        )
        layout.addWidget(self.content_text)

        self.agree_checkbox = QCheckBox("我同意以上免责条款")
        self.agree_checkbox.setChecked(True)
        layout.addWidget(self.agree_checkbox)

        button_frame = QFrame()
        layout.addWidget(button_frame)

        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 20, 0, 0)

        agree_button = QPushButton("同 意")
        agree_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        agree_button.setStyleSheet("background-color: #008000; color: white; font-size: 18px;")
        agree_button.clicked.connect(self.agree_action)
        button_layout.addWidget(agree_button)

        cancel_button = QPushButton("拒 绝")
        cancel_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        cancel_button.setStyleSheet("background-color: #FF0000; color: white; font-size: 18px;")
        # cancel_button.clicked.connect(self.close)
        cancel_button.clicked.connect(QApplication.quit)  # 退出应用程序
        button_layout.addWidget(cancel_button)

    def agree_action(self):
        if not self.agree_checkbox.isChecked():
            QMessageBox.warning(self, "请注意!", "请先同意以上免责条款")
        else:
            self.close()
            main_window = MainWindow(self.push_config)
            main_window.show()

            # 记录已经同意免责声明的日期到文件中
            today = datetime.now().strftime('%Y-%m-%d')
            with open(self.push_config["agreed_dates_log_file"], "a+") as file:
                file.seek(0)
                agreed_dates = file.read().splitlines()

                if today not in agreed_dates:
                    file.write(today + "\n")

def is_first_run(agreed_dates_log):
    
    today = datetime.now().strftime('%Y-%m-%d')
    if not os.path.isfile(agreed_dates_log):
        # 获取项目根目录路径
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # 获取日志目录
        directory_path = os.path.dirname(agreed_dates_log)
        # 创建日志目录
        log_dir = os.path.join(root_dir, directory_path)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        return True
    else:
        with open(agreed_dates_log, "r") as file:
            agreed_dates = file.read().splitlines()
            return today not in agreed_dates

def ReportGenX_main():
    push_config = yaml.safe_load(open("conf/config.yaml", "r", encoding="utf-8").read())
    app = QApplication([])
    agreed_dates_log = push_config["agreed_dates_log_file"]
    if is_first_run(agreed_dates_log):
        disclaimer_window = DisclaimerWindow(push_config)
        disclaimer_window.show()
    else:
        main_window = MainWindow(push_config)
        main_window.show()

    sys.exit(app.exec())