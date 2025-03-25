import sys
import os
import requests
import pythoncom

# 移除未使用的导入
from PyQt5.QtWidgets import (
    QApplication,
    QSystemTrayIcon,
    QMenu,
    QAction,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QTextEdit,
    QTabWidget,
    QDialog,
    QComboBox,  # 添加下拉选择控件
)
from PyQt5.QtGui import QIcon, QCursor, QFont
from PyQt5.QtCore import Qt, QPoint, QSize, QSettings, pyqtSignal, QTimer

# 导入新的文本提取器
from text_extractor import TextExtractor

# 初始化 COM
pythoncom.CoInitialize()

# 配置文件路径
APP_NAME = "ClickNow"
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), f".{APP_NAME.lower()}.ini")

# 默认提示词
DEFAULT_MAGNIFIER_PROMPT = "请通俗易懂地解释以下内容：\n{text}"
DEFAULT_DICTIONARY_PROMPT = "请将以下内容翻译成中文：\n{text}"

# AI API配置（示例使用，实际应用中需要替换为真实的API）
AI_API_URL = "http://192.168.20.63:11434"
AI_API_KEY = "ollama"
DEFAULT_AI_MODEL = "llama3"

# 在文件开头的常量定义区域添加默认模型配置
DEFAULT_MODELS = {
    "Ollama": "llama3",
    "DeepSeek": "deepseek-chat",
    "OpenAI": "gpt-3.5-turbo",
}


class FloatingButtons(QWidget):
    """悬浮按钮窗口"""

    magnifier_clicked = pyqtSignal(str)
    dictionary_clicked = pyqtSignal(str)

    def __init__(self, selected_text, parent=None):
        super().__init__(
            parent, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.selected_text = selected_text
        self.dragPosition = None

        # 获取DPI缩放因子
        screen = QApplication.primaryScreen()
        self.dpi_scale = screen.logicalDotsPerInch() / 96.0

        # 添加自动消失定时器
        self.hide_timer = QTimer(self)
        self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.close)
        self.hide_timer.start(5000)

        self.initUI()

    def enterEvent(self, event):  # 修改为正确的Qt事件名
        """鼠标进入时停止定时器"""
        self.hide_timer.stop()

    def leaveEvent(self, event):  # 修改为正确的Qt事件名
        """鼠标离开时重新开始定时器"""
        self.hide_timer.start(5000)

    def initUI(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # 放大镜按钮
        self.magnifier_btn = QPushButton()
        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        self.magnifier_btn.setIcon(QIcon(os.path.join(icon_dir, "magnifier.png")))
        self.magnifier_btn.setIconSize(QSize(48, 48))  # 增大图标尺寸
        self.magnifier_btn.setFixedSize(72, 72)  # 增大按钮尺寸
        self.magnifier_btn.setToolTip("解释所选文本")
        self.magnifier_btn.clicked.connect(self.on_magnifier_clicked)

        # 词典按钮
        self.dictionary_btn = QPushButton()
        self.dictionary_btn.setIcon(QIcon(os.path.join(icon_dir, "dictionary.png")))
        self.dictionary_btn.setIconSize(QSize(48, 48))  # 增大图标尺寸
        self.dictionary_btn.setFixedSize(72, 72)  # 增大按钮尺寸
        self.dictionary_btn.setToolTip("翻译所选文本")
        self.dictionary_btn.clicked.connect(self.on_dictionary_clicked)

        layout.addWidget(self.magnifier_btn)
        layout.addWidget(self.dictionary_btn)
        self.setLayout(layout)

        # 设置样式
        self.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50; /* 更显眼的绿色背景 */
                border-radius: 36px; /* 调整为按钮大小的一半，保持圆形 */
                border: 2px solid #2E7D32; /* 深绿色边框 */
                color: white; /* 白色文字 */
            }
            QPushButton:hover {
                background-color: #66BB6A; /* 悬停时颜色变亮 */
            }
            QPushButton:pressed {
                background-color: #388E3C; /* 按下时颜色变深 */
            }
        """)

        # 根据DPI缩放调整按钮大小
        button_size = int(27 * self.dpi_scale)  # 基础大小 * DPI缩放
        icon_size = int(18 * self.dpi_scale)  # 图标大小 * DPI缩放

        self.magnifier_btn.setIconSize(QSize(icon_size, icon_size))
        self.magnifier_btn.setFixedSize(button_size, button_size)

        self.dictionary_btn.setIconSize(QSize(icon_size, icon_size))
        self.dictionary_btn.setFixedSize(button_size, button_size)

    def on_magnifier_clicked(self):
        self.magnifier_clicked.emit(self.selected_text)
        self.close()

    def on_dictionary_clicked(self):
        self.dictionary_clicked.emit(self.selected_text)
        self.close()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and self.dragPosition is not None:
            self.move(event.globalPos() - self.dragPosition)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = None
            event.accept()


class ResultWindow(QWidget):
    """结果显示窗口"""

    def __init__(self, title, content, parent=None):
        super().__init__(parent, Qt.WindowStaysOnTopHint)
        self.setWindowTitle(title)

        # 获取DPI缩放因子
        screen = QApplication.primaryScreen()
        self.dpi_scale = screen.logicalDotsPerInch() / 96.0

        # 设置基础大小
        base_width = 500
        base_height = 400

        # 根据DPI缩放调整大小
        scaled_width = int(base_width * self.dpi_scale)
        scaled_height = int(base_height * self.dpi_scale)

        self.setMinimumSize(scaled_width, scaled_height)

        # 从设置中读取窗口大小
        settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)
        size = settings.value(
            f"result_window_size_{title}", QSize(scaled_width, scaled_height)
        )
        self.resize(size)

        # 初始化UI
        layout = QVBoxLayout()

        # 结果文本框
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Microsoft YaHei", 12))  # 增大字体大小

        # 直接设置纯文本内容
        self.result_text.setPlainText(content)

        # 确保内容可见
        self.result_text.setVisible(True)
        self.result_text.ensureCursorVisible()

        layout.addWidget(self.result_text)
        self.setLayout(layout)

        # 设置样式
        self.setStyleSheet("""
            QTextEdit {
                border: 1px solid #cccccc;
                background-color: #ffffff;
                padding: 10px;
                line-height: 1.6;
            }
        """)

    def closeEvent(self, event):
        """窗口关闭时保存大小"""
        settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)
        settings.setValue(f"result_window_size_{self.windowTitle()}", self.size())
        event.accept()

    def initUI(self, content):
        layout = QVBoxLayout()

        # 结果文本框
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Microsoft YaHei", 10))

        # 直接设置文本内容

        # 先设置纯文本
        self.result_text.setPlainText(content)

        # 确保内容可见
        self.result_text.setVisible(True)
        self.result_text.ensureCursorVisible()

        layout.addWidget(self.result_text)
        self.setLayout(layout)

        # 设置样式
        self.setStyleSheet("""
            QTextEdit {
                border: 1px solid #cccccc;
                background-color: #ffffff;
            }
        """)


class SettingsDialog(QDialog):
    """设置对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setMinimumSize(500, 400)
        self.settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)

        # 从设置中读取窗口大小和位置
        size = self.settings.value("settings_dialog_size", QSize(500, 400))
        pos = self.settings.value("settings_dialog_pos", QPoint(100, 100))

        self.resize(size)
        self.move(pos)

        self.initUI()
        self.loadSettings()

    def initUI(self):
        layout = QVBoxLayout()
        self.tabs = QTabWidget()

        # 放大镜和词典页签（保持不变）...
        self.magnifier_tab = QWidget()
        magnifier_layout = QVBoxLayout()
        magnifier_label = QLabel("放大镜按钮提示词：")
        self.magnifier_prompt = QTextEdit()
        magnifier_layout.addWidget(magnifier_label)
        magnifier_layout.addWidget(self.magnifier_prompt)
        self.magnifier_tab.setLayout(magnifier_layout)

        self.dictionary_tab = QWidget()
        dictionary_layout = QVBoxLayout()
        dictionary_label = QLabel("词典按钮提示词：")
        self.dictionary_prompt = QTextEdit()
        dictionary_layout.addWidget(dictionary_label)
        dictionary_layout.addWidget(self.dictionary_prompt)
        self.dictionary_tab.setLayout(dictionary_layout)

        # AI模型标签页（修改部分）
        self.ai_model_tab = QWidget()
        ai_model_layout = QVBoxLayout()

        # 模型提供方设置
        provider_layout = QHBoxLayout()
        provider_label = QLabel("模型提供方：")
        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Ollama", "DeepSeek", "OpenAI"])
        provider_layout.addWidget(provider_label)
        provider_layout.addWidget(self.provider_combo)

        # API KEY设置（新增）
        self.api_key_container = QWidget()
        api_key_layout = QHBoxLayout(self.api_key_container)
        api_key_layout.setContentsMargins(0, 0, 0, 0)
        api_key_label = QLabel("API KEY：")
        self.api_key_input = QTextEdit()
        self.api_key_input.setMaximumHeight(60)
        api_key_layout.addWidget(api_key_label)
        api_key_layout.addWidget(self.api_key_input)

        # API URL设置
        self.api_url_container = QWidget()
        api_url_layout = QHBoxLayout(self.api_url_container)
        api_url_layout.setContentsMargins(0, 0, 0, 0)
        api_url_label = QLabel("API URL：")
        self.api_url_input = QTextEdit()
        self.api_url_input.setMaximumHeight(60)
        api_url_layout.addWidget(api_url_label)
        api_url_layout.addWidget(self.api_url_input)

        # 模型名称设置
        model_name_layout = QHBoxLayout()
        model_name_label = QLabel("模型名称：")
        self.model_name_input = QTextEdit()
        self.model_name_input.setMaximumHeight(60)
        model_name_layout.addWidget(model_name_label)
        model_name_layout.addWidget(self.model_name_input)

        ai_model_layout.addLayout(provider_layout)
        ai_model_layout.addWidget(self.api_key_container)
        ai_model_layout.addWidget(self.api_url_container)
        ai_model_layout.addLayout(model_name_layout)
        ai_model_layout.addStretch(1)
        self.ai_model_tab.setLayout(ai_model_layout)

        self.tabs.addTab(self.magnifier_tab, "放大镜")
        self.tabs.addTab(self.dictionary_tab, "词典")
        self.tabs.addTab(self.ai_model_tab, "AI模型")

        button_layout = QHBoxLayout()
        self.save_btn = QPushButton("保存")
        self.cancel_btn = QPushButton("取消")
        self.save_btn.clicked.connect(self.saveSettings)
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addStretch(1)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)

        layout.addWidget(self.tabs)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        # 注册 provider 变化事件，根据选择调整显示内容
        self.provider_combo.currentTextChanged.connect(self.update_provider_fields)

    def update_provider_fields(self, provider):
        """根据不同的提供方更新界面显示"""
        self.api_key_container.setVisible(provider in ["DeepSeek", "OpenAI"])
        self.api_url_container.setVisible(provider == "Ollama")

        # 加载当前提供方对应的模型名
        model_name = self.settings.value(
            f"ai_model_{provider}", DEFAULT_MODELS.get(provider, "")
        )
        self.model_name_input.setPlainText(model_name)

    def loadSettings(self):
        magnifier_prompt = self.settings.value(
            "magnifier_prompt", DEFAULT_MAGNIFIER_PROMPT
        )
        dictionary_prompt = self.settings.value(
            "dictionary_prompt", DEFAULT_DICTIONARY_PROMPT
        )
        api_url = self.settings.value("ai_api_url", AI_API_URL)
        provider = self.settings.value("ai_provider", "Ollama")
        # 根据当前提供方加载对应的API Key
        api_key = self.settings.value(f"ai_api_key_{provider}", "")
        # 按提供方获取对应的模型名
        model_name = self.settings.value(
            f"ai_model_{provider}", DEFAULT_MODELS.get(provider)
        )

        self.magnifier_prompt.setText(magnifier_prompt)
        self.dictionary_prompt.setText(dictionary_prompt)
        self.api_url_input.setText(api_url)
        self.api_key_input.setText(api_key)
        self.model_name_input.setText(model_name)

        index = self.provider_combo.findText(provider)
        if index != -1:
            self.provider_combo.setCurrentIndex(index)
        self.update_provider_fields(provider)

    def saveSettings(self):
        current_provider = self.provider_combo.currentText()

        self.settings.setValue("magnifier_prompt", self.magnifier_prompt.toPlainText())
        self.settings.setValue(
            "dictionary_prompt", self.dictionary_prompt.toPlainText()
        )
        self.settings.setValue("ai_api_url", self.api_url_input.toPlainText())
        # 按提供方分别存储API Key
        self.settings.setValue(
            f"ai_api_key_{current_provider}", self.api_key_input.toPlainText()
        )
        # 按提供方分别存储模型名
        self.settings.setValue(
            f"ai_model_{current_provider}", self.model_name_input.toPlainText()
        )
        self.settings.setValue("ai_provider", current_provider)
        self.settings.setValue("settings_dialog_size", self.size())
        self.settings.setValue("settings_dialog_pos", self.pos())
        self.settings.sync()

        print("[Settings] 保存设置:")
        print(f"  - AI提供方: {current_provider}")
        print(f"  - 模型名称: {self.model_name_input.toPlainText()}")
        print(f"  - API URL: {self.api_url_input.toPlainText()}")

        self.accept()


# 使用新的TextExtractor类替代原来的ClipboardMonitor类


class ClickNowApp(QApplication):
    """主应用程序类"""

    def __init__(self, argv):
        super().__init__(argv)
        self.setQuitOnLastWindowClosed(False)
        self.settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)

        # 初始化系统托盘
        self.init_tray()

        # 初始化文本提取器
        self.clipboard_monitor = TextExtractor()
        self.clipboard_monitor.text_selected.connect(self.on_text_selected)

        # 悬浮按钮和结果窗口
        self.floating_buttons = None
        self.result_window = None
        self.settings_dialog = None  # 添加设置对话框变量

    def init_tray(self):
        # 创建系统托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        icon_path = os.path.join(icon_dir, "app_icon.png")

        # 预先加载图标，避免每次显示菜单时加载
        self.app_icon = QIcon(icon_path)
        self.tray_icon.setIcon(self.app_icon)
        self.tray_icon.setToolTip(APP_NAME)

        # 创建托盘菜单并保存为实例变量，避免每次右键时重新创建
        self.tray_menu = QMenu()

        settings_action = QAction("设置", self)
        settings_action.triggered.connect(self.show_settings)

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.quit)

        self.tray_menu.addAction(settings_action)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.show()

    def show_settings(self):
        """显示设置对话框"""
        if not hasattr(self, "settings_dialog") or not self.settings_dialog:
            self.settings_dialog = SettingsDialog()
        self.settings_dialog.show()
        self.settings_dialog.raise_()  # 确保窗口在最前面
        self.settings_dialog.activateWindow()  # 激活窗口

    def on_text_selected(self, text, pos):
        # 如果已有悬浮按钮，先关闭
        if self.floating_buttons and self.floating_buttons.isVisible():
            self.floating_buttons.close()

        # 确保文本不为空才显示悬浮按钮
        if not text or not text.strip():
            return

        print(f"[Selection] 检测到选中文本: {text[:100]}...")

        # 创建新的悬浮按钮，始终使用最新选中的文本
        self.floating_buttons = FloatingButtons(text)
        self.floating_buttons.magnifier_clicked.connect(self.on_magnifier_clicked)
        self.floating_buttons.dictionary_clicked.connect(self.on_dictionary_clicked)

        # 获取屏幕DPI缩放因子
        screen = self.primaryScreen()

        # 计算按钮位置，使其出现在选中文本的右上方
        button_width = self.floating_buttons.width()
        button_height = self.floating_buttons.height()

        # 计算按钮位置，确保不会超出屏幕边界
        screen_geometry = screen.geometry()
        adjusted_x = min(pos.x() + 20, screen_geometry.width() - button_width - 10)
        adjusted_y = max(pos.y() - button_height - 20, 10)  # 确保不会超出屏幕顶部

        # 移动按钮到计算出的位置
        self.floating_buttons.move(adjusted_x, adjusted_y)
        self.floating_buttons.show()

    def on_magnifier_clicked(self, text):
        # 获取放大镜提示词
        prompt = self.settings.value("magnifier_prompt", DEFAULT_MAGNIFIER_PROMPT)
        # 打印调试信息
        prompt = prompt.format(text=text)

        # 重置剪贴板监视器的last_selected_text，以便下次选中相同文本时也能触发
        if hasattr(self, "clipboard_monitor") and self.clipboard_monitor:
            self.clipboard_monitor.last_selected_text = ""

        # 调用AI API获取结果（示例实现）
        result = self.call_ai_api(prompt)

        # 显示结果窗口
        if self.result_window:
            self.result_window.close()

        # 保存悬浮按钮位置，用于结果窗口显示
        button_pos = None
        if self.floating_buttons:
            button_pos = self.floating_buttons.pos()

        self.result_window = ResultWindow("解释结果", result)

        # 如果有悬浮按钮位置，则在该位置显示结果窗口，否则在鼠标位置显示
        if button_pos:
            self.result_window.move(button_pos)
        else:
            self.result_window.move(QCursor.pos() + QPoint(20, 20))

        self.result_window.show()

        # 根据内容自动调整窗口大小
        self.result_window.adjustSize()

    def on_dictionary_clicked(self, text):
        # 获取词典提示词
        prompt = self.settings.value("dictionary_prompt", DEFAULT_DICTIONARY_PROMPT)
        # 打印调试信息
        prompt = prompt.format(text=text)

        # 重置剪贴板监视器的last_selected_text，以便下次选中相同文本时也能触发
        if hasattr(self, "clipboard_monitor") and self.clipboard_monitor:
            self.clipboard_monitor.last_selected_text = ""

        # 调用AI API获取结果（示例实现）
        result = self.call_ai_api(prompt)

        # 显示结果窗口
        if self.result_window:
            self.result_window.close()

        # 保存悬浮按钮位置，用于结果窗口显示
        button_pos = None
        if self.floating_buttons:
            button_pos = self.floating_buttons.pos()

        self.result_window = ResultWindow("翻译结果", result)

        # 如果有悬浮按钮位置，则在该位置显示结果窗口，否则在鼠标位置显示
        if button_pos:
            self.result_window.move(button_pos)
        else:
            self.result_window.move(QCursor.pos() + QPoint(20, 20))

        self.result_window.show()

        # 根据内容自动调整窗口大小
        self.result_window.adjustSize()

    def call_ai_api(self, prompt):
        try:
            api_url = self.settings.value("ai_api_url", AI_API_URL)
            provider = self.settings.value("ai_provider", "Ollama")
            # 获取当前提供方对应的API Key
            api_key = self.settings.value(f"ai_api_key_{provider}", "")
            # 获取当前提供方对应的模型名
            model_name = self.settings.value(
                f"ai_model_{provider}", DEFAULT_MODELS.get(provider, DEFAULT_AI_MODEL)
            )

            print("[API] 调用AI服务:")
            print(f"  - 提供方: {provider}")
            print(f"  - 模型: {model_name}")
            print(f"  - 提示词: {prompt[:100]}...")

            if provider == "Ollama":
                print(f"  - 接口URL: {api_url}/api/generate")
                api_endpoint = f"{api_url}/api/generate"
                data = {"model": model_name, "prompt": prompt, "stream": False}
                response = requests.post(api_endpoint, json=data)
                result = response.json().get("response", "")
            elif provider == "DeepSeek":
                print("  - 接口URL: https://api.deepseek.com/v1/chat/completions")
                api_endpoint = "https://api.deepseek.com/v1/chat/completions"
                data = {
                    "model": model_name,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.7,
                    "max_tokens": 2000,
                }
                headers = {
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                }
                response = requests.post(api_endpoint, json=data, headers=headers)
                result = (
                    response.json()
                    .get("choices", [{}])[0]
                    .get("message", {})
                    .get("content", "")
                )
            elif provider == "OpenAI":
                print("  - 接口URL: https://api.openai.com/v1/chat/completions")
                api_endpoint = "https://api.openai.com/v1/chat/completions"
                data = {
                    "model": model_name,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.7,
                    "max_tokens": 2000,
                }
                headers = {
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                }
                response = requests.post(api_endpoint, json=data, headers=headers)
                result = (
                    response.json()
                    .get("choices", [{}])[0]
                    .get("message", {})
                    .get("content", "")
                )

            if response.status_code == 200:
                # 清理HTML标签
                if result:
                    if "<think>" in result:
                        result = result.split("</think>")[-1].strip()
                    result = result.replace("<p>", "").replace("</p>", "\n")
                    result = result.replace("<br>", "\n").replace("<br/>", "\n")
                    result = result.replace("<div>", "").replace("</div>", "\n")
                    result = result.replace("<code>", "").replace("</code>", "")
                    result = result.replace("<pre>", "").replace("</pre>", "\n")

                print("[API] 请求成功")
                return result
            else:
                error_msg = (
                    f"请求失败: HTTP状态码 {response.status_code}\n{response.text}"
                )
                print(f"[API] {error_msg}")
                return error_msg
        except Exception as e:
            error_msg = f"请求失败: {str(e)}"
            print(f"[API] {error_msg}")
            return error_msg


def main():
    app = ClickNowApp(sys.argv)
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
