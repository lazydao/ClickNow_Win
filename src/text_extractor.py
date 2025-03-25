import time
import win32con
import win32api
import uiautomation as auto
from PyQt5.QtCore import QObject, pyqtSignal, QPoint, QTimer
from PyQt5.QtGui import QCursor


class TextExtractor(QObject):
    """使用UI自动化获取选中文本，不使用剪贴板或模拟按键"""

    text_selected = pyqtSignal(str, QPoint)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_selection)
        self.timer.start(300)  # 增加检查间隔，减少重复获取
        self.last_selected_text = ""
        self.last_cursor_pos = QCursor.pos()
        self.is_checking = False
        self.is_mouse_down = False
        self.mouse_down_position = None
        self.last_emit_time = time.time()  # 添加最后一次发送信号的时间记录

    def _get_text_from_element(self, element):
        """辅助函数：尝试从一个元素中获取选中文本"""
        try:
            if hasattr(element, "GetTextPattern"):
                print("元素支持TextPattern")
                text_pattern = element.GetTextPattern()
                if text_pattern:
                    selection = text_pattern.GetSelection()
                    if selection and len(selection) > 0:
                        text = selection[0].GetText(-1)
                        if text:
                            print(f"从元素获取到文本: {text}")
                            return text
        except Exception as e:
            print(f"从元素获取文本失败: {e}")
        return ""

    def get_selected_text_from_automation(self):
        """使用UI自动化获取选中文本，增加对子控件的遍历尝试"""
        try:
            print("\n=== 开始获取选中文本 ===")
            print(
                f"当前鼠标位置: ({self.last_cursor_pos.x()}, {self.last_cursor_pos.y()})"
            )

            # 获取当前鼠标位置下的控件
            element = auto.ControlFromPoint(
                self.last_cursor_pos.x(), self.last_cursor_pos.y()
            )
            if not element:
                print("未找到鼠标位置下的控件")
                return ""

            print(f"找到控件: {getattr(element, 'Name', '未知控件')}")
            print(f"控件类型: {getattr(element, 'ClassName', '未知类型')}")

            # 尝试直接从当前控件获取选中文本
            text = self._get_text_from_element(element)
            if text:
                print(f"成功从控件获取到文本: {text}")
                return text

            # 遍历子控件进行尝试
            children = element.GetChildren()
            if children:
                print("当前控件未获取到文本，尝试遍历子控件")
                for child in children:
                    text = self._get_text_from_element(child)
                    if text:
                        return text

            print("未能获取到选中文本")
            return ""
        except Exception as e:
            print(f"自动化获取文本失败: {str(e)}")
            return ""

    def check_selection(self):
        """检查是否有文本被选中"""
        if self.is_checking:
            return

        try:
            self.is_checking = True
            current_cursor_pos = QCursor.pos()
            current_time = time.time()

            # 获取鼠标按键状态
            mouse_down = win32api.GetKeyState(win32con.VK_LBUTTON) < 0

            # 记录鼠标按下的位置
            if mouse_down and not self.is_mouse_down:
                self.mouse_down_position = current_cursor_pos
                print("\n=== 鼠标按下 ===")
                print(f"按下位置: ({current_cursor_pos.x()}, {current_cursor_pos.y()})")

            # 检测鼠标释放的时刻
            if self.is_mouse_down and not mouse_down and self.mouse_down_position:
                print("\n=== 鼠标释放 ===")
                print(f"释放位置: ({current_cursor_pos.x()}, {current_cursor_pos.y()})")

                # 计算鼠标移动距离
                move_distance = (
                    (current_cursor_pos.x() - self.mouse_down_position.x()) ** 2
                    + (current_cursor_pos.y() - self.mouse_down_position.y()) ** 2
                ) ** 0.5

                print(f"鼠标移动距离: {move_distance}")

                # 只有当鼠标移动了一定距离才认为可能有文本选择
                if move_distance > 5:
                    print("移动距离超过阈值，可能进行了文本选择")
                    # 等待一小段时间确保选择完成
                    print("等待0.5秒确保选择完成...")
                    time.sleep(0.5)

                    # 获取选中文本
                    print("开始获取选中文本...")
                    selected_text = self.get_selected_text_from_automation()
                    print(f"获取到的文本: {selected_text}")

                    # 确保距离上次发送信号至少有1秒
                    time_since_last_emit = current_time - self.last_emit_time
                    print(f"距离上次发送时间: {time_since_last_emit:.2f}秒")

                    if (
                        selected_text
                        and selected_text.strip()
                        and selected_text != self.last_selected_text
                        and time_since_last_emit > 1
                    ):
                        print("文本有效且未重复，发送信号")
                        self.last_selected_text = selected_text
                        self.text_selected.emit(selected_text, current_cursor_pos)
                        self.last_emit_time = current_time
                        print(f"发送选中文本: {selected_text}")
                    else:
                        if not selected_text:
                            print("未获取到文本")
                        elif not selected_text.strip():
                            print("获取到的文本为空")
                        elif selected_text == self.last_selected_text:
                            print("文本与上次相同")
                        elif time_since_last_emit <= 1:
                            print("发送间隔太短")
                else:
                    print("移动距离不足，可能是点击而非选择")

                # 重置鼠标按下位置
                self.mouse_down_position = None
                print("重置鼠标按下位置")

            # 更新鼠标状态
            self.is_mouse_down = mouse_down

            # 更新鼠标位置
            self.last_cursor_pos = current_cursor_pos

        except Exception as e:
            print(f"检查选中文本时发生错误: {str(e)}")
        finally:
            self.is_checking = False
