"""企业微信窗口、搜索和资料提取。模板与偏移沿用原脚本。"""

from pathlib import Path
import time

from airtest.core import api
from airtest.core.cv import Template
import mss
from PIL import Image
import pyperclip
import win32gui
import win32process


ASSETS = Path(__file__).resolve().parent / "assets"
NICKNAME_OFFSETS = ((30, 90), (30, 95), (35, 85), (25, 100), (40, 90), (20, 90))


def windows(class_name, pid=None):
    matches = []

    def visit(handle, _):
        if win32gui.IsWindowVisible(handle) and win32gui.GetClassName(handle) == class_name:
            process = win32process.GetWindowThreadProcessId(handle)[1]
            if pid is None or process == pid:
                matches.append((handle, process))

    win32gui.EnumWindows(visit, None)
    return matches


class WeCom:
    def __init__(self, pid=None):
        self.pid = pid
        self.handle = None
        self.device = None

    def connect_main(self):
        candidates = windows("WeWorkWindow", self.pid)
        if not candidates:
            raise RuntimeError("未找到企业微信主窗口，请打开并登录后重试")
        if len(candidates) != 1:
            raise RuntimeError("存在多个企业微信窗口，请用 --pid 指定进程 ID")
        self.handle, self.pid = candidates[0]
        self.connect(self.handle)

    def connect(self, handle):
        if not win32gui.IsWindow(handle):
            raise RuntimeError("企业微信窗口已失效，请重新登录后重试")
        self.device = api.connect_device(f"Windows:///{handle}?foreground=False")
        self.device.set_foreground()
        self.device.move((0, 0))

    def find(self, *names):
        for name in names:
            point = api.exists(Template(str(ASSETS / name)))
            if point:
                return point
        return None

    def click(self, *names):
        point = self.find(*names)
        if not point:
            raise RuntimeError(f"未匹配到界面模板：{', '.join(names)}；请检查界面版本和显示缩放")
        api.touch(point)
        time.sleep(0.5)
        return point

    def search_window(self):
        if self.handle is None:
            self.connect_main()
        elif not win32gui.IsWindow(self.handle):
            raise RuntimeError("企业微信主窗口已关闭，请重新运行")
        matches = windows("SearchExternalsWnd", self.pid)
        if not matches:
            self.connect(self.handle)
            contacts = self.find("main/navbar_contact_button.png", "main/navbar_contact_button_hover.png",
                                 "main/navbar_contact_button_focus.png")
            if contacts:
                api.touch(contacts)
            else:
                self.click("main/more_button.png")
                self.click("main/contacts_menu.png")
            time.sleep(1)
            self.click("contact_new_friend_logo_btn.png")
            self.click("contact_new_friend_logo_btn2.png")
            for _ in range(10):
                time.sleep(0.5)
                matches = windows("SearchExternalsWnd", self.pid)
                if matches:
                    break
        if len(matches) != 1:
            raise RuntimeError("无法确定添加客户窗口，请关闭重复窗口后重试")
        self.connect(matches[0][0])
        time.sleep(1)

    def nickname(self, x, y, phone):
        for dx, dy in NICKNAME_OFFSETS:
            pyperclip.copy("")
            api.touch((x + dx, y + dy))
            api.double_click((x + dx, y + dy))
            api.keyevent("^c")
            time.sleep(0.4)
            name = pyperclip.paste().strip()
            if (name and len(name) < 50 and name != phone
                    and not any(word in name for word in
                                ("添加", "联系人", "好友", "确定", "取消", "搜索", "设置", "电话", "标签", "微信", "WECOM"))):
                return name
        raise RuntimeError("未能复制昵称，请检查资料弹窗与坐标偏移")

    def lookup(self, phone, avatar_dir):
        self.search_window()
        self.click("contact_new_friend_search_input1.png", "contact_new_friend_search_input4.png",
                   "contact_new_friend_search_input2.png")
        api.keyevent("^a")
        api.keyevent("{BACKSPACE}")
        pyperclip.copy(phone)
        api.keyevent("^v")
        api.keyevent("{ENTER}")
        time.sleep(3)
        if self.find("contact_new_friend_not_exist_btn_2.png"):
            self.click("contact_new_friend_not_exist_btn_2.png")
            return None, None
        # 展开资料图标，沿用原模板；不点击发送申请按钮。
        x, y = self.click("contact_new_friend_add_btn2.png")
        time.sleep(0.5)
        name = self.nickname(x, y, phone)
        api.touch((x + 250, y + 50))
        time.sleep(2)
        # 原流程截取头像放大后的 300×300 屏幕区域，需随客户端布局校准。
        region = {"left": int(x + 115), "top": int(y + 95), "width": 300, "height": 300}
        avatar_dir = Path(avatar_dir)
        avatar_dir.mkdir(parents=True, exist_ok=True)
        target = avatar_dir / f"{phone}.png"
        with mss.mss() as capture:
            shot = capture.grab(region)
            Image.frombytes("RGB", shot.size, shot.rgb).save(target)
        api.keyevent("{ESC}")
        return name, target
