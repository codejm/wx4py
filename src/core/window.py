# -*- coding: utf-8 -*-
"""微信窗口管理"""
import ctypes
import ctypes.wintypes
import time

import win32gui

from .uia_wrapper import UIAWrapper
from .exceptions import WeChatNotFoundError
from ..utils.win32 import (
    find_wechat_window,
    bring_window_to_front,
    get_window_title,
    get_window_class,
    check_and_fix_registry,
    ensure_screen_reader_flag,
    restart_wechat_process,
)
from ..utils.logger import get_logger
from ..config import OPERATION_INTERVAL

logger = get_logger(__name__)

# UIA 健康检查：控件树最少需要的节点数
# 正常微信窗口控件树远超此阈值；如果只有根窗口 + MMUIRenderSubWindowHW = 2 个节点，
# 说明 Qt 辅助功能未加载，需要重启微信。
_MIN_UIA_TREE_NODES = 5


def _count_uia_descendants(ctrl, max_depth=4, limit=20):
    """快速递归统计控件树节点数，用于健康检查。

    Args:
        ctrl: 根控件
        max_depth: 最大递归深度
        limit: 达到此数量后提前返回（无需全部遍历）

    Returns:
        int: 发现的控件节点数
    """
    count = 0
    stack = [(ctrl, 0)]
    while stack:
        node, depth = stack.pop()
        count += 1
        if count >= limit:
            return count
        if depth >= max_depth:
            continue
        try:
            children = node.GetChildren()
            if children:
                for ch in children:
                    stack.append((ch, depth + 1))
        except Exception:
            pass
    return count


# WM_GETOBJECT 消息常量
_WM_GETOBJECT = 0x003D
_OBJID_CLIENT = 0xFFFFFFFC  # -4 的无符号表示

# 唤醒尝试次数和每次等待时间（秒）
_WAKE_ATTEMPTS = 3
_WAKE_INTERVAL = 3


def _force_broadcast_screen_reader_flag():
    """无条件重新设置 SPI_SETSCREENREADER=1 并广播 WM_SETTINGCHANGE。

    即使当前值已经是 1，也重新设置一次，
    目的是让 Qt 收到 WM_SETTINGCHANGE 并重新评估辅助功能状态。
    """
    SPI_SETSCREENREADER = 0x0047
    SPIF_SENDCHANGE = 0x02
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_SETSCREENREADER, 1, 0, SPIF_SENDCHANGE
    )


def _send_wm_getobject(hwnd: int):
    """向窗口及其所有子窗口发送 WM_GETOBJECT 消息。

    Qt 在收到 WM_GETOBJECT(OBJID_CLIENT) 时会延迟初始化
    辅助功能接口（QAccessible），即使启动时没有激活。
    """
    try:
        ctypes.windll.user32.SendMessageW(
            hwnd, _WM_GETOBJECT, 0, _OBJID_CLIENT
        )
    except Exception:
        pass

    def _enum_child(child_hwnd, _):
        try:
            ctypes.windll.user32.SendMessageW(
                child_hwnd, _WM_GETOBJECT, 0, _OBJID_CLIENT
            )
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _enum_child, None)
    except Exception:
        pass


class WeChatWindow:
    """微信窗口管理器"""

    def __init__(self):
        """初始化微信窗口管理器"""
        self._hwnd: int = None
        self._uia: UIAWrapper = None
        self._initialized = False

    def _try_wake_uia(self) -> bool:
        """尝试唤醒微信的 UIA 辅助功能树，不重启微信。

        通过以下步骤尝试触发 Qt 的辅助功能延迟初始化：
        1. 强制重设 SPI_SETSCREENREADER=1 并广播 WM_SETTINGCHANGE
        2. 向微信窗口及子窗口发送 WM_GETOBJECT 消息
        3. 等待 Qt 处理后重新初始化 UIAWrapper
        4. 如果首次不成功，重复多次

        Returns:
            bool: 唤醒成功（UIA 控件树可用）返回 True
        """
        logger.info("尝试唤醒微信 UIA 辅助功能（不重启微信）...")

        for attempt in range(1, _WAKE_ATTEMPTS + 1):
            # 强制广播 SPI 变更通知
            _force_broadcast_screen_reader_flag()
            time.sleep(0.5)

            # 发送 WM_GETOBJECT 触发 Qt 延迟初始化
            _send_wm_getobject(self._hwnd)
            time.sleep(_WAKE_INTERVAL)

            # 重新绑定 UIA 并检查
            try:
                self._uia = UIAWrapper(self._hwnd)
                node_count = _count_uia_descendants(self._uia.root)
                logger.debug(
                    f"唤醒尝试 {attempt}/{_WAKE_ATTEMPTS}: "
                    f"UIA 节点数={node_count}"
                )
                if node_count >= _MIN_UIA_TREE_NODES:
                    logger.info("UIA 唤醒成功，无需重启微信")
                    return True
            except Exception as e:
                logger.debug(f"唤醒尝试 {attempt} 异常: {e}")

        logger.warning("UIA 唤醒失败，微信辅助功能未响应")
        return False

    def _restart_and_reconnect(self):
        """重启微信并等待重新连接。

        流程：
        1. 结束当前微信进程
        2. 等待新进程启动并出现窗口
        3. 重新绑定 UIA

        Raises:
            WeChatNotFoundError: 重启失败或等待超时时抛出
        """
        restarted = restart_wechat_process(self._hwnd)
        self.disconnect()
        if not restarted:
            raise WeChatNotFoundError(
                "辅助功能设置已变更但无法自动重启微信。"
                "请手动重启微信后重试。"
            )

        # 等待微信**主窗口**出现（最多等待 60 秒）
        # 微信重启后先出现 LoginWindow（自动登录中），
        # 登录完成后才会变为 MainWindow，HWND 也会改变。
        logger.info("微信已重启，等待主窗口出现...")
        hwnd = None
        for i in range(60):
            time.sleep(1)
            hwnd = find_wechat_window()
            if hwnd:
                cls = get_window_class(hwnd)
                if 'MainWindow' in cls:
                    logger.debug(f"检测到主窗口: HWND={hwnd}, ClassName={cls}")
                    break
                # 仍然是登录窗口，继续等待
                if i % 5 == 0:
                    logger.debug(f"等待登录完成... 当前窗口: {cls}")
                hwnd = None  # 重置，继续等待主窗口

        if not hwnd:
            # 最后兜底：接受任何微信窗口
            hwnd = find_wechat_window()

        if not hwnd:
            raise WeChatNotFoundError(
                "微信已重启但主窗口未出现，请确认微信已登录后重试。"
            )

        # 等待窗口完全加载
        time.sleep(3)
        bring_window_to_front(hwnd)
        time.sleep(1)

        self._hwnd = hwnd
        logger.info(f"微信重启完成，新窗口: HWND={hwnd}")

        # 重新初始化 UIA，并多次重试健康检查
        # 微信登录完成后 Qt 可能需要额外时间来完全初始化 UIA 控件树
        node_count = 0
        for check in range(5):
            self._uia = UIAWrapper(self._hwnd)
            node_count = _count_uia_descendants(self._uia.root)
            logger.debug(f"重启后 UIA 健康检查 ({check + 1}/5): 节点数={node_count}")
            if node_count >= _MIN_UIA_TREE_NODES:
                return
            # 尝试唤醒并等待
            _force_broadcast_screen_reader_flag()
            _send_wm_getobject(self._hwnd)
            time.sleep(3)

        raise WeChatNotFoundError(
            f"微信重启后 UIA 控件树仍然为空（{node_count} 个节点）。"
            "请确认微信已完全登录并显示主界面后重试。"
        )

    def connect(self) -> bool:
        """
        连接微信窗口。

        流程：
        1. 检查并修复注册表中的 UI Automation 设置
        2. 确保系统屏幕阅读器标志已开启
        3. 查找微信窗口
        4. 将窗口置于前台
        5. 如果设置有变更，重启微信
        6. 初始化 UIAutomation
        7. 健康检查：验证 UIA 控件树是否可用

        Returns:
            bool: 连接成功返回 True

        Raises:
            WeChatNotFoundError: 找不到微信窗口或需要重启微信时抛出
        """
        # 第1步：检查并修复注册表
        logger.info("正在检查注册表中的 UI Automation 设置...")
        registry_modified = False
        try:
            registry_modified = check_and_fix_registry()
            if registry_modified:
                logger.info("注册表 RunningState 已从 0 修改为 1")
            else:
                logger.debug("注册表 RunningState 已正确设置")
        except Exception as e:
            logger.warning(f"注册表检查失败: {e}")

        # 第2步：确保系统屏幕阅读器标志开启
        # Qt 应用（含微信 4.x）在启动时检查此标志，
        # 如果标志关闭则不会创建辅助功能对象。
        screen_reader_changed = False
        try:
            screen_reader_changed = ensure_screen_reader_flag()
            if screen_reader_changed:
                logger.info("系统屏幕阅读器标志原为关闭，已开启")
            else:
                logger.debug("系统屏幕阅读器标志已处于开启状态")
        except Exception as e:
            logger.warning(f"屏幕阅读器标志检查失败: {e}")

        # 如果任一设置被修改，微信需要重启才能生效
        settings_changed = registry_modified or screen_reader_changed

        # 第3步：查找微信窗口
        logger.info("正在查找微信窗口...")
        self._hwnd = find_wechat_window()
        if not self._hwnd:
            raise WeChatNotFoundError(
                "未找到微信窗口，请确保微信正在运行。"
            )

        logger.info(f"找到微信窗口: HWND={self._hwnd}")

        # 第4步：将窗口置于前台
        bring_window_to_front(self._hwnd)
        time.sleep(OPERATION_INTERVAL)

        # 第5步：如果设置有变更，重启微信并自动重连
        if settings_changed:
            logger.warning("辅助功能设置已变更，正在重启微信以使其生效...")
            self._restart_and_reconnect()
            self._initialized = True
            logger.info("成功连接到微信（重启后）")
            return True

        # 第6步：初始化 UIAutomation
        logger.info("正在初始化 UIAutomation...")
        self._uia = UIAWrapper(self._hwnd)

        # 第7步：UIA 健康检查
        # Qt 辅助功能仅在进程启动时根据 SPI_GETSCREENREADER 标志初始化。
        # 如果微信在标志关闭时已经启动，即使之后标志开启了，
        # 当前进程的控件树仍然可能为空。
        # 策略：先尝试通过 WM_GETOBJECT + WM_SETTINGCHANGE 唤醒，
        #        唤醒失败才重启微信。
        node_count = _count_uia_descendants(self._uia.root)
        logger.debug(f"UIA 健康检查: 控件树节点数={node_count}")
        if node_count < _MIN_UIA_TREE_NODES:
            logger.warning(
                f"UIA 控件树几乎为空（仅 {node_count} 个节点），"
                "尝试唤醒微信辅助功能..."
            )
            # 先尝试唤醒，不重启
            if not self._try_wake_uia():
                # 唤醒失败，必须重启微信
                logger.warning("UIA 唤醒失败，微信可能在屏幕阅读器启用前就已启动，正在重启微信...")
                try:
                    ensure_screen_reader_flag()
                except Exception:
                    pass
                self._restart_and_reconnect()

        self._initialized = True
        logger.info("成功连接到微信")
        return True

    def disconnect(self) -> None:
        """断开微信窗口连接"""
        self._hwnd = None
        self._uia = None
        self._initialized = False
        logger.info("已断开微信连接")

    @property
    def hwnd(self) -> int:
        """获取窗口句柄"""
        if not self._initialized:
            raise WeChatNotFoundError("未连接到微信")
        return self._hwnd

    @property
    def uia(self) -> UIAWrapper:
        """获取 UIAutomation 封装器"""
        if not self._initialized:
            raise WeChatNotFoundError("未连接到微信")
        return self._uia

    @property
    def is_connected(self) -> bool:
        """检查是否已连接"""
        return self._initialized and self._hwnd is not None

    @property
    def title(self) -> str:
        """获取窗口标题"""
        if self._hwnd:
            return get_window_title(self._hwnd)
        return ""

    @property
    def class_name(self) -> str:
        """获取窗口类名"""
        if self._hwnd:
            return get_window_class(self._hwnd)
        return ""

    def refresh(self) -> bool:
        """
        刷新微信窗口连接。

        Returns:
            bool: 刷新成功返回 True
        """
        self.disconnect()
        return self.connect()

    def activate(self) -> bool:
        """
        将微信窗口置于前台。

        Returns:
            bool: 成功时返回 True
        """
        if self._hwnd:
            return bring_window_to_front(self._hwnd)
        return False
