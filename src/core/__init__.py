# -*- coding: utf-8 -*-
"""核心模块 - 窗口管理与 UIAutomation 封装"""

from .window import WeChatWindow
from .uia_wrapper import UIAWrapper
from .exceptions import (
    WeChatError,
    WeChatNotFoundError,
    WeChatNotConnectedError,
    UIAError,
    ControlNotFoundError,
    TargetNotFoundError,
    RegistryError,
)

__all__ = [
    "WeChatWindow",
    "UIAWrapper",
    "WeChatError",
    "WeChatNotFoundError",
    "WeChatNotConnectedError",
    "UIAError",
    "ControlNotFoundError",
    "TargetNotFoundError",
    "RegistryError",
]
