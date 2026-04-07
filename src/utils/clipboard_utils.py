# -*- coding: utf-8 -*-
"""剪贴板工具"""
import os
import struct
import win32clipboard
import win32con


def set_files_to_clipboard(file_paths):
    """
    将文件路径以 CF_HDROP 格式设置到剪贴板。

    这允许将文件粘贴到微信等应用程序的聊天输入框中。

    Args:
        file_paths: 单个文件路径字符串或文件路径列表

    Returns:
        bool: 成功时返回 True

    Raises:
        ValueError: 文件路径不存在时抛出
    """
    # 将单个字符串转换为列表
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    # 验证文件路径
    valid_paths = []
    for path in file_paths:
        if os.path.exists(path):
            valid_paths.append(os.path.abspath(path))
        else:
            raise ValueError(f"File not found: {path}")

    if not valid_paths:
        return False

    try:
        # 打开剪贴板
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()

        # 构建 DROPFILES 头部（20 字节）
        # pFiles offset, pt.x, pt.y, fNC, fWide
        offset = 20
        dropfiles_header = struct.pack('<LLLLL', offset, 0, 0, 0, 1)

        # 构建文件路径列表（Unicode，双 null 结尾）
        file_list = []
        for path in valid_paths:
            file_list.append(path.encode('utf-16le'))
            file_list.append(b'\x00\x00')  # null 结束符

        # 额外的双 null 作为列表结束标记
        file_list.append(b'\x00\x00')

        # 组合所有数据
        file_data = b''.join(file_list)
        hdrop_data = dropfiles_header + file_data

        # 设置到剪贴板
        win32clipboard.SetClipboardData(win32con.CF_HDROP, hdrop_data)

        return True

    except Exception:
        return False

    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def set_text_to_clipboard(text: str) -> bool:
    """
    将 Unicode 文本设置到剪贴板。

    Args:
        text: 文本内容

    Returns:
        bool: 成功时返回 True
    """
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        return True
    except Exception:
        return False
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
