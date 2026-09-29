# -*- coding: utf-8 -*-
"""
右键菜单注册/卸载模块（DLP 一键解密）
- 写入 HKCU 的 Software\\Classes，无需管理员权限；
- 对每个支持的扩展名，同时在"扩展名键"和"其 ProgID 键"两处注册 shell 动词，
  最大化右键菜单显示兼容性；
- 命令指向工具自身（打包后为 exe），右键后把 "%*" 传给它——
  %* 会把本次选中的所有文件路径一次性传入（%1 则每个文件单独启动一次，
  导致选中 N 个文件开 N 个窗口）。
"""
import os
import sys
import winreg

# 菜单名称（右键菜单里显示的文字）
MENU_NAME = "DLP 一键解密"

# 需要注册右键菜单的扩展名（与 decrypt_core.EXT_MAP 保持一致）
EXTENSIONS = [
    # Excel
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xlt",
    # Word
    ".doc", ".docx", ".docm", ".dotx", ".dotm", ".dot",
    # PowerPoint
    ".ppt", ".pptx", ".pptm", ".ppsx", ".ppsm", ".pps",
    # PDF
    ".pdf",
]

# 基础键（HKCU\Software\Classes 之下）
_BASE = r"Software\Classes"


def get_tool_path():
    """
    返回工具可执行文件路径。
    打包后（PyInstaller）：sys.executable 即 exe 本身。
    脚本运行时：返回本脚本，配合 python 使用。
    """
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.abspath(sys.argv[0])


def _command():
    """右键菜单执行命令：用工具打开被右键的文件。
    用 %*（全部选中文件一次传入）而非 %1（每个文件单独启动一次）。
    """
    return '"%s" "%%*"' % get_tool_path()


def _progid_of(ext):
    """读取扩展名关联的 ProgID（HKCU 优先，其次 HKLM）。"""
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(root, _BASE + "\\" + ext) as k:
                v, _ = winreg.QueryValueEx(k, "")
                if v:
                    return v
        except OSError:
            continue
    return None


def _userchoice_progid(ext):
    """读取用户选择的默认打开程序 ProgID（如 Applications\\PDFXEdit.exe）。"""
    key = (r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
           + "\\" + ext + r"\UserChoice")
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            v, _ = winreg.QueryValueEx(k, "ProgId")
            return v or None
    except OSError:
        return None


def _target_keys(ext):
    """
    返回该扩展名需要注册菜单的所有基础键（相对 _BASE）。
    1) 扩展名键本身（如 .pdf、.docx）
    2) 其关联 ProgID 键（如 Word.Document.12）
    3) 用户默认打开程序的 ProgID 键（如 Applications\\PDFXEdit.exe，覆盖无 ProgID 的扩展名）
    """
    keys = [ext]
    progid = _progid_of(ext)
    if progid:
        keys.append(progid)
    uch = _userchoice_progid(ext)
    if uch and uch not in keys:
        keys.append(uch)
    return keys


def register():
    """注册右键菜单。返回 (成功数, 失败原因列表)。"""
    ok = 0
    fails = []
    cmd = _command()
    tool = get_tool_path()
    for ext in EXTENSIONS:
        for base_key in _target_keys(ext):
            try:
                menu_key = _BASE + "\\" + base_key + "\\shell\\" + MENU_NAME
                # 命令
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                                      menu_key + "\\command") as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, cmd)
                # 图标（可选，用工具自身图标）
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, menu_key,
                                    0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ,
                                      '"%s",0' % tool)
                ok += 1
            except OSError as e:
                fails.append("%s: %s" % (ext, e))
    return ok, fails


def unregister():
    """卸载右键菜单。返回 (移除数, 失败原因列表)。"""
    ok = 0
    fails = []
    for ext in EXTENSIONS:
        for base_key in _target_keys(ext):
            menu_key = _BASE + "\\" + base_key + "\\shell\\" + MENU_NAME
            try:
                try:
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                                     menu_key + "\\command")
                except FileNotFoundError:
                    pass
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, menu_key)
                ok += 1
            except FileNotFoundError:
                continue  # 本就没注册过
            except OSError as e:
                fails.append("%s: %s" % (ext, e))
    return ok, fails


if __name__ == "__main__":
    # 命令行使用：python context_menu.py [register|unregister]
    action = sys.argv[1] if len(sys.argv) > 1 else "register"
    if action == "unregister":
        n, _ = unregister()
        print("已移除 %d 个右键菜单项" % n)
    else:
        n, _ = register()
        print("已注册 %d 个右键菜单项" % n)
    if fs:
        print("失败：", fs)
