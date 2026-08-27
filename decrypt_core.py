# -*- coding: utf-8 -*-
"""
DLP 批量解密核心模块（pywin32 COM 分组批量版 + WPS PDF 启动器）
机制：透明加密驱动仅对白名单进程解密——Office 文件用微软 Office（COM），
     PDF 文件用 WPS(wps.exe)（在 DLP 白名单内），打开后驱动落盘解密。
加速：按文件类型分组（Excel/Word/PPT/PDF），每组启动一个解密程序实例逐个处理。
注意：Office 必须用 Dispatch（附加已注册实例）而非 DispatchEx（强制新实例），
     实测 DispatchEx 启动的实例不在 DLP 解密路径上，无法触发解密。
     PDF 用命令行启动 wps.exe 打开，wps 读取文件时驱动即落盘解密，wps 随后自行退出。
"""
import os
import shutil
import subprocess
import time

import pythoncom
import win32com.client

# ---------- 常量 ----------
PLAINTEXT_DIR = r"C:\解密文件"
ENCRYPTED_MAGIC = b"\x88\x7d\x1c"          # DLP 加密特征头
OLE_MAGIC = b"\xd0\xcf\x11\xe0"            # 旧版 Office (xls/doc/ppt)
ZIP_MAGIC = b"PK\x03\x04"                  # 新版 Office (xlsx/docx/pptx)
PDF_MAGIC = b"%PDF"                        # PDF 明文特征头

# 扩展名 → 文件类型
EXT_MAP = {
    # Excel
    ".xls": "excel", ".xlsx": "excel", ".xlsm": "excel",
    ".xlsb": "excel", ".xltx": "excel", ".xlt": "excel",
    # Word
    ".doc": "word", ".docx": "word", ".docm": "word",
    ".dotx": "word", ".dotm": "word", ".dot": "word",
    # PowerPoint
    ".ppt": "ppt", ".pptx": "ppt", ".pptm": "ppt",
    ".ppsx": "ppt", ".ppsm": "ppt", ".pps": "ppt",
    # PDF（靠 WPS 的 wps.exe 触发驱动解密）
    ".pdf": "pdf",
}

# COM 类型 → ProgID（仅 Office 使用）
COM_PROGID = {"excel": "Excel.Application",
              "word": "Word.Application",
              "ppt": "PowerPoint.Application"}

# WPS 主程序查找位置（PDF 解密依赖 wps.exe 在白名单内）
WPS_SEARCH_PATTERNS = [
    r"D:\PDF\*\office6\wps.exe",                              # 本机：WPS PDF 模块
    r"D:\WPS Office\*\office6\wps.exe",
    r"C:\Program Files (x86)\Kingsoft\WPS Office\*\office6\wps.exe",
    r"C:\Program Files\Kingsoft\WPS Office\*\office6\wps.exe",
    r"C:\Program Files (x86)\WPS Office\*\office6\wps.exe",
    r"C:\Program Files\WPS Office\*\office6\wps.exe",
    r"%LocalAppData%\Kingsoft\WPS Office\*\office6\wps.exe",
]

# PDF 打开后等待驱动落盘解密的轮询参数
PDF_POLL_INTERVAL = 0.25
PDF_POLL_TIMEOUT = 20          # 单个 PDF 最长等待（秒）
PDF_EXIT_GRACE = 2.5           # wps.exe 退出后仍允许落盘的宽限时间（秒）
PDF_CLOSE_WAIT = 3.0           # 解密成功后等待 wps.exe 自行退出的时间（秒）
PDF_CONFIRM_DELAY = 0.5        # 判定成功后二次确认间隔（过滤瞬时明文被回滚）

# 打开每个文件后等待驱动解密落盘的秒数
OPEN_SETTLE = 0.8


# ---------- 文件检测 ----------
def is_supported(path):
    return os.path.splitext(path)[1].lower() in EXT_MAP


def get_office_type(path):
    return EXT_MAP.get(os.path.splitext(path)[1].lower())


def read_header(path, n=8):
    """读取文件头字节，失败返回 None。"""
    try:
        with open(path, "rb") as f:
            return f.read(n)
    except Exception:
        return None


def is_encrypted(path):
    """DLP 加密判定：文件头为 88 7D 1C 特征。"""
    head = read_header(path, 3)
    return head is not None and head == ENCRYPTED_MAGIC


def is_decrypted_ok(path):
    """解密成功判定：文件头必须是已知明文格式（ZIP/OLE/PDF），
    且不再是加密特征。比"只要不是加密头"更严格，避免空文件/损坏文件误判。"""
    head = read_header(path, 8)
    if head is None or head[:3] == ENCRYPTED_MAGIC:
        return False
    return head[:4] in (ZIP_MAGIC, OLE_MAGIC, PDF_MAGIC)


# ---------- WPS 定位（PDF 解密依赖） ----------
_wps_cache = None
_wps_tried = False


def find_wps_exe():
    """
    定位 WPS 主程序 wps.exe（在 DLP 白名单内，可触发 PDF 落盘解密）。
    依次尝试常见安装路径与注册表，找不到返回 None。结果缓存。
    """
    global _wps_cache, _wps_tried
    if _wps_tried:
        return _wps_cache
    _wps_tried = True

    import glob
    for pat in WPS_SEARCH_PATTERNS:
        try:
            found = glob.glob(os.path.expandvars(pat))
            if found:
                _wps_cache = found[0]
                return _wps_cache
        except Exception:
            continue
    # 注册表 App Paths
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, r"Software\Microsoft\Windows\CurrentVersion\App Paths\wps.exe") as k:
                    v, _ = winreg.QueryValueEx(k, "")
                    if v and os.path.exists(v):
                        _wps_cache = v
                        return _wps_cache
            except Exception:
                continue
    except Exception:
        pass
    # PATH
    try:
        import shutil as _sh
        p = _sh.which("wps.exe")
        if p:
            _wps_cache = p
            return _wps_cache
    except Exception:
        pass
    _wps_cache = None
    return None


# ---------- 明文区工具 ----------
def unique_dest_name(src_path):
    """在明文区生成不重名的目标路径（重名自动加序号）。"""
    name = os.path.basename(src_path)
    base, ext = os.path.splitext(name)
    dest = os.path.join(PLAINTEXT_DIR, name)
    i = 1
    while os.path.exists(dest):
        dest = os.path.join(PLAINTEXT_DIR, f"{base}_{i}{ext}")
        i += 1
    return dest


def prepare_dest(src_path):
    """把源文件复制/定位到明文区目标路径。返回 (dest, 错误消息或None)。"""
    if os.path.dirname(os.path.abspath(src_path)) == os.path.abspath(PLAINTEXT_DIR):
        return src_path, None
    dest = unique_dest_name(src_path)
    try:
        shutil.copyfile(src_path, dest)
        return dest, None
    except Exception as e:
        return None, f"复制到明文区失败: {e}"


def cleanup_office_lock(dest_path, retries=8, delay=0.5):
    """清理 Office 打开文件产生的 ~$ 临时锁文件残留（延迟+重试）。"""
    d = os.path.dirname(dest_path)
    name = os.path.basename(dest_path)
    lock = os.path.join(d, "~$" + name.lstrip("~$"))
    for _ in range(retries):
        try:
            if os.path.exists(lock):
                os.remove(lock)
            return
        except Exception:
            time.sleep(delay)


# ---------- COM 批量处理 ----------
# 单个文件打开后等待驱动落盘解密的最长秒数（打开→轮询文件头→关闭）
COM_FILE_TIMEOUT = 15


def com_open_close(progid, filepaths, settle=OPEN_SETTLE, timeout=COM_FILE_TIMEOUT):
    """
    用 COM（Dispatch）启动一个 Office 实例，逐个打开/关闭一组文件。
    每个文件打开后**轮询文件头**直到驱动落盘解密成功（最长 timeout 秒），
    确认解密完成后再关闭文档——避免 Close 太早导致驱动不再写盘而失败。
    返回 (打开成功数, 失败数)。失败时抛出异常由调用方处理。
    """
    pythoncom.CoInitialize()
    app = None
    opened = 0
    try:
        app = win32com.client.Dispatch(progid)
        # 抑制弹窗
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass

        for f in filepaths:
            try:
                if progid == "Excel.Application":
                    doc = app.Workbooks.Open(f, 0, True)  # ReadOnly
                elif progid == "Word.Application":
                    doc = app.Documents.Open(f, False, True)  # ReadOnly
                else:  # PowerPoint
                    doc = app.Presentations.Open(f, True, False, False)  # ReadOnly
                # 轮询等待驱动解密落盘（文件头 88 7D 1C → 明文 magic）
                deadline = time.time() + max(timeout, 5)
                while time.time() < deadline and not is_decrypted_ok(f):
                    time.sleep(0.3)
                # 关闭文档（不保存）
                try:
                    if progid == "Excel.Application":
                        doc.Close(False)
                    elif progid == "Word.Application":
                        doc.Close(False)
                    else:
                        doc.Close()
                except Exception:
                    pass
                opened += 1
            except Exception:
                continue  # 单个文件失败不中断整组
        return opened, len(filepaths) - opened
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------- PDF 批量处理（WPS 启动器） ----------
# 单个 PDF 解密失败后的重试次数（wps 触发驱动解密是概率性的，重试兜底）
PDF_MAX_RETRY = 3


def pdf_open_close(pdf_paths, wps_exe, timeout=PDF_POLL_TIMEOUT):
    """
    用 WPS(wps.exe) 逐个打开 PDF，触发 DLP 驱动落盘解密。
    关键经验（真机验证）：
    1) wps 打开 PDF 触发驱动解密是**概率性**的——失败就重试（最多 PDF_MAX_RETRY 次）；
    2) **解密成功后绝不能强杀 wps**（taskkill /F 会导致驱动把文件回滚加密，
       出现"判定成功但文件头仍 88 7D 1C"）；只等它自然退出；
    3) 失败重试前才可强杀残留 wps（文件本就没解密，无回滚损失）；
    4) WPS 是单实例软件——启动下一个文件前必须等上一个进程自然退出。
    返回 (成功数, 失败数)。
    """
    opened = 0
    last_proc = None

    def _wait_exit(proc, secs):
        """等待进程自然退出，返回是否已退出。"""
        if proc is None or proc.poll() is not None:
            return True
        try:
            proc.wait(timeout=secs)
        except Exception:
            pass
        return proc.poll() is not None

    for f in pdf_paths:
        # 1. 等上一个 wps 自然退出（防单实例劫持）
        _wait_exit(last_proc, 10)
        ok = False
        proc = None
        for _ in range(PDF_MAX_RETRY):
            try:
                proc = subprocess.Popen([wps_exe, f])
            except Exception:
                break
            last_proc = proc
            deadline = time.time() + timeout
            done = False
            exited_at = None
            while time.time() < deadline:
                if is_decrypted_ok(f):
                    # 二次确认：防止"瞬时明文后被驱动回滚"的误判
                    time.sleep(PDF_CONFIRM_DELAY)
                    if is_decrypted_ok(f):
                        done = True
                        break
                if proc.poll() is not None:
                    if exited_at is None:
                        exited_at = time.time()
                    # 进程已退出，再给驱动一点落盘时间
                    if time.time() - exited_at > PDF_EXIT_GRACE:
                        break
                time.sleep(PDF_POLL_INTERVAL)
            if done:
                ok = True
                break
            # 失败：等进程自然退出（超时才强杀；文件没解密，强杀无回滚损失）
            if not _wait_exit(proc, 8):
                try:
                    subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                                   capture_output=True, timeout=10)
                except Exception:
                    pass
                time.sleep(0.5)
        if ok:
            opened += 1
            # 解密成功：等 wps 自然退出（绝不强杀，防驱动回滚）
            _wait_exit(proc, PDF_CLOSE_WAIT)
    # 收尾：等最后一个 wps 自然退出
    _wait_exit(last_proc, 10)
    return opened, len(pdf_paths) - opened


# ---------- 批量解密 ----------
def decrypt_batch(paths, progress_cb=None, log_cb=None, timeout=120):
    """
    批量解密入口（COM 分组批量版）。
    1) 收集文件并复制到明文区；
    2) 按 Office 类型分组，每组 COM 启动一个 Office 实例逐个处理；
    3) 处理完统一轮询确认解密落盘、清理锁文件。
    progress_cb(done, total, current_path)
    log_cb(msg)
    返回 (成功数, 失败数, 结果列表)
    """
    files = collect_files(paths)
    total = len(files)
    if log_cb:
        log_cb(f"共发现 {total} 个支持的文件")

    # 1. 过滤：只处理加密文件，并准备明文区副本
    prepared = []  # [(src, dest)]
    for i, f in enumerate(files, 1):
        if log_cb:
            log_cb(f"[准备 {i}/{total}] {os.path.basename(f)}")
        if progress_cb:
            progress_cb(i - 1, total, f)
        if not is_encrypted(f):
            if log_cb:
                log_cb(f"    ↷ 未加密，跳过: {os.path.basename(f)}")
            if progress_cb:
                progress_cb(i, total, f)
            continue
        dest, err = prepare_dest(f)
        if err:
            if log_cb:
                log_cb(f"    ✗ {err}")
            if progress_cb:
                progress_cb(i, total, f)
            continue
        prepared.append((f, dest))
        if progress_cb:
            progress_cb(i, total, f)

    if not prepared:
        if log_cb:
            log_cb("没有需要解密的加密文件。")
        return 0, 0, []

    # 2. 按 Office 类型分组
    groups = {}
    for s, d in prepared:
        t = get_office_type(s)
        groups.setdefault(t, []).append((s, d))

    ok = fail = 0
    results = []
    done_count = 0
    total_work = len(prepared)

    for office_type, items in groups.items():
        dests = [d for _, d in items]
        if office_type == "pdf":
            # PDF：用 WPS(wps.exe) 触发驱动解密（wps 在白名单内）
            wps_exe = find_wps_exe()
            if not wps_exe:
                for s, d in items:
                    fail += 1
                    done_count += 1
                    results.append((s, False, d, "未找到 WPS(wps.exe)，无法解密 PDF"))
                    if log_cb:
                        log_cb(f"    ✗ {os.path.basename(s)}")
                    if progress_cb:
                        progress_cb(done_count, total_work, s)
                continue
            if log_cb:
                log_cb(f"\n[pdf] 启动 WPS 处理 {len(dests)} 个 PDF…")
            try:
                pdf_open_close(dests, wps_exe, timeout=PDF_POLL_TIMEOUT)
            except Exception as e:
                for s, d in items:
                    fail += 1
                    done_count += 1
                    results.append((s, False, d, f"启动 WPS 失败: {e}"))
                    if log_cb:
                        log_cb(f"    ✗ {os.path.basename(s)}")
                    if progress_cb:
                        progress_cb(done_count, total_work, s)
                continue
        else:
            progid = COM_PROGID[office_type]
            if log_cb:
                log_cb(f"\n[{office_type}] 启动 Office（COM）处理 {len(dests)} 个文件…")
            try:
                _opened, _failed = com_open_close(progid, dests, timeout=COM_FILE_TIMEOUT)
            except Exception as e:
                for s, d in items:
                    fail += 1
                    done_count += 1
                    results.append((s, False, d, f"启动 Office 失败: {e}"))
                    if log_cb:
                        log_cb(f"    ✗ {os.path.basename(s)}")
                    if progress_cb:
                        progress_cb(done_count, total_work, s)
                continue

        # 3. 处理完后统一轮询确认解密落盘（COM 关闭后驱动已落盘，这里做兜底确认）
        deadline = time.time() + 30
        pending = list(items)
        while pending and time.time() < deadline:
            still = []
            for s, d in pending:
                if is_decrypted_ok(d):
                    ok += 1
                    results.append((s, True, d, f"解密成功 → {d}"))
                    if log_cb:
                        log_cb(f"    ✓ {os.path.basename(s)}")
                else:
                    still.append((s, d))
            pending = still
            if pending:
                time.sleep(0.5)
        for s, d in pending:
            fail += 1
            results.append((s, False, d, "解密未完成（文件可能已损坏或未被 Office 读取）"))
            if log_cb:
                log_cb(f"    ✗ {os.path.basename(s)}")
        # 清理锁文件
        for _, d in items:
            cleanup_office_lock(d)
            done_count += 1
            if progress_cb:
                progress_cb(done_count, total_work, d)

    if log_cb:
        log_cb(f"\n完成：成功 {ok}，失败 {fail}")
    return ok, fail, results


# ---------- 文件收集 ----------
def collect_files(paths):
    """收集待处理文件列表：文件直接加入；文件夹递归扫描支持的 Office 文件。"""
    files = []
    for p in paths:
        p = p.strip().strip('"')
        if not p or not os.path.exists(p):
            continue
        if os.path.isfile(p):
            if is_supported(p):
                files.append(p)
        elif os.path.isdir(p):
            for root, _, fnames in os.walk(p):
                for fn in fnames:
                    full = os.path.join(root, fn)
                    if is_supported(full):
                        files.append(full)
    # 去重保序
    seen, result = set(), []
    for f in files:
        key = os.path.abspath(f).lower()
        if key not in seen:
            seen.add(key)
            result.append(f)
    return result


# ---------- 单文件解密（兼容旧调用） ----------
def decrypt_one(src_path, timeout=60, poll_interval=1.0):
    """解密单个文件（复用 COM 批量逻辑）。返回 (成功?, 目标路径, 消息)。"""
    if not os.path.exists(src_path):
        return False, "", "源文件不存在"
    if not is_supported(src_path):
        return False, "", "不支持的文件类型"
    if not is_encrypted(src_path):
        return False, "", "文件未加密（已是明文）"
    dest, err = prepare_dest(src_path)
    if err:
        return False, "", err
    office_type = get_office_type(src_path)
    if office_type == "pdf":
        wps_exe = find_wps_exe()
        if not wps_exe:
            cleanup_office_lock(dest)
            return False, dest, "未找到 WPS(wps.exe)，无法解密 PDF"
        try:
            pdf_open_close([dest], wps_exe, timeout=timeout)
        except Exception as e:
            cleanup_office_lock(dest)
            return False, dest, f"启动 WPS 失败: {e}"
        time.sleep(1)
        if is_decrypted_ok(dest):
            cleanup_office_lock(dest)
            return True, dest, f"解密成功 → {dest}"
        cleanup_office_lock(dest)
        return False, dest, "解密失败：文件未被 WPS 解密"
    progid = COM_PROGID[office_type]
    try:
        com_open_close(progid, [dest], timeout=timeout)
    except Exception as e:
        cleanup_office_lock(dest)
        return False, dest, f"启动 Office 失败: {e}"
    time.sleep(1)
    if is_decrypted_ok(dest):
        cleanup_office_lock(dest)
        return True, dest, f"解密成功 → {dest}"
    cleanup_office_lock(dest)
    return False, dest, "解密失败：文件未被 Office 解密"
