# -*- coding: utf-8 -*-
"""
DLP 批量解密核心模块（pywin32 COM 另存为落盘版 + WPS PDF 启动器）
机制：透明加密驱动只对白名单进程解密，且**只有"另存为"才会把明文落盘**——
     真机实测（2026-09-29）：
       · 打开加密文件后持续观察 23 秒，磁盘文件始终为密文（88 7D 1C）→ 单纯打开无效；
       · 打开后立即 SaveAs 到明文区，约 3.5 秒即产出明文 → 另存为是唯一可靠路径。
     故流程为：复制密文副本到明文区 → Office(COM) 打开副本 → 另存为到最终路径 → 校验明文。
     Office 必须用 Dispatch（附加已注册实例）而非 DispatchEx（强制新实例），
     实测 DispatchEx 启动的实例不在 DLP 解密路径上，无法触发解密。
     PDF 用命令行启动 wps.exe 打开（wps 在白名单内），依赖驱动落盘。
"""
import glob
import os
import shutil
import subprocess
import time
import uuid
import winreg

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

# COM 类型 → ProgID
COM_PROGID = {"excel": "Excel.Application",
              "word": "Word.Application",
              "ppt": "PowerPoint.Application"}

# PowerPoint 另存为格式（Presentation 无 SaveFormat 属性，只能按扩展名映射）
# 1=ppSaveAsPresentation(.ppt) 24=OpenXMLPresentation(.pptx) 25=MacroEnabled
# 28=OpenXMLSlideShow(.ppsx) 29=OpenXMLSlideShowMacroEnabled(.ppsm)
PPT_SAVE_FORMAT = {".ppt": 1, ".pps": 1, ".pptx": 24, ".pptm": 25,
                   ".ppsx": 28, ".ppsm": 29}

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
PDF_MAX_RETRY = 3              # 单个 PDF 失败重试次数（wps 触发驱动解密是概率性的）

# 另存为后等待明文落盘的校验窗口（秒）
SAVEAS_VERIFY_TIMEOUT = 15
# 另存为失败时的"打开+等待"兜底窗口（秒）——部分环境驱动会在打开时落盘
FALLBACK_WAIT = 10


# ---------- 文件检测 ----------
def _ext(path):
    return os.path.splitext(path)[1].lower()


def is_supported(path):
    return _ext(path) in EXT_MAP


def get_office_type(path):
    return EXT_MAP.get(_ext(path))


def read_header(path, n=8):
    """读取文件头字节，失败返回 None。"""
    try:
        with open(path, "rb") as f:
            return f.read(n)
    except OSError:
        return None


def is_encrypted(path):
    """DLP 加密判定：文件头为 88 7D 1C 特征。"""
    return read_header(path, 3) == ENCRYPTED_MAGIC


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
    依次尝试常见安装路径、注册表 App Paths、PATH，找不到返回 None。结果缓存。
    """
    global _wps_cache, _wps_tried
    if _wps_tried:
        return _wps_cache
    _wps_tried = True

    for pat in WPS_SEARCH_PATTERNS:
        found = glob.glob(os.path.expandvars(pat))
        if found:
            _wps_cache = found[0]
            return _wps_cache
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(root, r"Software\Microsoft\Windows\CurrentVersion"
                                     r"\App Paths\wps.exe") as k:
                v, _ = winreg.QueryValueEx(k, "")
                if v and os.path.exists(v):
                    _wps_cache = v
                    return _wps_cache
        except OSError:
            continue
    _wps_cache = shutil.which("wps.exe")
    return _wps_cache


# ---------- 明文区工具 ----------
def ensure_plaintext_dir():
    """确保明文区存在。"""
    os.makedirs(PLAINTEXT_DIR, exist_ok=True)


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


def make_work_copy(src_path):
    """把源文件复制为明文区临时工作副本（密文），供 Office 打开后另存为。"""
    ensure_plaintext_dir()
    work = os.path.join(
        PLAINTEXT_DIR,
        f"__dlp_work_{os.getpid()}_{uuid.uuid4().hex[:6]}{_ext(src_path)}")
    shutil.copyfile(src_path, work)
    return work


def discard(path):
    """删除临时工作副本（失败静默）。"""
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


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
        except OSError:
            time.sleep(delay)


# ---------- Office 另存为解密 ----------
def _office_open(app, office_type, path):
    """以只读方式打开工作副本，返回文档对象。"""
    if office_type == "excel":
        return app.Workbooks.Open(path, 0, True)              # UpdateLinks=0, ReadOnly
    if office_type == "word":
        return app.Documents.Open(path, False, True, False)  # ConfirmConversions/ReadOnly/AddToRecent
    return app.Presentations.Open(path, True, False, False)  # ReadOnly/Untitled/WithWindow


def _office_save_as(doc, office_type, src_path, out_path):
    """
    另存为到 out_path，保持原文件格式。
    Word/Excel 取文档自身格式常量（SaveFormat/FileFormat），PowerPoint 按扩展名映射。
    """
    ext = _ext(src_path)
    if office_type == "word":
        doc.SaveAs2(out_path, getattr(doc, "SaveFormat", 0) or 0)
    elif office_type == "excel":
        doc.SaveAs(out_path, getattr(doc, "FileFormat", -4143) or -4143)
    else:
        doc.SaveAs(out_path, PPT_SAVE_FORMAT.get(ext, 1))


def _office_close(doc):
    """关闭文档（不保存），失败静默。"""
    for args in ((0,), ()):
        try:
            doc.Close(*args)
            return
        except Exception:
            continue


def _wait_plaintext(path, timeout):
    """等待另存为产生的文件落盘为明文，返回是否成功。"""
    deadline = time.time() + timeout
    while True:
        if is_decrypted_ok(path):
            return True
        if time.time() >= deadline:
            return False
        time.sleep(0.3)


def decrypt_one_office(app, office_type, src_path, out_path):
    """
    单个 Office 文件的解密：工作副本 → 打开 → 另存为 → 校验明文。
    这是本工具真正生效的解密动作（仅打开不会让驱动落盘）。
    返回 (成功?, 消息)。异常由调用方捕获计入失败。
    """
    work = make_work_copy(src_path)
    doc = None
    try:
        doc = _office_open(app, office_type, work)
        _office_save_as(doc, office_type, work, out_path)
        if not _wait_plaintext(out_path, SAVEAS_VERIFY_TIMEOUT):
            # 兜底：个别环境驱动在"打开"阶段落盘，再等一会儿
            if not _wait_plaintext(out_path, FALLBACK_WAIT):
                return False, "另存为后文件仍为密文（驱动未落盘明文）"
        return True, f"解密成功 → {out_path}"
    finally:
        if doc is not None:
            _office_close(doc)
        discard(work)


def office_decrypt_group(office_type, items, log_cb=None, should_stop=None):
    """
    用一个 Office 实例逐个解密一组文件。
    items: [(src, out)]，out 为 None 时自动分配明文区目标名。
    返回 {src: (成功?, out_path, 消息)}。
    """
    progid = COM_PROGID[office_type]
    results = {}
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.Dispatch(progid)
        for attr, val in (("Visible", False), ("DisplayAlerts", False),
                          ("AutomationSecurity", 3)):  # 3=禁用宏，避免宏弹窗
            try:
                setattr(app, attr, val)
            except Exception:
                pass

        for src, out in items:
            if should_stop is not None and should_stop():
                if log_cb:
                    log_cb("⏹ 已停止，剩余文件未处理")
                break
            if out is None:
                out = unique_dest_name(src)
            t0 = time.time()
            try:
                ok, msg = decrypt_one_office(app, office_type, src, out)
            except Exception as e:
                ok, msg = False, f"{type(e).__name__}: {e}"
            results[src] = (ok, out, msg)
            if log_cb:
                mark = "✓" if ok else "✗"
                log_cb(f"    {mark} {os.path.basename(src)}  "
                       f"[{time.time() - t0:.1f}s]")
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
    return results


# ---------- PDF 批量处理（WPS 启动器） ----------
def _wait_exit(proc, secs):
    """等待进程自然退出，返回是否已退出。"""
    if proc is None or proc.poll() is not None:
        return True
    try:
        proc.wait(timeout=secs)
    except Exception:
        pass
    return proc.poll() is not None


def _wait_pdf_decrypted(f, proc, timeout):
    """轮询等待单个 PDF 落盘解密成功，返回是否成功。"""
    deadline = time.time() + timeout
    exited_at = None
    while time.time() < deadline:
        if is_decrypted_ok(f):
            # 二次确认：防止"瞬时明文后被驱动回滚"的误判
            time.sleep(PDF_CONFIRM_DELAY)
            if is_decrypted_ok(f):
                return True
        if proc.poll() is not None:
            if exited_at is None:
                exited_at = time.time()
            # 进程已退出，再给驱动一点落盘时间
            if time.time() - exited_at > PDF_EXIT_GRACE:
                break
        time.sleep(PDF_POLL_INTERVAL)
    return False


def pdf_open_close(pdf_paths, wps_exe, timeout=PDF_POLL_TIMEOUT):
    """
    用 WPS(wps.exe) 逐个打开 PDF，触发 DLP 驱动落盘解密。
    关键经验（真机验证）：
    1) wps 打开 PDF 触发驱动解密是**概率性**的——失败就重试（最多 PDF_MAX_RETRY 次）；
    2) **解密成功后绝不能强杀 wps**（taskkill /F 会导致驱动把文件回滚加密，
       出现"判定成功但文件头仍 88 7D 1C"）；只等它自然退出；
    3) 失败重试前才可强杀残留 wps（文件本就没解密，无回滚损失）；
    4) WPS 是单实例软件——启动下一个文件前必须等上一个进程自然退出。
    注意：Office 侧已实证"仅打开不落盘"，PDF 侧同样存疑——若实测不通，
         需改为 WPS 的"另存为/导出"路径。
    返回 (成功数, 失败数)。
    """
    opened = 0
    last_proc = None

    for f in pdf_paths:
        # 等上一个 wps 自然退出（防单实例劫持）
        _wait_exit(last_proc, 10)
        ok = False
        proc = None
        for _ in range(PDF_MAX_RETRY):
            try:
                proc = subprocess.Popen([wps_exe, f])
            except Exception:
                break
            last_proc = proc
            if _wait_pdf_decrypted(f, proc, timeout):
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


# ---------- 文件收集 ----------
def collect_files(paths):
    """收集待处理文件列表：文件直接加入；文件夹递归扫描支持的文件。"""
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
                files.extend(os.path.join(root, fn) for fn in fnames
                             if is_supported(os.path.join(root, fn)))
    # 去重保序
    seen, result = set(), []
    for f in files:
        key = os.path.abspath(f).lower()
        if key not in seen:
            seen.add(key)
            result.append(f)
    return result


# ---------- 批量解密 ----------
def decrypt_batch(paths, progress_cb=None, log_cb=None, should_stop=None):
    """
    批量解密入口。
    1) 收集文件并过滤出加密文件；
    2) 按类型分组：Office 走 COM 另存为落盘，PDF 走 WPS；
    3) 汇总结果。
    progress_cb(done, total, current_path, phase)  phase: prepare / decrypt
    log_cb(msg)
    返回 (成功数, 失败数, 结果列表 [(src, ok, out, msg)])
    """
    files = collect_files(paths)
    total = len(files)
    if log_cb:
        log_cb(f"共发现 {total} 个支持的文件")

    # 1. 过滤：只处理加密文件
    todo = []
    for i, f in enumerate(files, 1):
        if log_cb:
            log_cb(f"[准备 {i}/{total}] {os.path.basename(f)}")
        if progress_cb:
            progress_cb(i - 1, total, f, "prepare")
        if not is_encrypted(f):
            if log_cb:
                log_cb(f"    ↷ 未加密，跳过: {os.path.basename(f)}")
        else:
            todo.append(f)
        if progress_cb:
            progress_cb(i, total, f, "prepare")

    if not todo:
        if log_cb:
            log_cb("没有需要解密的加密文件。")
        return 0, 0, []

    ensure_plaintext_dir()
    groups = {}
    for f in todo:
        groups.setdefault(get_office_type(f), []).append(f)

    ok = fail = done = 0
    results = []
    work_total = len(todo)

    def _emit(src, out, success, msg):
        nonlocal ok, fail, done
        done += 1
        ok += 1 if success else 0
        fail += 0 if success else 1
        results.append((src, success, out, msg))
        if log_cb:
            log_cb(f"[{done}/{work_total}] {'✓' if success else '✗'} "
                   f"{os.path.basename(src)}")
        if progress_cb:
            progress_cb(done, work_total, src, "decrypt")

    for office_type, srcs in groups.items():
        if should_stop is not None and should_stop():
            if log_cb:
                log_cb("⏹ 已停止，剩余文件未处理")
            for s in srcs:
                _emit(s, None, False, "用户停止")
            continue

        if office_type == "pdf":
            wps_exe = find_wps_exe()
            if not wps_exe:
                if log_cb:
                    log_cb("[pdf] 未找到 WPS(wps.exe)，无法解密 PDF")
                for s in srcs:
                    _emit(s, None, False, "未找到 WPS(wps.exe)，无法解密 PDF")
                continue
            if log_cb:
                log_cb(f"\n[pdf] 启动 WPS 处理 {len(srcs)} 个 PDF…")
            pairs = []
            for s in srcs:
                try:
                    pairs.append((s, make_work_copy(s)))
                except OSError as e:
                    if log_cb:
                        log_cb(f"    ✗ 复制失败：{e}")
                    _emit(s, None, False, f"复制到明文区失败: {e}")
            if pairs:
                try:
                    pdf_open_close([w for _, w in pairs], wps_exe,
                                   timeout=PDF_POLL_TIMEOUT)
                except Exception as e:
                    if log_cb:
                        log_cb(f"    ✗ 启动 WPS 失败：{e}")
                for s, w in pairs:
                    success = is_decrypted_ok(w)
                    _emit(s, w if success else None, success,
                          f"解密成功 → {w}" if success else "解密未完成（PDF 依赖 WPS 落盘）")
                    if success:
                        cleanup_office_lock(w)
                    else:
                        discard(w)
            continue

        # Office：按类型分组，用一个实例逐个"打开→另存为"
        if log_cb:
            log_cb(f"\n[{office_type}] 启动 Office 另存为解密 {len(srcs)} 个文件…")
        outcome = office_decrypt_group(office_type, [(s, None) for s in srcs],
                                       log_cb=log_cb, should_stop=should_stop)
        for s in srcs:
            got = outcome.get(s)
            if got is None:
                _emit(s, None, False, "未处理（已停止或异常）")
                continue
            success, out, msg = got
            cleanup_office_lock(out)
            _emit(s, out if success else None, success, msg)

    if log_cb:
        log_cb(f"\n完成：成功 {ok}，失败 {fail}")
    return ok, fail, results


# ---------- 单文件解密（对外接口） ----------
def decrypt_one(src_path):
    """解密单个文件。返回 (成功?, 目标路径, 消息)。"""
    if not os.path.exists(src_path):
        return False, "", "源文件不存在"
    if not is_supported(src_path):
        return False, "", "不支持的文件类型"
    if not is_encrypted(src_path):
        return False, "", "文件未加密（已是明文）"

    office_type = get_office_type(src_path)
    ensure_plaintext_dir()
    out = unique_dest_name(src_path)

    if office_type == "pdf":
        wps_exe = find_wps_exe()
        if not wps_exe:
            return False, "", "未找到 WPS(wps.exe)，无法解密 PDF"
        try:
            work = make_work_copy(src_path)
        except OSError as e:
            return False, "", f"复制到明文区失败: {e}"
        try:
            pdf_open_close([work], wps_exe, timeout=PDF_POLL_TIMEOUT)
        except Exception as e:
            discard(work)
            return False, "", f"启动 WPS 失败: {e}"
        if is_decrypted_ok(work):
            return True, work, f"解密成功 → {work}"
        discard(work)
        return False, "", "解密失败：文件未被 WPS 解密"

    try:
        outcome = office_decrypt_group(office_type, [(src_path, out)])
    except Exception as e:
        return False, out, f"启动 Office 失败: {e}"
    success, path, msg = outcome.get(src_path, (False, out, "未处理"))
    cleanup_office_lock(path)
    return success, path, msg
