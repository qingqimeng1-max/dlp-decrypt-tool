# DLP 批量解密工具

针对透明加密（DLP）文件系统（例如绿盾）的**批量解密工具**。把加密的 Excel / Word / PPT / PDF 文件解密为明文副本，输出到明文区 `C:\解密文件`。

## 功能特性

- **四种格式全支持**：Excel（xls/xlsx/xlsm/xlsb/xlt/xltx）、Word（doc/docx/docm/dot/dotx/dotm）、PowerPoint（ppt/pptx/pptm/pps/ppsx/ppsm）、PDF
- **批量解密**：一次处理多个文件/整个文件夹，按类型分组并行驱动解密程序
- **右键菜单一键解密**：在资源管理器中选中文件右键 → 「DLP 一键解密」，支持多选（单实例，只开一个窗口）
- **企业极简 UI**：近黑主色、状态徽章、进度条、应用内通知，无弹窗打断
- **单实例架构**：全系统只有一个主窗口，后续右键操作自动转发给已开窗口

## 工作原理

DLP 透明加密驱动只对**白名单进程**解密：

| 文件类型 | 解密程序 | 触发方式 |
|---|---|---|
| Excel/Word/PPT | 微软 Office | COM（Dispatch）打开 → 驱动落盘明文 |
| PDF | WPS `wps.exe` | 命令行启动 → 驱动落盘明文（`wpspdf.exe` 不在白名单） |

工具把源文件复制到明文区后，用白名单程序打开一次，驱动即把副本解密为明文。

## 使用说明

### 直接使用（无需安装）

1. 下载 `dist/DLP批量解密工具.exe`
2. 双击运行，拖入文件/文件夹，或点「添加文件」，再点「开始解密」
3. 解密副本输出到 `C:\解密文件`
4. 首次使用可双击 `注册右键菜单.bat` 启用右键菜单

### 从源码运行

```bash
pip install pywin32 tkinterdnd2 pillow
python decrypt_gui.py
```

### 重新打包 exe

```bash
pip install pyinstaller
pyinstaller --noconfirm --clean "DLP批量解密工具.spec"
# 产物：dist/DLP批量解密工具.exe
```

## 项目结构

```
batch_decrypt/
├── decrypt_core.py        # 解密核心（COM 批量 + WPS PDF 启动器 + 重试机制）
├── decrypt_gui.py         # 图形界面（企业极简风格 + 单实例锁 + 队列转发）
├── context_menu.py        # 右键菜单注册/卸载（HKCU，免管理员）
├── DLP批量解密工具.spec   # PyInstaller 打包配置
├── 注册右键菜单.bat        # 一键注册右键菜单
├── 卸载右键菜单.bat        # 一键卸载右键菜单
└── dist/
    └── DLP批量解密工具.exe # 打包好的可执行文件
```

## 已知限制

- PDF 解密依赖本机安装 WPS（`wps.exe`），且该程序必须在 DLP 白名单内
- 右键菜单指向 exe 当前位置，移动 exe 后需重新注册（双击 `注册右键菜单.bat`）
- 输出目录固定为 `C:\解密文件`（可在 `decrypt_core.py` 的 `PLAINTEXT_DIR` 修改）
