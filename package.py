#!/usr/bin/env python3
"""
package.py — 将 build_ppt.py 打包为独立可执行文件。

用法：
    python package.py              # 为当前平台打包
    python package.py --clean      # 清理构建缓存后打包

依赖 PyInstaller，首次运行会自动安装。
注意：PyInstaller 不支持跨平台编译，请在目标 OS 上运行本脚本。
"""

import os
import sys
import subprocess
import platform
import shutil


PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
ENTRY_SCRIPT = os.path.join(PROJECT_ROOT, "build_ppt.py")
SPEC_FILE    = os.path.join(PROJECT_ROOT, "build_ppt.spec")
DIST_DIR     = os.path.join(PROJECT_ROOT, "dist")
BUILD_DIR    = os.path.join(PROJECT_ROOT, "build")

# ── 根据平台确定输出文件名 ──────────────────────────────────

IS_WINDOWS = platform.system() == "Windows"
EXE_NAME   = "build_ppt.exe" if IS_WINDOWS else "build_ppt"


# ── 依赖检查与安装 ──────────────────────────────────────────

def _run(cmd: list[str], desc: str):
    """运行命令并打印进度。"""
    print(f"\n> {desc}")
    print(f"  {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    if result.returncode != 0:
        print(f"\n错误：{desc} 失败（退出码 {result.returncode}）")
        sys.exit(result.returncode)


def ensure_pyinstaller():
    """确保 PyInstaller 已安装。"""
    try:
        import PyInstaller  # noqa: F401
        print("[信息] PyInstaller 已安装")
    except ImportError:
        print("[信息] 正在安装 PyInstaller...")
        _run([sys.executable, "-m", "pip", "install", "pyinstaller"],
             "安装 PyInstaller")


# ── 打包 ────────────────────────────────────────────────────

def clean_build():
    """清理上一次的构建产物。"""
    for p in [SPEC_FILE, BUILD_DIR, DIST_DIR]:
        if os.path.isdir(p):
            shutil.rmtree(p)
            print(f"[清理] 删除目录 {os.path.basename(p)}/")
        elif os.path.isfile(p):
            os.remove(p)
            print(f"[清理] 删除文件 {os.path.basename(p)}")


def package():
    """使用 PyInstaller 打包为单文件可执行程序。"""
    ensure_pyinstaller()

    # 构建命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # 单文件输出
        "--name", EXE_NAME.replace(".exe", ""),
        "--hidden-import", "lxml.etree",
        "--hidden-import", "PIL.Image",
        "--clean",
        ENTRY_SCRIPT,
    ]

    _run(cmd, "PyInstaller 打包中")

    # ── 验证输出 ──
    output_exe = os.path.join(DIST_DIR, EXE_NAME)
    if os.path.isfile(output_exe):
        size_mb = os.path.getsize(output_exe) / (1024 * 1024)
        print(f"\n{'=' * 50}")
        print(f"  打包成功！")
        print(f"  输出文件: {output_exe}")
        print(f"  文件大小: {size_mb:.1f} MB")
        print(f"{'=' * 50}")
    else:
        print(f"\n错误：未找到输出文件 {output_exe}")
        sys.exit(1)


# ── 入口 ────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--clean" in sys.argv:
        clean_build()
    package()
