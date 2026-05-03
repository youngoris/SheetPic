"""
SheetPic 打包脚本
用法:
  macOS:   python3 build.py           → 产出 dist/SheetPic.dmg
  Windows: python build.py            → 产出 dist/SheetPic.exe
  Windows ARM64: python build.py arm64 → 产出 dist/SheetPic.exe
"""
import os
import sys
import subprocess
import platform
import shutil

APP_NAME = "SheetPic"
MAIN_SCRIPT = "sheetpic.py"
DIST_DIR = "dist"
BUILD_DIR = "build"


def run(cmd, check=True):
    print(f"  > {cmd}")
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if result.returncode != 0 and check:
        print(f"  ERROR: {result.stderr}")
        if result.stdout:
            print(f"  STDOUT: {result.stdout}")
        sys.exit(1)
    return result


def build_macos():
    """macOS: PyInstaller → 签名 → DMG"""
    app_path = os.path.join(DIST_DIR, f"{APP_NAME}.app")
    dmg_path = os.path.join(DIST_DIR, f"{APP_NAME}.dmg")

    # 清理旧构建
    for d in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(d):
            shutil.rmtree(d)

    # PyInstaller
    print("=" * 50)
    print("步骤 1: PyInstaller 打包")
    print("=" * 50)
    icon_args = f"--icon=icon.icns" if os.path.exists("icon.icns") else ""
    excludes = " ".join(f"--exclude-module {m}" for m in [
        "tqdm", "setuptools", "unittest", "test", "distutils",
        "lib2to3", "pydoc", "tkinter.test", "numpy.testing",
    ])
    run(f'{sys.executable} -m PyInstaller --windowed --onedir --noconfirm --clean '
        f'--name={APP_NAME} {icon_args} {excludes} {MAIN_SCRIPT}')

    # 写入版本号到 Info.plist
    import re as _re
    version = "1.0.0"
    try:
        with open(MAIN_SCRIPT, "r") as f:
            m = _re.search(r'APP_VERSION\s*=\s*"([^"]+)"', f.read())
            if m:
                version = m.group(1)
    except:
        pass
    plist_path = os.path.join(DIST_DIR, f"{APP_NAME}.app", "Contents", "Info.plist")
    if os.path.exists(plist_path):
        with open(plist_path, "r") as f:
            plist = f.read()
        plist = plist.replace(
            "<string>0.0.0</string>",
            f"<string>{version}</string>", 1
        )
        # 确保有 CFBundleShortVersionString
        if "CFBundleShortVersionString" not in plist:
            plist = plist.replace(
                "</dict>\n</plist>",
                f"\t<key>CFBundleShortVersionString</key>\n\t<string>{version}</string>\n</dict>\n</plist>"
            )
        with open(plist_path, "w") as f:
            f.write(plist)
        print(f"  版本号: {version}")

    # 签名
    print()
    print("=" * 50)
    print("步骤 2: ad-hoc 签名")
    print("=" * 50)
    if os.path.exists(app_path):
        run(f'codesign --force --deep --sign - "{app_path}"')
        result = run(f'codesign --verify --deep --strict "{app_path}"', check=False)
        print(f"  签名{'成功' if result.returncode == 0 else '验证失败'} ✓")
    else:
        print(f"  ERROR: 未找到 {app_path}")
        sys.exit(1)

    # 创建 DMG
    print()
    print("=" * 50)
    print("步骤 3: 创建 DMG")
    print("=" * 50)
    _create_dmg(app_path, dmg_path)

    # 输出结果
    print()
    print("=" * 50)
    print("打包完成")
    print("=" * 50)
    if os.path.exists(dmg_path):
        size = os.path.getsize(dmg_path) / 1024 / 1024
        print(f"  输出: {dmg_path} ({size:.1f} MB)")
        print(f"  用法: 双击 DMG → 拖动 SheetPic 到 Applications 文件夹")


def _create_dmg(app_path, dmg_path):
    """用 hdiutil 创建 DMG（包含 Applications 快捷方式）"""
    # 临时目录
    dmg_staging = os.path.join(DIST_DIR, "dmg_staging")
    if os.path.exists(dmg_staging):
        shutil.rmtree(dmg_staging)
    os.makedirs(dmg_staging)

    # 复制 .app 到临时目录
    shutil.copytree(app_path, os.path.join(dmg_staging, f"{APP_NAME}.app"))

    # 创建 /Applications 的符号链接（用户可拖拽安装）
    os.symlink("/Applications", os.path.join(dmg_staging, "Applications"))

    # 删除旧 DMG
    if os.path.exists(dmg_path):
        os.remove(dmg_path)

    # 用 hdiutil 创建 DMG
    run(f'hdiutil create -volname "{APP_NAME}" '
        f'-srcfolder "{dmg_staging}" '
        f'-ov -format UDZO '
        f'"{dmg_path}"')

    # 清理临时目录
    shutil.rmtree(dmg_staging)
    print(f"  DMG 创建成功 ✓")


def build_windows():
    """Windows: PyInstaller → 单 exe"""
    target_arch = sys.argv[1] if len(sys.argv) > 1 else ""
    arch_label = "ARM64" if target_arch.lower() == "arm64" else "x86_64"

    print(f"平台: Windows ({arch_label})")
    print(f"Python: {sys.version.split()[0]}")
    print(f"打包: {MAIN_SCRIPT} → {APP_NAME}")
    print()

    # 清理旧构建
    for d in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(d):
            shutil.rmtree(d)

    # PyInstaller: onefile 单 exe
    print("=" * 50)
    print("步骤 1: PyInstaller 打包")
    print("=" * 50)
    icon_args = "--icon=icon.ico" if os.path.exists("icon.ico") else ""
    run(f'{sys.executable} -m PyInstaller --windowed --onefile --noconfirm --clean '
        f'--name={APP_NAME} {icon_args} {MAIN_SCRIPT}')

    # 输出结果
    print()
    print("=" * 50)
    print("打包完成")
    print("=" * 50)
    exe_path = os.path.join(DIST_DIR, f"{APP_NAME}.exe")
    if os.path.exists(exe_path):
        size = os.path.getsize(exe_path) / 1024 / 1024
        print(f"  输出: {exe_path} ({size:.1f} MB)")
        print(f"  架构: {arch_label}")
        print(f"  签名: 无（SmartScreen 会弹警告，点仍要运行即可）")


def build():
    system = platform.system()
    print(f"平台: {system}")
    print(f"Python: {sys.version.split()[0]}")
    print(f"打包: {MAIN_SCRIPT} → {APP_NAME}")
    print()

    if system == "Darwin":
        build_macos()
    elif system == "Windows":
        build_windows()
    else:
        print(f"不支持的平台: {system}")
        sys.exit(1)


if __name__ == "__main__":
    build()
