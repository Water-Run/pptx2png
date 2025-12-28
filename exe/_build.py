"""
pptx2png 打包脚本
使用 PyInstaller 将程序打包为单个 .exe 文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

# 配置
APP_NAME = "pptx2png"
VERSION = "2025.1"
AUTHOR = "WaterRun"
DESCRIPTION = "PowerPoint to PNG Converter"

# 路径配置
SCRIPT_DIR = Path(__file__).parent.absolute()
ROOT_DIR = SCRIPT_DIR.parent
LOGO_PATH = ROOT_DIR / "logo.png"
MAIN_SCRIPT = SCRIPT_DIR / "pptx2png-exe.py"
OUTPUT_DIR = SCRIPT_DIR / "dist"
BUILD_DIR = SCRIPT_DIR / "build"
SPEC_FILE = SCRIPT_DIR / f"{APP_NAME}.spec"

def print_banner():
    """打印横幅"""
    banner = f"""
╔════════════════════════════════════════════════════════════╗
║                    pptx2png 打包工具                       ║
║                      Version {VERSION}                         ║
║                    by {AUTHOR}                           ║
╚════════════════════════════════════════════════════════════╝
"""
    print(banner)

def check_requirements():
    """检查依赖"""
    print("📦 检查依赖...")
    
    # 检查 PyInstaller
    try:
        import PyInstaller
        print(f"  ✓ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("  ✗ PyInstaller 未安装")
        print("\n正在安装 PyInstaller...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        print("  ✓ PyInstaller 安装完成")
    
    # 检查必要文件
    if not LOGO_PATH.exists():
        print(f"  ✗ Logo 文件不存在: {LOGO_PATH}")
        return False
    else:
        print(f"  ✓ Logo 文件: {LOGO_PATH}")
    
    if not MAIN_SCRIPT.exists():
        print(f"  ✗ 主脚本不存在: {MAIN_SCRIPT}")
        return False
    else:
        print(f"  ✓ 主脚本: {MAIN_SCRIPT}")
    
    return True

def create_version_file():
    """创建 Windows 版本信息文件"""
    version_file = SCRIPT_DIR / "version_info.txt"
    
    version_content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(2025, 1, 0, 0),
    prodvers=(2025, 1, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'{AUTHOR}'),
        StringStruct(u'FileDescription', u'{DESCRIPTION}'),
        StringStruct(u'FileVersion', u'{VERSION}'),
        StringStruct(u'InternalName', u'{APP_NAME}'),
        StringStruct(u'LegalCopyright', u'Copyright (C) {AUTHOR} 2025'),
        StringStruct(u'OriginalFilename', u'{APP_NAME}.exe'),
        StringStruct(u'ProductName', u'{APP_NAME}'),
        StringStruct(u'ProductVersion', u'{VERSION}')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""
    
    with open(version_file, 'w', encoding='utf-8') as f:
        f.write(version_content)
    
    return version_file

def clean_build():
    """清理构建目录"""
    print("\n🧹 清理旧文件...")
    
    dirs_to_clean = [BUILD_DIR, OUTPUT_DIR]
    for dir_path in dirs_to_clean:
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"  ✓ 已删除: {dir_path.name}/")
    
    if SPEC_FILE.exists():
        SPEC_FILE.unlink()
        print(f"  ✓ 已删除: {SPEC_FILE.name}")

def build_executable():
    """构建可执行文件"""
    print("\n🔨 开始打包...")
    
    # 创建版本信息文件
    version_file = create_version_file()
    
    # PyInstaller 参数
    args = [
        'pyinstaller',
        '--name', APP_NAME,
        '--onefile',
        '--windowed',
        '--clean',
        '--noconfirm',
        f'--icon={LOGO_PATH}',
        '--add-data', f'{LOGO_PATH};.',  # 将 logo.png 添加到根目录
        '--optimize', '2',
        '--version-file', str(version_file),
        str(MAIN_SCRIPT)
    ]
    
    # 运行 PyInstaller
    try:
        result = subprocess.run(args, check=True, capture_output=True, text=True)
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失败:")
        print(e.stderr)
        return False
    finally:
        # 清理版本文件
        if version_file.exists():
            version_file.unlink()

def copy_to_root():
    """将生成的 exe 复制到脚本同目录"""
    print("\n📁 移动文件...")
    
    source_exe = OUTPUT_DIR / f"{APP_NAME}.exe"
    target_exe = SCRIPT_DIR / f"{APP_NAME}.exe"
    
    if source_exe.exists():
        # 如果目标文件存在，先删除
        if target_exe.exists():
            target_exe.unlink()
        
        # 复制文件
        shutil.copy2(source_exe, target_exe)
        print(f"  ✓ 已复制到: {target_exe}")
        
        # 获取文件大小
        size_mb = target_exe.stat().st_size / (1024 * 1024)
        print(f"  ✓ 文件大小: {size_mb:.2f} MB")
        
        return True
    else:
        print(f"  ✗ 未找到生成的文件: {source_exe}")
        return False

def cleanup_after_build():
    """构建后清理"""
    print("\n🧹 清理临时文件...")
    
    # 保留 exe，删除其他构建文件
    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)
        print("  ✓ 已删除: build/")
    
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
        print("  ✓ 已删除: dist/")
    
    if SPEC_FILE.exists():
        SPEC_FILE.unlink()
        print("  ✓ 已删除: .spec 文件")

def main():
    """主函数"""
    print_banner()
    
    # 检查依赖
    if not check_requirements():
        print("\n❌ 依赖检查失败，请修复后重试")
        return 1
    
    # 清理旧文件
    clean_build()
    
    # 构建
    if not build_executable():
        print("\n❌ 构建失败")
        return 1
    
    # 复制文件
    if not copy_to_root():
        print("\n❌ 文件复制失败")
        return 1
    
    # 清理临时文件
    cleanup_after_build()
    
    # 完成
    print("\n" + "="*60)
    print("✅ 打包完成！")
    print("="*60)
    print(f"\n📦 输出文件: {SCRIPT_DIR / f'{APP_NAME}.exe'}")
    print("\n提示:")
    print("  - 首次运行可能被杀毒软件拦截，请添加信任")
    print("  - 确保目标机器已安装 Microsoft PowerPoint")
    print("\n🎉 祝您使用愉快！\n")
    
    return 0

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)