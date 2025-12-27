import os
import sys
import subprocess
import shutil

def build_exe():
    """Build pptx2img.exe using PyInstaller"""
    
    print("=" * 60)
    print("Building pptx2img.exe")
    print("=" * 60)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("Error: PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Get script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_dir, "pptx2img.py")
    
    # Build command
    build_cmd = [
        "pyinstaller",
        "--onefile",                          # Single file executable
        "--windowed",                         # No console window
        "--name=pptx2img",                    # Output name
        "--icon=NONE",                        # No icon (can be added later)
        "--clean",                            # Clean cache
        script_path
    ]
    
    print("\nRunning PyInstaller...")
    print(" ".join(build_cmd))
    print()
    
    try:
        subprocess.check_call(build_cmd)
        
        # Move the exe to current directory
        dist_dir = os.path.join(script_dir, "dist")
        exe_path = os.path.join(dist_dir, "pptx2img.exe")
        target_path = os.path.join(script_dir, "pptx2img.exe")
        
        if os.path.exists(exe_path):
            shutil.copy2(exe_path, target_path)
            print(f"\n✓ Build successful!")
            print(f"✓ Executable created: {target_path}")
            
            # Cleanup
            print("\nCleaning up build files...")
            build_dir = os.path.join(script_dir, "build")
            spec_file = os.path.join(script_dir, "pptx2img.spec")
            
            if os.path.exists(build_dir):
                shutil.rmtree(build_dir)
            if os.path.exists(dist_dir):
                shutil.rmtree(dist_dir)
            if os.path.exists(spec_file):
                os.remove(spec_file)
            
            print("✓ Cleanup complete!")
            
        else:
            print(f"\n✗ Build failed: {exe_path} not found")
            return False
        
    except subprocess.CalledProcessError as e:
        print(f"\n✗ Build failed with error: {e}")
        return False
    
    print("\n" + "=" * 60)
    print("Build process completed!")
    print("=" * 60)
    return True


if __name__ == "__main__":
    build_exe()