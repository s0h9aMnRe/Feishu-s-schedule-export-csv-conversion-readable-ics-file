import os
import sys
import subprocess
from aipyapp import runtime

def svg_to_ico_and_repackage():
    try:
        # 1. 验证SVG文件是否存在
        svg_path = r"C:\Users\DESKTOP-43TDRJE\aipywork\te37qwwutuw\转换.svg"
        if not os.path.exists(svg_path):
            raise FileNotFoundError(f"SVG图标文件不存在：{svg_path}")
        
        # 2. 安装SVG转ICO所需依赖
        if not runtime.install_packages("cairosvg"):
            raise Exception("安装cairosvg失败，无法转换SVG")
        
        # 3. 将SVG转换为ICO（临时文件）
        import cairosvg
        ico_temp_path = "temp_icon.ico"
        cairosvg.svg2ico(url=svg_path, write_to=ico_temp_path, size=(256, 256))  # 高分辨率图标
        
        # 4. 备份当前EXE（防止转换失败）
        current_exe = "CSV转换ics工具.exe"
        backup_exe = current_exe + ".bak"
        if os.path.exists(current_exe):
            os.rename(current_exe, backup_exe)
            print(f"已备份当前EXE：{backup_exe}")
        
        # 5. 重新打包EXE（使用转换后的ICO）
        source_file = "csv_to_ics_converter.py"
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"源文件 {source_file} 不存在")
        
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile", "--windowed",
            "--name", current_exe,  # 保持原中文名称
            "--distpath", ".",
            "--clean",
            f"--icon={ico_temp_path}",  # 使用转换后的ICO
            source_file
        ]
        
        print(f"正在使用新图标打包：{' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        # 6. 清理临时文件
        if os.path.exists(ico_temp_path):
            os.remove(ico_temp_path)
        
        # 7. 验证结果
        if result.returncode == 0 and os.path.exists(current_exe):
            # 删除备份（成功则无需备份）
            if os.path.exists(backup_exe):
                os.remove(backup_exe)
            print(f"✅ 图标替换成功！新EXE：{os.path.abspath(current_exe)}")
            runtime.set_state(True, exe_path=os.path.abspath(current_exe))
        else:
            # 恢复备份（失败则回滚）
            if os.path.exists(backup_exe):
                os.rename(backup_exe, current_exe)
                print(f"⚠️ 打包失败，已恢复原EXE")
            raise Exception(f"打包失败：{result.stderr}")
            
    except Exception as e:
        error_msg = f"❌ 图标替换失败：{str(e)}"
        print(error_msg, file=sys.stderr)
        runtime.set_state(False, error=error_msg)

if __name__ == "__main__":
    svg_to_ico_and_repackage()