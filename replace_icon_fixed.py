import os
import sys
import subprocess
from PIL import Image  # 用于PNG转ICO
import cairosvg  # 用于SVG转PNG
from aipyapp import runtime

def svg_to_ico_and_repackage():
    try:
        # 1. 验证SVG文件是否存在
        svg_path = r"C:\Users\DESKTOP-43TDRJE\aipywork\te37qwwutuw\转换.svg"
        if not os.path.exists(svg_path):
            raise FileNotFoundError(f"SVG图标文件不存在：{svg_path}")
        
        # 2. 安装必要依赖
        if not runtime.install_packages("cairosvg", "Pillow"):
            raise Exception("安装依赖失败，无法转换图标")
        
        # 3. SVG→PNG转换（高分辨率）
        png_temp_path = "temp_icon.png"
        cairosvg.svg2png(url=svg_path, write_to=png_temp_path, output_width=256, output_height=256)
        print(f"✅ SVG转换为PNG：{png_temp_path}")
        
        # 4. PNG→ICO转换（支持多尺寸）
        ico_temp_path = "temp_icon.ico"
        with Image.open(png_temp_path) as img:
            img.save(ico_temp_path, format='ICO', sizes=[(256,256), (128,128), (64,64), (32,32), (16,16)])
        print(f"✅ PNG转换为ICO：{ico_temp_path}")
        
        # 5. 备份当前EXE
        current_exe = "CSV转换ics工具.exe"
        backup_exe = current_exe + ".bak"
        if os.path.exists(current_exe):
            os.rename(current_exe, backup_exe)
            print(f"ℹ️ 已备份当前EXE：{backup_exe}")
        
        # 6. 重新打包EXE（使用新ICO）
        source_file = "csv_to_ics_converter.py"
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"源文件 {source_file} 不存在")
        
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile", "--windowed",
            "--name", current_exe,  # 保持中文名称
            "--distpath", ".",
            "--clean",
            f"--icon={ico_temp_path}",  # 使用转换后的ICO
            source_file
        ]
        
        print(f"正在使用新图标打包：{' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        # 7. 清理临时文件
        for temp in [png_temp_path, ico_temp_path]:
            if os.path.exists(temp):
                os.remove(temp)
        
        # 8. 验证结果
        if result.returncode == 0 and os.path.exists(current_exe):
            if os.path.exists(backup_exe):
                os.remove(backup_exe)  # 成功则删除备份
            print(f"🎉 图标替换成功！新EXE：{os.path.abspath(current_exe)}")
            runtime.set_state(True, exe_path=os.path.abspath(current_exe))
        else:
            if os.path.exists(backup_exe):
                os.rename(backup_exe, current_exe)  # 失败则恢复备份
                print(f"⚠️ 打包失败，已恢复原EXE")
            raise Exception(f"打包失败：{result.stderr[:200]}...")  # 显示部分错误日志
            
    except Exception as e:
        error_msg = f"❌ 图标替换失败：{str(e)}"
        print(error_msg, file=sys.stderr)
        runtime.set_state(False, error=error_msg)

if __name__ == "__main__":
    svg_to_ico_and_repackage()