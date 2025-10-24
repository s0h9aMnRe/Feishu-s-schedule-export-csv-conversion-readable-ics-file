import os
import glob
import sys
from aipyapp import runtime

def improved_cleanup():
    try:
        # 要保留的文件（严格控制）
        keep_files = {
            "csv_to_ics_converter.py",  # 终极版源代码
            "CSVToICSConverter_Ultimate.exe"  # 终极版可执行文件（如果已存在）
        }
        
        # 删除所有旧版.py文件
        for file in glob.glob("*.py"):
            if file not in keep_files and not file.startswith("cleanup_") and not file.startswith("package_"):
                try:
                    if os.path.exists(file):
                        os.remove(file)
                        print(f"✅ 删除旧版Python文件：{file}")
                except Exception as e:
                    print(f"⚠️ 无法删除 {file}（可能被占用）：{str(e)}")
        
        # 删除所有旧版.exe文件（除了要保留的终极版）
        for file in glob.glob("CSVToICSConverter_*.exe"):
            if file not in keep_files:
                try:
                    if os.path.exists(file):
                        os.remove(file)
                        print(f"✅ 删除旧版EXE文件：{file}")
                except Exception as e:
                    print(f"⚠️ 无法删除 {file}（可能被占用）：{str(e)}")
        
        # 删除打包残留文件和目录
        for item in ["build", "dist", "__pycache__", "CSVToICSConverter_Ultimate.spec"]:
            try:
                if os.path.exists(item):
                    if os.path.isdir(item):
                        import shutil
                        shutil.rmtree(item, ignore_errors=True)
                    else:
                        os.remove(item)
                    print(f"✅ 清理残留文件：{item}")
            except Exception as e:
                print(f"⚠️ 无法清理 {item}：{str(e)}")
        
        print("✅ 主要清理完成，已保留最新版核心文件")
        runtime.set_state(True)
        
    except Exception as e:
        print(f"❌ 清理失败：{str(e)}", file=sys.stderr)
        runtime.set_state(False, error=str(e))

if __name__ == "__main__":
    improved_cleanup()