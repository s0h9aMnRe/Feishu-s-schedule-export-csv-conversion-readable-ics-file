import os
import sys
from aipyapp import runtime

def cleanup_and_rename():
    try:
        # 1. 删除初始版EXE（CSVToICSConverter.exe）
        old_exe = "CSVToICSConverter.exe"
        if os.path.exists(old_exe):
            os.remove(old_exe)
            print(f"✅ 已删除初始版文件：{old_exe}")
        else:
            print(f"ℹ️ 初始版文件 {old_exe} 不存在，无需删除")
        
        # 2. 将终极版EXE重命名为"CSV转换ics工具.exe"
        ultimate_exe = "CSVToICSConverter_Ultimate.exe"
        new_name = "CSV转换ics工具.exe"
        
        if os.path.exists(ultimate_exe):
            # 如果目标文件名已存在，先删除（避免冲突）
            if os.path.exists(new_name):
                os.remove(new_name)
                print(f"ℹ️ 已清理同名文件：{new_name}")
            
            # 执行重命名
            os.rename(ultimate_exe, new_name)
            print(f"✅ 已重命名：{ultimate_exe} → {new_name}")
            
            # 验证结果
            if os.path.exists(new_name):
                print(f"🎉 最终文件：{os.path.abspath(new_name)}")
                runtime.set_state(True, final_exe_path=os.path.abspath(new_name))
            else:
                raise Exception("重命名后文件不存在")
        else:
            raise Exception(f"终极版文件 {ultimate_exe} 不存在")
            
    except Exception as e:
        print(f"❌ 操作失败：{str(e)}", file=sys.stderr)
        runtime.set_state(False, error=str(e))

if __name__ == "__main__":
    cleanup_and_rename()