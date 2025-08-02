# uninstall_backup.py
import os
import sys
import shutil
from datetime import datetime

def backup_database_on_uninstall():
    """在卸载时备份数据库文件"""
    try:
        # 确定数据库文件路径
        appdata_path = os.getenv('APPDATA')
        if not appdata_path:
            print("无法获取APPDATA环境变量")
            return False
        
        db_dir = os.path.join(appdata_path, "ReagentManagementSystem")
        db_file = os.path.join(db_dir, "ReagentWarehouseData.db")
        
        if not os.path.exists(db_file):
            print("数据库文件不存在，无需备份")
            return True
        
        # 创建备份目录
        backup_dir = os.path.join(db_dir, "Backups")
        os.makedirs(backup_dir, exist_ok=True)
        
        # 生成带时间戳的备份文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(backup_dir, f"ReagentWarehouseData_{timestamp}.db")
        
        # 备份数据库文件
        shutil.copyfile(db_file, backup_file)
        print(f"数据库已备份到: {backup_file}")
        
        # 删除原数据库文件
        os.remove(db_file)
        print("原数据库文件已删除")
        
        # 如果数据库目录为空，则删除目录
        if not os.listdir(db_dir):
            os.rmdir(db_dir)
            print("数据库目录已删除")
        
        return True
    except Exception as e:
        print(f"备份过程中出错: {str(e)}")
        return False

if __name__ == "__main__":
    if backup_database_on_uninstall():
        sys.exit(0)
    else:
        sys.exit(1)