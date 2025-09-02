# 文件名: database_migration.py
# 功能: 数据库迁移脚本，用于更新试剂信息表的唯一约束

import os
import sys
import sqlite3
import tempfile
import shutil

def get_database_path():
    """获取数据库文件路径"""
    if getattr(sys, 'frozen', False):
        # 打包后的可执行文件 - 使用AppData目录
        appdata_path = os.getenv('APPDATA')
        if appdata_path:
            db_dir = os.path.join(appdata_path, "ReagentManagementSystem")
            db_file = os.path.join(db_dir, "ReagentWarehouseData.db")
            return db_file
        else:
            # 备用方案：使用可执行文件目录
            db_dir = os.path.dirname(sys.executable)
            db_file = os.path.join(db_dir, "ReagentWarehouseData.db")
            return db_file
    else:
        # 开发环境
        db_dir = os.path.dirname(os.path.abspath(__file__))
        db_file = os.path.join(db_dir, "ReagentWarehouseData.db")
        return db_file

def migrate_database():
    """执行数据库迁移"""
    db_file = get_database_path()
    
    # 检查数据库文件是否存在
    if not os.path.exists(db_file):
        print(f"数据库文件不存在: {db_file}")
        return False
    
    print(f"开始迁移数据库: {db_file}")
    
    try:
        # 创建数据库备份
        backup_dir = os.path.join(os.path.dirname(db_file), "Backups")
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(backup_dir, f"ReagentWarehouseData_backup_{timestamp}.db")
        shutil.copyfile(db_file, backup_file)
        print(f"已创建数据库备份: {backup_file}")
        
        # 连接数据库
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 检查当前表结构
        cursor.execute("PRAGMA table_info(reagents)")
        columns = cursor.fetchall()
        print("当前表结构:")
        for col in columns:
            print(f"  列: {col[1]}, 类型: {col[2]}")
        
        # 检查是否存在旧的唯一约束（仅batch）
        cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='reagents'")
        table_sql = cursor.fetchone()[0]
        print(f"当前表定义: {table_sql}")
        
        # 检查是否已经存在复合唯一约束
        if "UNIQUE (name, batch)" in table_sql:
            print("数据库已是最新结构，无需迁移")
            conn.close()
            return True
        
        # 创建临时表（新结构）
        print("创建新结构临时表...")
        cursor.execute('''
            CREATE TABLE reagents_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                batch TEXT NOT NULL,
                production_date TEXT NOT NULL,
                expiry_date TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                gtin TEXT,
                qrcode TEXT,
                UNIQUE (name, batch)
            )
        ''')
        
        # 复制数据到临时表
        print("迁移数据到新表...")
        cursor.execute('''
            INSERT OR IGNORE INTO reagents_new 
            (id, name, batch, production_date, expiry_date, quantity, gtin, qrcode)
            SELECT id, name, batch, production_date, expiry_date, quantity, gtin, qrcode
            FROM reagents
            ORDER BY id
        ''')
        
        # 检查是否有重复数据被忽略
        cursor.execute('SELECT COUNT(*) FROM reagents')
        old_count = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM reagents_new')
        new_count = cursor.fetchone()[0]
        
        if old_count != new_count:
            print(f"警告: 有 {old_count - new_count} 条重复数据被忽略")
        
        # 删除旧表
        print("删除旧表...")
        cursor.execute('DROP TABLE reagents')
        
        # 重命名新表
        print("重命名新表...")
        cursor.execute('ALTER TABLE reagents_new RENAME TO reagents')
        
        # 提交更改
        conn.commit()
        conn.close()
        
        print("数据库迁移成功完成!")
        return True
        
    except Exception as e:
        print(f"数据库迁移失败: {str(e)}")
        # 尝试恢复备份
        try:
            if 'backup_file' in locals() and os.path.exists(backup_file):
                shutil.copyfile(backup_file, db_file)
                print("已从备份恢复数据库")
        except Exception as restore_error:
            print(f"恢复备份失败: {str(restore_error)}")
        return False

if __name__ == "__main__":
    from datetime import datetime
    success = migrate_database()
    if success:
        print("迁移完成")
        sys.exit(0)
    else:
        print("迁移失败")
        sys.exit(1)