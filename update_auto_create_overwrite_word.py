import os
import time
import win32com.client as win32
import pythoncom
import psutil
from datetime import datetime
import sys

class RealTimeWordSync:
    def __init__(self):
        # 获取可执行文件所在目录
        if getattr(sys, 'frozen', False):
            # 打包后：获取可执行文件所在目录
            self.script_dir = os.path.dirname(sys.executable)
        else:
            # 开发阶段：获取脚本所在目录
            self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 路径配置
        self.backup_folder = os.path.join(self.script_dir, "SyncBackup")
        self.live_doc_path = os.path.join(self.script_dir, "WorkingDocument.docx")  # 主编辑文件
        self.backup_path = os.path.join(self.backup_folder, "RealTimeBackup.docx")  # 实时备份
        
        # 初始化环境
        self._prepare_folders()
        self.word_app = None
        self.doc = None
        self._init_word()

    def _prepare_folders(self):
        """创建必要目录"""
        os.makedirs(self.backup_folder, exist_ok=True)

    def _force_kill_word(self):
        """彻底清理Word进程"""
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'WINWORD.EXE':
                try:
                    proc.kill()
                except psutil.AccessDenied:
                    pass

    def _init_word(self):
        """初始化Word及主文档"""
        self._force_kill_word()
        try:
            pythoncom.CoInitialize()
            self.word_app = win32.DispatchEx("Word.Application")
            self.word_app.Visible = True
            
            # 优先打开已有主文档，不存在则创建新文档
            if os.path.exists(self.live_doc_path):
                self.doc = self.word_app.Documents.Open(self.live_doc_path)
                print(f"已载入现有主文档: {self.live_doc_path}")
            else:
                self.doc = self.word_app.Documents.Add()
                self.doc.SaveAs(self.live_doc_path)  # 首次创建主文档
                print(f"新建主文档: {self.live_doc_path}")
            return True
        except Exception as e:
            print(f"Word初始化失败: {e}")
            return False

    def _sync_backup(self):
        """执行覆盖式备份"""
        try:
            # 先保存主文档的修改
            self.doc.Save()
            
            # 另存为备份文件（覆盖模式）
            self.doc.SaveAs(self.backup_path)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] 实时备份更新 -> {self.backup_path}")
        except Exception as e:
            print(f"备份失败: {str(e)}")
            self._force_kill_word()
            time.sleep(1)
            self._init_word()

    def start_sync(self):
        """启动同步服务"""
        print(f"实时备份服务运行中...\n主文档: {self.live_doc_path}\n备份文件: {self.backup_path}")
        time.sleep(30)  # 初始等待
        
        while True:
            try:
                self._sync_backup()
                time.sleep(30)
            except KeyboardInterrupt:
                print("\n安全退出...")
                self.doc.Close(SaveChanges=True)
                self.word_app.Quit()
                break

if __name__ == "__main__":
    syncer = RealTimeWordSync()
    syncer.start_sync()
