import sys
import os
import sqlite3
import time
from datetime import datetime
import win32api
import win32con
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QPushButton, QLabel, QMessageBox,
                             QListWidget, QListWidgetItem, QAbstractItemView,
                             QSystemTrayIcon, QMenu, QAction, QStackedWidget, QGraphicsDropShadowEffect)
from PyQt5.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QIcon, QColor

# --- 核心配置与资源处理 ---
APP_NAME = "MorningMemo"
ICON_NAME = "112.ico"


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), relative_path)


def get_db_path():
    path = os.path.join(os.path.expanduser("~"), "Documents", "MyMemoApp")
    if not os.path.exists(path): os.makedirs(path)
    return os.path.join(path, "memo_data.db")


# --- 数据库管理 ---
class DBManager:
    def __init__(self):
        self.conn = sqlite3.connect(get_db_path())
        self.cursor = self.conn.cursor()
        self.cursor.execute(
            'CREATE TABLE IF NOT EXISTS memos (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, content TEXT)')
        self.conn.commit()

    def add_memo(self, content):
        self.cursor.execute("INSERT INTO memos (date, content) VALUES (?, ?)",
                            (datetime.now().strftime("%Y-%m-%d %H:%M"), content))
        self.conn.commit()

    def get_all(self):
        self.cursor.execute("SELECT * FROM memos ORDER BY id DESC")
        return self.cursor.fetchall()

    def get_latest(self):
        self.cursor.execute("SELECT content FROM memos ORDER BY id DESC LIMIT 1")
        res = self.cursor.fetchone()
        return res[0] if res else "开启新的一天吧！"

    def delete(self, m_id):
        self.cursor.execute("DELETE FROM memos WHERE id = ?", (m_id,))
        self.conn.commit()


# --- 样式精致的列表项 ---
class HistoryItem(QWidget):
    def __init__(self, m_id, date, content, parent, db):
        super().__init__()
        self.m_id, self.db, self.parent_list = m_id, db, parent
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)

        info_layout = QVBoxLayout()
        date_lbl = QLabel(date)
        date_lbl.setStyleSheet("color: #999; font-size: 11px;")
        cont_lbl = QLabel(content.replace('\n', ' ')[:25] + ("..." if len(content) > 25 else ""))
        cont_lbl.setStyleSheet("color: #333; font-size: 13px; font-weight: 500;")
        info_layout.addWidget(date_lbl)
        info_layout.addWidget(cont_lbl)

        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setStyleSheet("""
            QPushButton { border-radius: 12px; background: #fee; color: #f55; border: none; font-weight: bold; }
            QPushButton:hover { background: #f55; color: white; }
        """)
        del_btn.clicked.connect(self.do_delete)

        layout.addLayout(info_layout)
        layout.addStretch()
        layout.addWidget(del_btn)

    def do_delete(self):
        if QMessageBox.question(self, "删除", "确定删除这条记录吗？") == QMessageBox.Yes:
            self.db.delete(self.m_id)
            for i in range(self.parent_list.count()):
                item = self.parent_list.item(i)
                if self.parent_list.itemWidget(item) == self:
                    self.parent_list.takeItem(i);
                    break


# --- 主窗口：极简卡片 UI ---
class ModernUI(QWidget):
    def __init__(self):
        super().__init__()
        self.db = DBManager()
        self.icon_obj = QIcon(get_resource_path(ICON_NAME))
        self.setup_ui()
        self.init_tray()
        self.auto_run_setup()

    def setup_ui(self):
        self.setWindowTitle("Morning Memo")
        self.setWindowIcon(self.icon_obj)
        self.setFixedSize(420, 580)
        self.setStyleSheet("QWidget { background-color: #f4f6f9; font-family: 'Segoe UI', 'Microsoft YaHei'; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(25, 30, 25, 30)
        main_layout.setSpacing(20)

        # 头部标题区
        header = QLabel("✨ Hello, Today")
        header.setStyleSheet("font-size: 24px; font-weight: 800; color: #1e293b; background: transparent;")
        main_layout.addWidget(header)

        # 中央卡片容器
        self.container = QStackedWidget()
        self.container.setStyleSheet("background: white; border-radius: 20px;")

        # 阴影效果
        shadow = QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=10, color=QColor(0, 0, 0, 30))
        self.container.setGraphicsEffect(shadow)

        # 页面1：今日输入
        self.page_write = QWidget()
        pw_layout = QVBoxLayout(self.page_write)
        pw_layout.setContentsMargins(20, 20, 20, 20)

        self.edit = QTextEdit()
        self.edit.setPlaceholderText("在此输入明天的计划...")
        self.edit.setText(self.db.get_latest())
        self.edit.setStyleSheet("""
            QTextEdit { border: none; font-size: 15px; color: #334155; line-height: 1.5; }
        """)
        pw_layout.addWidget(self.edit)

        save_btn = QPushButton("保存并同步明天")
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setStyleSheet("""
            QPushButton { background: #0f172a; color: white; border-radius: 12px; height: 45px; font-weight: bold; font-size: 14px; }
            QPushButton:hover { background: #334155; }
        """)
        save_btn.clicked.connect(self.save_data)
        pw_layout.addWidget(save_btn)

        # 页面2：历史列表
        self.page_list = QWidget()
        pl_layout = QVBoxLayout(self.page_list)
        pl_layout.setContentsMargins(10, 10, 10, 10)
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(
            "QListWidget { border: none; } QScrollBar:vertical { width: 5px; background: transparent; }")
        pl_layout.addWidget(self.list_widget)

        self.container.addWidget(self.page_write)
        self.container.addWidget(self.page_list)
        main_layout.addWidget(self.container)

        # 底部导航切换
        nav_layout = QHBoxLayout()
        self.btn_goto_write = QPushButton("📝 计划")
        self.btn_goto_hist = QPushButton("📜 历史")
        for b in [self.btn_goto_write, self.btn_goto_hist]:
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet("""
                QPushButton { background: transparent; color: #64748b; font-weight: bold; border: none; padding: 5px; }
                QPushButton:checked { color: #0f172a; border-bottom: 2px solid #0f172a; }
            """)
        self.btn_goto_write.setChecked(True)
        self.btn_goto_write.clicked.connect(lambda: self.switch_page(0))
        self.btn_goto_hist.clicked.connect(lambda: self.switch_page(1))

        nav_layout.addStretch()
        nav_layout.addWidget(self.btn_goto_write)
        nav_layout.addSpacing(40)
        nav_layout.addWidget(self.btn_goto_hist)
        nav_layout.addStretch()
        main_layout.addLayout(nav_layout)

    def switch_page(self, index):
        self.container.setCurrentIndex(index)
        self.btn_goto_write.setChecked(index == 0)
        self.btn_goto_hist.setChecked(index == 1)
        if index == 1: self.refresh_history()

    def refresh_history(self):
        self.list_widget.clear()
        for m_id, date, content in self.db.get_all():
            item = QListWidgetItem(self.list_widget)
            item.setSizeHint(QSize(0, 60))
            self.list_widget.setItemWidget(item, HistoryItem(m_id, date, content, self.list_widget, self.db))

    def save_data(self):
        txt = self.edit.toPlainText().strip()
        if txt:
            self.db.add_memo(txt)
            self.hide()
            self.tray.showMessage("已保存", "明天开机将再次提醒您。", QSystemTrayIcon.Information, 2000)
        else:
            QMessageBox.warning(self, "提示", "内容不能为空哦")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.icon_obj)
        menu = QMenu()
        a_show = QAction("显示窗口", self)
        a_show.triggered.connect(self.showNormal)
        a_quit = QAction("退出程序", self)
        a_quit.triggered.connect(QApplication.quit)
        menu.addAction(a_show);
        menu.addAction(a_quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.DoubleClick else None)
        self.tray.show()

    def auto_run_setup(self):
        # 依旧使用任务计划程序（Task Scheduler）确保最高权限自启
        raw_path = os.path.realpath(sys.argv[0])
        if not raw_path.endswith(".exe"): return
        task_name = "MorningMemoAutostart"
        flag = os.path.join(os.path.dirname(get_db_path()), ".task_ok")
        if not os.path.exists(flag):
            try:
                cmd = f'schtasks /create /f /tn "{task_name}" /tr "\\"{raw_path}\\"" /sc onlogon /rl highest /it'
                os.system(cmd)
                with open(flag, "w") as f:
                    f.write("ok")
            except:
                pass

    def closeEvent(self, event):
        if self.tray.isVisible():
            self.hide();
            event.ignore()
        else:
            event.accept()


if __name__ == '__main__':
    # 模拟开机延迟，防止环境未就绪
    if len(sys.argv) > 1: time.sleep(5)
    app = QApplication(sys.argv)
    ex = ModernUI()
    ex.show()
    sys.exit(app.exec_())