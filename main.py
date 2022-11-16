import sys
from PyQt5.QtCore import QLockFile, QDir, Qt, QPropertyAnimation, QEvent
from window import *
import win32gui
from PyQt5.QtWidgets import QMainWindow, QApplication, QSystemTrayIcon, QMenu, QAction, QMessageBox, QStyle
from PyQt5.QtGui import QIcon, QColor
import pywinctl

apps = {}


def _access_error(func):
    def decorator(*args, **kwargs):
        try:
            self = args[0]
            func(*args, **kwargs)
        except win32gui.error as exc:
            if exc.args[0] == 5:
                QMessageBox.information(self, 'Notification', 'Access Denied\nRun as administrator')
        except Exception as exc:
            print('list_clicked error', exc)
        else:
            self.hide()

    return decorator


class ToTop:
    def __init__(self, fg):
        self.fg = fg

    def top(self):
        win32gui.SetWindowPos(self.fg, -1, 0, 0, 0, 0, 2 | 1)

    def cancel(self):
        win32gui.SetWindowPos(self.fg, -2, 0, 0, 0, 0, 2 | 1)


class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setWindowIcon(QtGui.QIcon('plum.ico'))
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.ui = Ui_MainWindow()
        self.setWindowTitle('PlumTopper')
        self.ui.setupUi(self)
        self.ui.listWidget.itemDoubleClicked.connect(self.list_clicked)
        self.ui.listWidget.itemChanged.connect(self.list_checked)
        self.tray_menu()
        self.tray.activated.connect(self.tray_activated)
        self.move(QApplication.desktop().availableGeometry().width() - 500,
                  QApplication.desktop().availableGeometry().height() - 600)

        pixmapi = QStyle.SP_TitleBarUnshadeButton
        icon = self.style().standardIcon(pixmapi)
        self.ui.close_btn.setIcon(icon)
        self.ui.close_btn.clicked.connect(lambda: self.hide())

        self.animation = QPropertyAnimation(self, b'windowOpacity')
        self.animation.setDuration(200)

        self.ui.listWidget.installEventFilter(self)

    def tray_activated(self, reason):
        if reason == 3:
            if self.isHidden():
                self.showNormal()
            else:
                self.activateWindow()

    def tray_menu(self):
        self.menu = QMenu()
        self.tray = QSystemTrayIcon(QIcon("tray.svg"))
        self.action_exit = QAction("Exit")
        self.action_exit.triggered.connect(app.exit)
        self.menu.addAction(self.action_exit)
        self.tray.setToolTip("PlumTopper")
        self.tray.show()
        self.tray.setContextMenu(self.menu)

    def event(self, event):
        if event.type() == QtCore.QEvent.WindowActivate:
            self.create_list()
        return QtWidgets.QWidget.event(self, event)

    @_access_error
    def sizer(self, clicked_window, location):
        handle = clicked_window.value
        max_w = QApplication.desktop().availableGeometry().width()+28
        max_h = QApplication.desktop().availableGeometry().height()+16
        win32gui.ShowWindow(handle, 1)
        win32gui.SetForegroundWindow(handle)
        if location == 'left_top':
            win32gui.MoveWindow(handle, -7, 0, max_w // 2, max_h // 2, True)
        if location == 'right_top':
            win32gui.MoveWindow(handle, (max_w // 2) - 22, 0, max_w // 2, max_h // 2, True)
        if location == 'left_bottom':
            win32gui.MoveWindow(handle, -7, max_h // 2 - 10, max_w // 2, max_h // 2, True)
        if location == 'right_bottom':
            win32gui.MoveWindow(handle, (max_w // 2) - 22, max_h // 2 - 10, max_w // 2, max_h // 2, True)
        self.create_list()

    def action_clicked(self):
        action = self.sender()
        self.sizer(action.data(), action.text())

    def eventFilter(self, source, event):
        if event.type() == QEvent.ContextMenu and source is self.ui.listWidget:
            item = source.itemAt(event.pos())
            if item != None:
                self.con_menu = QMenu()
                sides = 'left_top', 'right_top', 'left_bottom', 'right_bottom'
                for side in sides:
                    self.menu_action = QAction(side, self)
                    self.menu_action.setObjectName(side)
                    self.menu_action.setData(item)
                    self.menu_action.setIcon(QIcon(f'{side}.png'))
                    self.menu_action.triggered.connect(self.action_clicked)
                    self.con_menu.addAction(self.menu_action)
                self.con_menu.exec_(event.globalPos())
            return True
        return super().eventFilter(source, event)

    def list_clicked(self, clicked_window):
        if clicked_window.checkState() == 2:
            clicked_window.setCheckState(0)
        else:
            clicked_window.setCheckState(2)

    @_access_error
    def list_checked(self, clicked_window):
        handle = clicked_window.value
        to_top = ToTop(handle)
        if clicked_window.checkState() == 2:
            to_top.top()
            clicked_window.setBackground(QColor('#7fc97f'))
        else:
            to_top.cancel()
            clicked_window.setBackground(QColor('white'))

    # animation
    def showNormal(self):
        super(MainWindow, self).showNormal()
        self.animation.stop()
        self.animation.setStartValue(0)
        self.animation.setEndValue(1)
        self.animation.start()

    def create_list(self):
        self.ui.listWidget.clear()
        self.get_apps()
        for handle, title in apps.items():
            it = QtWidgets.QListWidgetItem(title)
            it.value = handle
            style = win32gui.GetWindowLong(handle, -20)
            # check if window is always on top
            if style & 8 == 8:
                it.setCheckState(2)
                it.setBackground(QColor('#7fc97f'))
            else:
                it.setCheckState(0)
            self.ui.listWidget.addItem(it)

    def get_apps(self):
        apps.clear()
        all_windows = pywinctl.getAllWindows()
        for window in all_windows:
            title = window.title
            handle = window.getHandle()
            self_title = self.windowTitle()
            if (title != self_title) and len(title) > 0:
                apps[handle] = title


if __name__ == '__main__':
    lockfile = QLockFile(QDir.temp().absoluteFilePath('plumtopper.lock'))
    try:
        app = QApplication(sys.argv)
        if lockfile.tryLock():
            app.setQuitOnLastWindowClosed(False)
            dlgMain = MainWindow()
            dlgMain.hide()
            sys.exit(app.exec_())
        else:
            error = QMessageBox()
            error.setWindowTitle("Error")
            error.setIcon(QMessageBox.Warning)
            error.setText("The application is already running!")
            error.setStandardButtons(QMessageBox.Ok)
            error.exec()
    finally:
        lockfile.unlock()
