import os, time
import sys
import pyautogui as pag
from PyQt5.QtCore import QSize, QPoint, Qt
from PyQt5.QtGui import QMouseEvent, QMovie, QCursor
from PyQt5.QtWidgets import QWidget, QMessageBox, QApplication, QMenu, qApp
from qtpy import QtWidgets, QtCore
from threading import Thread

class Main(QWidget):
    _startPos = None
    _endPos = None
    _isTracking = False
    all_bytes = 0
    about = "Show mouse coord in pyautogui"

    def __init__(self):
        super().__init__()
        self._initUI()

    def _initUI(self):
        self.setFixedSize(QSize(259, 270))
        self.setGeometry(1600,50,259,270)
        self.setWindowFlags(Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | Qt.Tool)

        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)  # 设置窗口背景透明

        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(0, 0, 259, 111))
        self.label.setMinimumSize(QtCore.QSize(259, 111))
        self.label.setBaseSize(QtCore.QSize(259, 111))
        self.label.setStyleSheet('font: 75 20pt "Adobe Arabic";color:rgb(255,0,0)')
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        # self.label.setText(self.getPos())

        self.timer = QtCore.QTimer(self)
        self.timer.start(1)
        self.timer.timeout.connect(self.start)

        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.show()

    def mouseMoveEvent(self, e: QMouseEvent):  # 重写移动事件
        self._endPos = e.pos() - self._startPos
        self.move(self.pos() + self._endPos)

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._isTracking = True
            self._startPos = QPoint(e.x(), e.y())

        if e.button() == Qt.RightButton:
            menu = QMenu(self)
            quitAction = menu.addAction("Quit")
            action = menu.exec_(self.mapToGlobal(e.pos()))
            if action == quitAction:
                qApp.quit()

    def mouseReleaseEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._isTracking = False
            self._startPos = None
            self._endPos = None
        if e.button() == Qt.RightButton:
            self._isTracking = False
            self._startPos = None
            self._endPos = None

    def start(self):
        Thread(target=self.getPos, daemon=True).start()

    def getPos(self):
        x,y = pag.position() # 返回鼠标的坐标
        posStr="Position:"+str(x).rjust(4)+','+str(y).rjust(4)
        self.label.setText(posStr)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Main()
    sys.exit(app.exec_())

# try:
#     while True:
#             x,y = pag.position() #返回鼠标的坐标
#             posStr="Position:"+str(x).rjust(4)+','+str(y).rjust(4)
#             print(posStr)#打印坐标
#             time.sleep(0.2)
#             os.system('cls')#清楚屏幕
# except  KeyboardInterrupt:
#     print('end....')
