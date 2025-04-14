# --*coding=utf-8*--
# 这个程序是用于便利地将规整的PDF文件转化为Excel\word\图片等 
# 许钦滔 2024.3.13/3.18/5.10/5.15/5.16
import os
import shutil
import sys

import fitz
import pdfplumber
import pytesseract
from docx import Document
from openpyxl import Workbook
from pdf2docx import Converter as docxConverter
from PIL import Image
from pptx import Presentation
from PyPDF2 import PdfReader, PdfWriter
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QCheckBox, QFileDialog, QMessageBox

file_load = os.getcwd()



class NewQLineEdit(QtWidgets.QLineEdit):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)  # 删除没有影响，目前不确定（因为True和False测试结果一样）
        self.setDragEnabled(True)  # 删除没有影响，（因为True和False测试结果一样）

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():  # 当文件拖入此区域时为True
            event.accept()  # 接受拖入文件
        else:
            event.ignore()  # 忽略拖入文件

    def dropEvent(self, event):  # 本方法为父类方法，本方法中的event为鼠标放事件对象
        urls = [u for u in event.mimeData().urls()]  # 范围文件路径的Qt内部类型对象列表，由于支持多个文件同时拖入所以使用列表存放
        for url in urls:
            self.setText(url.path()[1:])  # 将Qt内部类型转换为字符串类型

class NewQLineEdit(QtWidgets.QLineEdit):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)  # 删除没有影响，目前不确定（因为True和False测试结果一样）
        self.setDragEnabled(True)  # 删除没有影响，（因为True和False测试结果一样）

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():  # 当文件拖入此区域时为True
            event.accept()  # 接受拖入文件
        else:
            event.ignore()  # 忽略拖入文件

    def dropEvent(self, event):  # 本方法为父类方法，本方法中的event为鼠标放事件对象
        urls = [u for u in event.mimeData().urls()]  # 范围文件路径的Qt内部类型对象列表，由于支持多个文件同时拖入所以使用列表存放
        for url in urls:
            self.setText(url.path()[1:])  # 将Qt内部类型转换为字符串类型


class Ui_MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self._startPos = None
        self._endPos = None
        self._tracking = False
        self.resize(733, 500)
        self.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.setAcceptDrops(True)
        self.checked = 0
        

        # 设置背景图
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(0, 0, 681, 251))
        self.label.setText("")
        # 还需要获取绝对路径？
        file_load = os.getcwd()
        self.label.setPixmap(QtGui.QPixmap(file_load + '/image/background2.png'))
        self.label.setScaledContents(True)
        self.label.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)

        # 设置加载动画
        self.loading_movie = QtGui.QMovie("loading.gif")  # 加载动画
        self.loading_label = QtWidgets.QLabel(self)
        self.loading_label.setMovie(self.loading_movie)
        self.loading_label.hide() 
        self.loading_movie.start()
        self.loading_movie.setScaledSize(QtCore.QSize(100, 100))
        self.loading_label.setGeometry(300, 180, 100, 100)
        self.loading_label.raise_()
    

        # label = QLabel("我是一个窗体，但我没有边框！！！", self)
        # label.move(6, 6)

        # 设置输入框
        self.lineEdit = NewQLineEdit(self)  # 此处更改
        self.lineEdit.setGeometry(QtCore.QRect(255, 38, 355, 50))
        self.lineEdit.setAcceptDrops(True)
        self.lineEdit.setStyleSheet("font: 12pt \"Arial\";")
        self.lineEdit.setText("(等待中)>>")
        self.lineEdit.setObjectName("lineEdit")

        # 设置标签
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(35, 40, 211, 41))
        self.label_3 = QtWidgets.QLabel(self)
        self.label_3.setGeometry(QtCore.QRect(280, 200, 192, 41))
        self.label_5 = QtWidgets.QLabel(self)
        self.label_5.setGeometry(QtCore.QRect(125, 170, 250, 41))
        self.label_4 = QtWidgets.QLabel(self)
        self.label_4.setGeometry(QtCore.QRect(440, 140, 651, 180))
        self.label_4.setStyleSheet("color: gray")

        # 设置按钮
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(350, 113, 131, 41))
        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        self.pushButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton.setStyleSheet("QPushButton {\n"
                                      "    background-color: #ffffff;\n"
                                      "    border: 1px solid #dcdfe6;\n"
                                      "    padding: 10px;\n"
                                      "    border-radius: 5px;\n"
                                      "}\n"
                                      "\n"
                                      "QPushButton:hover {\n"
                                      "    background-color: #ecf5ff;\n"
                                      "    color: #409eff;\n"
                                      "}\n"
                                      "\n"
                                      "QPushButton:pressed, QPushButton:checked {\n"
                                      "    border: 1px solid #3a8ee6;\n"
                                      "    color: #409eff;\n"
                                      "}\n"
                                      "\n"
                                      "#button3 {\n"
                                      "    border-radius: 20px;\n"
                                      "}")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setText("转化为Excel")
        self.pushButton.clicked.connect(self.pdfToExcel)

        self.pushButton_2 = QtWidgets.QPushButton(self)
        self.pushButton_2.setGeometry(QtCore.QRect(200, 103, 131, 41))
        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(12)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_2.setStyleSheet("QPushButton {\n"
                                        "    background-color: #ffffff;\n"
                                        "    border: 1px solid #dcdfe6;\n"
                                        "    padding: 10px;\n"
                                        "    border-radius: 5px;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:hover {\n"
                                        "    background-color: #ecf5ff;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:pressed, QPushButton:checked {\n"
                                        "    border: 1px solid #3a8ee6;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "#button3 {\n"
                                        "    border-radius: 20px;\n"
                                        "}")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.pdfToWord)
        self.pushButton_2.setText("转化为Word")

        # self.pushButton_2.clicked.connect(self.close)
        self.pushButton_3 = QtWidgets.QPushButton(self)
        self.pushButton_3.setGeometry(QtCore.QRect(50, 113, 131, 41))
        self.pushButton_3.setFont(font)
        self.pushButton_3.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_3.setStyleSheet("QPushButton {\n"
                                        "    background-color: #ffffff;\n"
                                        "    border: 1px solid #dcdfe6;\n"
                                        "    padding: 10px;\n"
                                        "    border-radius: 5px;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:hover {\n"
                                        "    background-color: #ecf5ff;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:pressed, QPushButton:checked {\n"
                                        "    border: 1px solid #3a8ee6;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "#button3 {\n"
                                        "    border-radius: 20px;\n"
                                        "}")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.setText("转化为图片")
        self.pushButton_3.clicked.connect(self.pdfToImage)  # 转化为图片

        self.pushButton_5 = QtWidgets.QPushButton(self)
        self.pushButton_5.setGeometry(QtCore.QRect(500, 113, 131, 41))
        self.pushButton_5.setFont(font)
        self.pushButton_5.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_5.setStyleSheet("QPushButton {\n"
                                        "    background-color: #ffffff;\n"
                                        "    border: 1px solid #dcdfe6;\n"
                                        "    padding: 10px;\n"
                                        "    border-radius: 5px;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:hover {\n"
                                        "    background-color: #ecf5ff;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:pressed, QPushButton:checked {\n"
                                        "    border: 1px solid #3a8ee6;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "#button3 {\n"
                                        "    border-radius: 20px;\n"
                                        "}")
        self.pushButton_5.setObjectName("pushButton_3")
        self.pushButton_5.setText("转化为PPT")
        self.pushButton_5.clicked.connect(self.pdfToPPT)  # 转化为ppt

        self.pushButton_4 = QtWidgets.QPushButton(self)
        self.pushButton_4.setGeometry(QtCore.QRect(370, 170, 185, 41))
        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(13)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton_4.setStyleSheet("QPushButton {\n"
                                        "    background-color: #ffffff;\n"
                                        "    border: 1px solid #dcdfe6;\n"
                                        "    padding: 10px;\n"
                                        "    border-radius: 5px;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:hover {\n"
                                        "    background-color: #ecf5ff;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "QPushButton:pressed, QPushButton:checked {\n"
                                        "    border: 1px solid #3a8ee6;\n"
                                        "    color: #409eff;\n"
                                        "}\n"
                                        "\n"
                                        "#button3 {\n"
                                        "    border-radius: 20px;\n"
                                        "}")
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.pdfStrip)
        self.pushButton_4.setText("拆分出一个新的PDF")

        # 设置进度条
        # self.progressBar = QtWidgets.QProgressBar(self)
        # # self.progressBar.setGeometry(QtCore.QRect(40, 400, 632, 10))
        # self.progressBar.setGeometry(QtCore.QRect(260, 223, 192, 10))
        # self.progressBar.setProperty("value", 0)
        # self.progressBar.setObjectName("progressBar")

        # 设置勾选框
        self.checkBox = QCheckBox("使用OCR请打钩\n OCR速度更慢", self)
        self.checkBox.move(213,143)
        self.checkBox.setChecked(False)
        self.checkBox.stateChanged.connect(self.checkBoxState)

        # 设置小输入框
        self.textbox_start = QLineEdit(self)
        self.textbox_start.setGeometry(210, 175, 30, 30)
        self.textbox_start.setText("1")

        self.textbox_end = QLineEdit(self)
        self.textbox_end.setGeometry(305, 175, 30, 30)
        self.textbox_end.setText("1")

        # 设置字体和内容
        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setText("请拖一个PDF文件到此：")
        self.label_2.setObjectName("label_2")
        self.label_3.setFont(font)
        self.label_3.setText("")     # 隐藏
        self.label_3.setObjectName("label_3")
        self.label_5.setFont(font)
        self.label_5.setText("我想把第       页到第       页")     # 隐藏
        self.label_5.setObjectName("label_3")
        
        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(11)
        font.setBold(False)
        font.setWeight(50)

        font = QtGui.QFont()
        font.setFamily("Microsoft Yahei")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(50)
        self.label_4.setFont(font)
        self.label_4.setText("PDF工具箱 (温岭第一检察部Win测试版)")
        self.label_4.setObjectName("label_4")

        # 添加红色圆圈(关闭按钮)
        self.close_button = QtWidgets.QPushButton(self)
        self.close_button.setGeometry(QtCore.QRect(643, 7, 27, 27))
        self.close_button.setStyleSheet("background-color:DeepPink;border-radius:13px;")
        self.close_button.clicked.connect(self.close)

        # 窗体置顶(窗体置顶，仅仅为了方便测试)，去边框
        # self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setWindowFlags(Qt.FramelessWindowHint)
        # 窗体透明，控件不透明
        self.setAttribute(Qt.WA_TranslucentBackground)
        # 设置窗口透明度
        self.setWindowOpacity(0.91)

    def mouseMoveEvent(self, e: QMouseEvent):  # 重写移动事件
        if self._tracking:
            self._endPos = e.pos() - self._startPos
            self.move(self.pos() + self._endPos)

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._startPos = QPoint(e.x(), e.y())
            self._tracking = True

    def mouseReleaseEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._tracking = False
            self._startPos = None
            self._endPos = None

    def checkBoxState(self):
        if self.checkBox.isChecked():
            self.checked = 1
            print(self.checked)
            print("已勾选")
        else:
            self.checked = 0
            print(self.checked)
            print("已取消勾选")

    # def close_all_buttons(self):
    #     self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False)

    def pdfToExcel(self):
        self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False), self.pushButton_5.setEnabled(False) 
        # self.label_3.setText("开始读取数据，请稍等")
        msgBox = QMessageBox()
        msgBox.setWindowTitle("请仔细阅读说明")
        msgBox.setText("点击确定后开始转化！请耐心等待，中途不要有其他操作，以免卡死。\n     完成后会有弹窗提示！100页内一般在半分钟内可以完成\n   只能转化规整的PDF，如电子账单等，拍摄的可能失效！")
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()
        print('调试信息：进入pdfToExcel函数')

        self.loading_label.show()   # 显示加载动画

        self.pdf_file_path = self.lineEdit.text()  # 将lineEdit中的值设置为默认 Windows
        # if not hasattr(self, 'zip_file_path'):
        # print('调试信息：进入if not hasattr(selfzip_file_path')
        pdf_file_path = self.lineEdit.text()  # Windows
        # pdf_file_path = '/' + self.lineEdit.text() # Mac
        print(pdf_file_path)
        if not os.path.exists(pdf_file_path):
            self.lineEdit.setText("错误：文件或被移动，请重新拖入！")
            self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True), self.pushButton_5.setEnabled(True) 
            return
        # self.progressBar.setValue(0) 
        # 获取文件名
        pdf_file_name = os.path.basename(pdf_file_path)
        self.pdf_file_name = pdf_file_name
        file_name = pdf_file_name.split(".")[-2]
        file_path = pdf_file_path.rstrip(pdf_file_name)
        print("pdf_file_name",file_name)

        self.worker = WorkerThread_Excel(self.lineEdit.text())
        # 信号连接到槽函数
        self.worker.finished.connect(self.onFinished)
        self.worker.error.connect(self.onError)
        # 启动线程
        self.worker.start()
        
        #     # 完成提示
        #     msgBox = QMessageBox()
        #     msgBox.setWindowTitle("成功！")
        #     msgBox.setText("Excel文件已成功生成至原来目录 ： " + "\n" + endfile )
        #     msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        #     msgBox.exec_()
        # except:
        #     msgBox = QMessageBox()
        #     msgBox.setWindowTitle("生成错误！")
        #     msgBox.setText("Excel文件因不明原因生成失败！请检查是否拖入了错误的PDF文件或联系开发人员：86086182。")
        #     msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        #     msgBox.exec_()
        
        # self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True)

    def pdfToWord(self):
        self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False), self.pushButton_5.setEnabled(False) 
        msgBox = QMessageBox()
        msgBox.setWindowTitle("请仔细阅读说明")
        msgBox.setText("点击确定后开始转化！请耐心等待，中途不要有其他操作，以免卡死。\n     完成后会有弹窗提示！5页内一般在半分钟内可以完成\n 如未勾选OCR，只能转化非图片的PDF，如案件文书等，拍摄的可能失效！")
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()
        self.pdf_file_path = self.lineEdit.text()  # 将lineEdit中的值设置为默认 Windows
        pdf_file_path = self.lineEdit.text()  # Windows
        print(pdf_file_path)
        if not os.path.exists(pdf_file_path):
            self.lineEdit.setText("错误：文件或被移动，请重新拖入！")
            self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True), self.pushButton_5.setEnabled(True) 
            return
        self.loading_label.show()   # 显示加载动画
        # 获取文件名
        pdf_file_name = os.path.basename(pdf_file_path)
        file_name = pdf_file_name.split(".")[-2]
        print("!!!",self.checked)
        self.worker = WorkerThread_WORD(self.lineEdit.text(),self.checked)
        # 信号连接到槽函数
        self.worker.finished.connect(self.onFinished)
        self.worker.error.connect(self.onError)
        # 启动线程
        self.worker.start()


    def pdfToImage(self):
        self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False), self.pushButton_5.setEnabled(False) 
        msgBox = QMessageBox()
        msgBox.setWindowTitle("请仔细阅读说明")
        msgBox.setText("点击确定后开始转化！请耐心等待，中途不要有其他操作，以免卡死。\n     完成后会有弹窗提示！50页内一般在半分钟内可以完成\n   只能转化非图片的PDF，如案件文书等，拍摄的可能失效！")
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()
        self.pdf_file_path = self.lineEdit.text()  # 将lineEdit中的值设置为默认 Windows
        pdf_file_path = self.lineEdit.text()  # Windows
        print(pdf_file_path)
        if not os.path.exists(pdf_file_path):
            self.lineEdit.setText("错误：文件或被移动，请重新拖入！")
            self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True), self.pushButton_4.setEnabled(True), self.pushButton_5.setEnabled(True) 
            return
        # self.loading_label.show()   # 显示加载动画
        # 获取文件名
        pdf_file_name = os.path.basename(pdf_file_path)
        file_name = pdf_file_name.split(".")[-2]
        file_path = pdf_file_path.rstrip(pdf_file_name)
        img_start = 0
        # img_start = img_start - 1
        img_end = 0
        # try:
        with pdfplumber.open(pdf_file_path) as pdf:
            pages_len = len(pdf.pages)
            img_end = pages_len
            # for i, page in enumerate(pdf.pages[:2]):
            for i, page in enumerate(pdf.pages[img_start:img_end]):
                im = page.to_image(resolution=300)
                im.save(file_path +'{}.png'.format(i + 1))
                print('------分割线，第%d页-------'%(int(i)+1))

        msgBox = QMessageBox()
        msgBox.setWindowTitle("成功！")
        msgBox.setText("图片已成功生成至原来目录" )
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()
        # except:
        #     msgBox = QMessageBox()
        #     msgBox.setWindowTitle("生成错误！")
        #     msgBox.setText("Excel文件因不明原因生成失败！请检查是否拖入了错误的PDF文件或联系开发人员：86086182。")
        #     msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        #     msgBox.exec_()
        self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 

    def pdfStrip(self):
        self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False) ,self.pushButton_5.setEnabled(False) 
        # msgBox = QMessageBox()
        # msgBox.setWindowTitle("请仔细阅读说明")
        # msgBox.setText("点击确定后开始转化！请耐心等待，中途不要有其他操作，以免卡死。\n     完成后会有弹窗提示！一般在半分钟内可以完成\n   只能转化非图片的PDF，如案件文书等，拍摄的可能失效！")
        # msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        # msgBox.exec_()
        self.pdf_file_path = self.lineEdit.text()  # 将lineEdit中的值设置为默认 Windows
        pdf_file_path = self.lineEdit.text()  # Windows
        print(pdf_file_path)
        if not os.path.exists(pdf_file_path):
            self.lineEdit.setText("错误：文件或被移动，请重新拖入！")
            self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 
            return
        # self.loading_label.show()   # 显示加载动画
        # 获取文件名
        pdf_file_name = os.path.basename(pdf_file_path)
        file_name = pdf_file_name.split(".")[-2]
        file_path = pdf_file_path.rstrip(pdf_file_name)
        page_start = int(self.textbox_start.text()) - 1
        # img_start = img_start - 1
        page_end = int(self.textbox_end.text())
        print(page_start,page_end)
        pdf_writer = PdfWriter()

        # if page_start and page_end:
        #     print('进入if')
        try:
            with open(pdf_file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                for i in range(page_start, page_end):
                    pdf_writer.add_page(pdf_reader.pages[i])

                with open(file_path + f'(第{page_start + 1}到{page_end}页)' + file_name + '.pdf', 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)

                # output_file = pdfplumber.open(file_path + "（拆分）" + file_name + ".pdf", 'wb')
                # output_file.save()
                msgBox = QMessageBox()
                msgBox.setWindowTitle("成功！")
                msgBox.setText("拆分的PDF已成功生成至原来目录" )
                msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
                msgBox.exec_()
        except:
            msgBox = QMessageBox()
            msgBox.setWindowTitle("生成错误！")
            msgBox.setText("文件因不明原因生成失败！请检查是否拖入了错误的PDF文件或输入了错误的页码。联系开发人员：86086182。")
            msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
            msgBox.exec_()
        self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 
        # else:
        #     msgBox = QMessageBox()
        #     msgBox.setWindowTitle("输入错误！")
        #     msgBox.setText("文件因不明原因生成失败！请检查是否拖入了错误的PDF文件或输入了错误的页码。可联系开发人员：86086182。")
        #     msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        #     msgBox.exec_()

    # def update_label(self):
    #     self.accept()

    def pdfToPPT(self):
        print('开始处理PPT')
        self.pushButton.setEnabled(False), self.pushButton_2.setEnabled(False), self.pushButton_3.setEnabled(False), self.pushButton_4.setEnabled(False) ,self.pushButton_5.setEnabled(False) 
        self.pdf_file_path = self.lineEdit.text()  # 将lineEdit中的值设置为默认 Windows
        pdf_file_path = self.lineEdit.text()  # Windows
        print(pdf_file_path)
        if not os.path.exists(pdf_file_path):
            self.lineEdit.setText("错误：文件或被移动，请重新拖入！")
            self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 
            return
        # pdf_file = '222.pdf'
        pdf_file = pdf_file_path
        
        # 获取文件名
        pdf_file_name = os.path.basename(pdf_file_path)
        file_name = pdf_file_name.split(".")[-2]
        file_path = pdf_file_path.rstrip(pdf_file_name)

        ppt = Presentation()

        if not os.path.exists("ppt_img_temps"):     # 检查是否已经存在文件夹
            os.mkdir("ppt_img_temps")
            print("成功创建文件夹1")
        else:
            shutil.rmtree("ppt_img_temps")
            os.mkdir("ppt_img_temps")
            print("成功创建文件夹2")
        path_load = os.getcwd() + r'\ppt_img_temps/'
        
        # 调整为A4纸大小
        ppt.slide_width = round(210 / 25.4 * 914400)
        ppt.slide_height = round(297 / 25.4 * 914400)

        doc = fitz.open(pdf_file)
        for i in range(doc.page_count):
            page = doc.load_page(i)
            pix = page.get_pixmap(zoom=2)
            img = path_load + f'page_{i}.png'
            pix.save(img)
            print(img)

            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            slide.shapes.add_picture(img, 0, 0, ppt.slide_width, ppt.slide_height)

        ppt.save(file_path + file_name + '.pptx')
        print("成功创建PPT")
        try:
            shutil.rmtree("ppt_img_temps")
        except:
            pass
        msgBox = QMessageBox()
        msgBox.setWindowTitle("成功！")
        msgBox.setText("生成的PPT已成功生成至原来目录" )
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()
        self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 

       

    def onFinished(self):
        # 任务完成后的处理，比如弹窗提示完成，重启按钮等
        self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 
        self.loading_label.hide()   # 隐藏加载动画
        msgBox = QMessageBox()
        msgBox.setWindowTitle("成功！")
        msgBox.setText("处理的PDF已成功生成至原来目录" )
        msgBox.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.exec_()

    def onError(self, message):
        # 出错处理，显示错误信息等
        QMessageBox.critical(self, "Error", message)
        self.loading_label.hide()   # 隐藏加载动画
        self.pushButton.setEnabled(True), self.pushButton_2.setEnabled(True), self.pushButton_3.setEnabled(True), self.pushButton_4.setEnabled(True) ,self.pushButton_5.setEnabled(True) 

class WorkerThread_Excel(QThread):
    finished = pyqtSignal()  # 任务完成信号
    error = pyqtSignal(str)  # 错误信号，可以传递错误信息

    def __init__(self, pdf_file_path):
        super().__init__()
        self.pdf_file_path = pdf_file_path
        pdf_file_name = os.path.basename(pdf_file_path)
        self.pdf_file_name = pdf_file_name

    def run(self):   
        try:
            pdf_file_name = self.pdf_file_name
            file_name = pdf_file_name.split(".")[-2]
            # file_path = '/' + self.pdf_file_path.rstrip(pdf_file_name)      # mac
            file_path = self.pdf_file_path.rstrip(pdf_file_name)      # Windows
            # 将 pdfToExcel 方法中的 PDF 处理逻辑放在这里
            wb = Workbook()  # 创建文件对象
            ws = wb.active  # 获取第一个sheet
            # path = os.getcwd()+"/微信支付交易明细证明(20200101-20201231).pdf" #当前路径下的pdf文件
            # path = '/' + self.pdf_file_path # mac
            path = self.pdf_file_path      # Windows
            print("测试：读取路径成功：", path)
            pdf = pdfplumber.open(path) #打开pdf文件
            print('\n')
            print('开始读取数据')
            
            # 第一页第一行标题，解析只对规整的表格有用，凸(*皿* )！！
            # ws.append(pdf.pages[0].extract_tables()[0][0])
            for page in pdf.pages:
                # 获取当前页面的全部文本信息，包括表格中的文字
                #print(page.extract_text())
                for table in page.extract_tables():
                    # print(table)
                    for row in table:
                        # print(row)
                        #把列表拆了
                        rowlist=str(row).replace("[","",).replace("]","").replace("'","").replace("\\n","").split(",")
                        #print(rowlist)
                        ws.append(rowlist)
                    print('---------- 分割线 ----------')
            pdf.close()
            # 保存Excel表到22.xlsx,直接替换,注意保存
            endfile = file_path + file_name + '.xlsx'
            wb.save(endfile)

            self.finished.emit()  # 任务完成后发出完成信号
        except Exception as e:
            self.error.emit(str(e))  # 出错时发出错误信号

class WorkerThread_WORD(QThread):
    finished = pyqtSignal()  # 任务完成信号
    error = pyqtSignal(str)  # 错误信号，可以传递错误信息

    def __init__(self, pdf_file_path, checked):
        super().__init__()
        self.pdf_file_path = pdf_file_path
        pdf_file_name = os.path.basename(pdf_file_path)
        self.pdf_file_name = pdf_file_name
        self.checked = checked

    def run(self):   
        pdf_file_path = self.pdf_file_path
        try:
            if not self.checked:       # 确定是否勾选OCR
                # doc = docx.Document()
                # paragraph = doc.add_paragraph()
                # pdf = pdfplumber.open(pdf_file_path)
                # a = len(pdf.pages)       # 获取文件总页数
                # for i in range(0,a):
                #     page = pdf.pages[i]
                #     text = page.extract_text()
                #     print(text)
                #     print('------正在存入第'+ str(i)+ '页内容--------')
                #     paragraph.add_run(text)
                
                docx_file = pdf_file_path.rstrip(".pdf") + ".docx"
                # endfile = file_name + '.docx'
                # doc.save(endfile)
                cv = docxConverter(pdf_file_path)
                cv.convert(docx_file, start=0, end=None)
                cv.close
                self.finished.emit()
                
            else:       # 使用OCR
                # os.environ['TESSDATA_PREFIX'] = r'Tesseract-OCR\tessdata'
                pdf_file_name = os.path.basename(pdf_file_path)
                file_name = pdf_file_name.split(".")[-2]
                file_path = pdf_file_path.rstrip(pdf_file_name)
                img_start = 0
                # img_start = img_start - 1
                img_end = 0
                # try:
                if not os.path.exists("img_temps"):     # 检查是否已经存在文件夹
                    os.mkdir("img_temps")
                else:
                    shutil.rmtree("img_temps")
                    os.mkdir("img_temps")
                print("进行到这里了1111")
                with pdfplumber.open(pdf_file_path) as pdf:
                    pages_len = len(pdf.pages)
                    img_end = pages_len
                    # for i, page in enumerate(pdf.pages[:2]):
                    for i, page in enumerate(pdf.pages[img_start:img_end]):
                        im = page.to_image(resolution=300)
                        im.save(r"img_temps/" +'{}.png'.format(i + 1))
                        print('------分割线，第%d页-------'%(int(i)+1))
                print("进行到这里了2222")
                MainWindow_path = os.getcwd()   # 获取当前路径
                pytesseract.pytesseract.tesseract_cmd = MainWindow_path + r'\Tesseract-OCR\tesseract.exe'
                os.environ['TESSDATA_PREFIX'] = MainWindow_path + r'\Tesseract-OCR\tessdata'
                img_dir = 'img_temps'
                doc = Document()

                for img_name in os.listdir(img_dir):
                    if img_name.endswith('.jpg') or img_name.endswith('.png'):
                        image_path = os.path.join(img_dir, img_name)
                        print(image_path)
                        print("进行到这里了3333")
                        # text = str(((pytesseract.image_to_string(Image.open(image_path), lang='chi_sim',config='--tessdata-dir \Tesseract-OCR\tessdata'))))
                        text = str(((pytesseract.image_to_string(Image.open(image_path), lang='chi_sim'))))
                        doc.add_paragraph(text)
                print("进行到这里了4444")
                doc.save(file_path + file_name + ".docx")

                shutil.rmtree("img_temps")  # 删除原来文件夹
                
                
                self.finished.emit()  # 任务完成后发出完成信号
        except Exception as e:
            self.error.emit(str(e))  # 出错时发出错误信号


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = Ui_MainWindow()
    icon_path = QIcon(file_load + r'/image/icon.ico')
    app.setWindowIcon(icon_path)
    ui.show()
    sys.exit(app.exec_())
