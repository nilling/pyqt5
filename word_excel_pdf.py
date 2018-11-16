# -*- coding: utf-8 -*-


# Author: nil
# Date: 2018/11/15
# Doc: 打开word，excel，pdf文件

from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox


class AxWidget(QWidget):
    def __init__(self):
        super(AxWidget, self).__init__()
        self.resize(800, 600)
        layout = QVBoxLayout(self)
        self.axWidget = QAxWidget(self)
        layout.addWidget(self.axWidget)
        layout.addWidget(QPushButton('选择excel, word, pdf文件', self, clicked=self.onOpenFile))

    def onOpenFile(self):
        print 45
        path, _ = QFileDialog.getOpenFileName(self, '请选择文件', '', 'excel(*.xlsx *.xls);;word(*.docx *.doc);;pdf(*.pdf)')
        print path
        print _
        if not path:
            return
        if _.find('.doc') != -1:
            print 2
            return self.openOffice(path, 'Word.Application')
        if _.find('.xls') != -1:
            print 3
            return self.openOffice(path, 'Excel.Application')
        if _.find('.pdf') != -1:
            print 11
            return self.openPdf(path)

    def openOffice(self, path, app):
        self.axWidget.clear()
        if not self.axWidget.setControl(app):
            return QMessageBox.critical(self, '错误', '没有安装 %s' % app )
        self.axWidget.dynamicCall('SetVisible (bool Visible)', 'false')
        self.axWidget.setProperty('DisplayAlters', False)
        self.axWidget.setControl(path)

    def openPdf(self, path):
        self.axWidget.clear()
        if not self.axWidget.setControl('Adobe PDF Reader'):
            return QMessageBox.critical(self, '错误', '没有安装 adobe pdf reader')
        print 1
        self.axWidget.dynamicCall('LoadFile(const QString&)', path)

    def closeEvent(self, event):
        self.axWidget.close()
        self.axWidget.clear()
        self.layout().removeWidget(self.axWidget)
        del self.axWidget
        super(AxWidget, self).closeEvent(event)




if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    w = AxWidget()
    w.show()
    sys.exit(app.exec_())