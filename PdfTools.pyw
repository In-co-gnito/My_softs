from PyQt5.QtWidgets import (QGroupBox,QWidget,QDockWidget,QMessageBox,QFileDialog,QLabel,QAbstractItemView,QListWidget,
                            QLineEdit,QFormLayout,QApplication,QPushButton,QMainWindow,QVBoxLayout,QHBoxLayout)
from PyQt5 import QtGui
from PyQt5 import QtCore
from PyPDF2 import PdfFileMerger,PdfFileReader,PdfFileWriter
import sys

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Window
        self.setWindowIcon(QtGui.QIcon("book.png"))
        self.setWindowTitle("PDF Tools")
        self.setFixedSize(800,350)
        
        
        # Error message box
        self.errmsg = QMessageBox()
        self.errmsg.setIcon(QMessageBox.Critical)
        self.errmsg.setWindowIcon(QtGui.QIcon("angry.png"))
        self.errmsg.setWindowTitle("Error !")
        self.errmsg.setStandardButtons(QMessageBox.Ok)

        # Opening Dockwidget
        self.dockwindow = QDockWidget()
        self.dockwindow.setMinimumSize(400,200)
        self.addDockWidget(QtCore.Qt.RightDockWidgetArea,self.dockwindow)
        self.dockwindow.setFeatures(QDockWidget.DockWidgetFloatable|QDockWidget.DockWidgetMovable)
        
        # Initializing UI
        self.win1 = QWidget(self)
        self.setCentralWidget(self.win1)
        self.initUI()
        self.extractui()
    
    def extract(self):
        path = self.filepath.text()
        outpath = self.outfilepath.text()
        start = None
        end = None
        try:
            pdf = PdfFileReader(path,"rb")
            pdf_writer = PdfFileWriter()
            start = int(self.startpage.text())
            end = int(self.endpage.text())
        except Exception:
            self.errmsg.setText("Enter valid entries !")
            self.errmsg.exec_()

        if type(start) == int and type(end) == int and start >0 and end<=pdf.numPages:
            try:        
                if start <= end:
                    for page in range(start-1,end):
                        pdf_writer.addPage(pdf.getPage(page))
                else:
                    for page in range(start-1,end-2,-1):
                        pdf_writer.addPage(pdf.getPage(page))
                with open(outpath,"wb") as out:
                    pdf_writer.write(out)
            except Exception:
                self.errmsg.setText("Enter valid entries ! !")
                self.errmsg.exec_()
    
    def getnewfilepath(self):
        self.filepath.clear()
        newfile = QFileDialog.getOpenFileName(self,"Open File",".","(*.pdf)")
        self.filepath.insert(newfile[0])
    
    def getnewfilespath(self):
        self.files = QFileDialog.getOpenFileNames(self,"Select files",".","(*.pdf)")
        self.newfilespath.addItems(self.files[0])

    def getsavefilepath(self):
        self.outfilepath.clear()
        savefile = QFileDialog.getSaveFileName(self,"Open Folder",".","(*.pdf)")
        self.outfilepath.insert(savefile[0])
    
    def extractui(self):
        self.dockwindow.setWindowTitle("Extract PDF")       
        w1 = QWidget(self)
        layout = QFormLayout()

        row1 = QHBoxLayout()
        self.filepath = QLineEdit()
        
        self.filebutton = QPushButton("Browse ")
        self.filebutton.clicked.connect(self.getnewfilepath)
        row1.addWidget(self.filepath)
        row1.addWidget(self.filebutton)
        layout.addRow(QLabel("New File Path"),row1)

        row2 = QHBoxLayout()
        self.outfilepath = QLineEdit()
        
        self.outfilebutton = QPushButton("Browse ")
        self.outfilebutton.clicked.connect(self.getsavefilepath)
        row2.addWidget(self.outfilepath)
        row2.addWidget(self.outfilebutton)
        layout.addRow(QLabel("Output File Path"),row2)

        self.startpage = QLineEdit()
        self.endpage = QLineEdit()
        extractbut = QPushButton("Extract")
        extractbut.clicked.connect(self.extract)
        layout.addRow("Starting Page : ",self.startpage)
        layout.addRow("Ending Page : ",self.endpage)
        layout.addRow(extractbut)
        w1.setLayout(layout)
        self.dockwindow.setWidget(w1)

    def splitui(self):
        self.dockwindow.setWindowTitle("Extract PDF")
        w2 = QWidget(self)
        layout = QFormLayout()

        row1 = QHBoxLayout()
        self.filepath = QLineEdit()
        self.filepath.setReadOnly(True)
        self.filebutton = QPushButton("Browse ")
        self.filebutton.clicked.connect(self.getnewfilepath)
        row1.addWidget(self.filepath)
        row1.addWidget(self.filebutton)
        layout.addRow(QLabel("New File Path"),row1)

        row2 = QHBoxLayout()
        self.outfilepath = QLineEdit()
        self.outfilepath.setReadOnly(True)
        self.outfilebutton = QPushButton("Browse ")
        self.outfilebutton.clicked.connect(self.getsavefilepath)
        row2.addWidget(self.outfilepath)
        row2.addWidget(self.outfilebutton)
        layout.addRow(QLabel("Output File Path"),row2)

        self.startpage = QLineEdit()
        self.endpage = QLineEdit()
        extractbut = QPushButton("Extract")
        extractbut.clicked.connect(self.extract)
        layout.addRow("First Limit : ",self.startpage)
        layout.addRow("Second Limit : ",self.endpage)
        layout.addRow(extractbut)
        w2.setLayout(layout)
        self.dockwindow.setWidget(w2)
    
    def addwatermark(self):
        self.dockwindow.setWindowTitle("Add Watermark")

    def reorderpdf(self):
        self.dockwindow.setWindowTitle("Reorder Pdf")

    def mergeui(self):
        self.dockwindow.setWindowTitle("Merge Pdf")
        w2 = QWidget(self)
        layout = QFormLayout()

        row1 = QHBoxLayout()
        self.outfilepath = QLineEdit()
        self.filebutton = QPushButton("Browse")
        self.filebutton.clicked.connect(self.getsavefilepath)
        row1.addWidget(self.outfilepath)
        row1.addWidget(self.filebutton)
        layout.addRow(QLabel("Output File Location"),row1)

        row2 = QHBoxLayout()
        row3 = QVBoxLayout()
        self.newfilespath = QListWidget()
        self.newfilespath.setMinimumSize(300,200)
        self.newfilespath.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.addbutton = QPushButton()
        self.addbutton.setFixedSize(40,40)
        self.addbutton.setIcon(QtGui.QIcon("PLUS.png"))
        self.deletebutton = QPushButton()
        self.deletebutton.setIcon(QtGui.QIcon("delete.png"))
        self.deletebutton.setFixedSize(40,40)
        self.upbutton = QPushButton()
        self.upbutton.setIcon(QtGui.QIcon("Up.png"))
        self.upbutton.setFixedSize(40,40)
        self.downbutton = QPushButton()
        self.downbutton.setIcon(QtGui.QIcon("down.png"))
        self.downbutton.setFixedSize(40,40)
        row2.addWidget(self.addbutton)
        row2.addWidget(self.deletebutton)
        row2.addWidget(self.upbutton)
        row2.addWidget(self.downbutton)
        row2.setAlignment(QtCore.Qt.AlignHCenter)
        self.addbutton.clicked.connect(self.getnewfilespath)
        self.deletebutton.clicked.connect(self.removeitem)
        row3.addWidget(self.newfilespath)
        layout.addRow(row2)
        layout.addRow(row3)

        mergebut = QPushButton("Merge")
        mergebut.clicked.connect(self.mergepdf)
        layout.addRow(mergebut)
        w2.setLayout(layout)
        self.dockwindow.setWidget(w2)   

    def removeitem(self):
        listitems = self.newfilespath.selectedItems()
        for item in listitems:
            self.newfilespath.takeItem(self.newfilespath.row(item))

    def mergepdf(self):
        merger = PdfFileMerger()      
        pdfs = []
        for i in range(self.newfilespath.count()):
            pdfs.append(self.newfilespath.item(i).text())
        print(pdfs)
        try:
            for pdf in pdfs:
                with open(pdf,"rb") as p:
                    merger.append(PdfFileReader(p))
            with open(self.outfilepath.text(),"wb") as out:
                merger.write(out)
        except Exception:
            self.errmsg.setText("Enter valid entries !")
            self.errmsg.exec_()

    def pdftotxt(self):
        self.dockwindow.setWindowTitle("Pdf to Text")
   
    def rotatepdf(self):
        self.dockwindow.setWindowTitle("Rotate Pdf pages")

    # ENCRYPT METHODS:

    def encryptgui(self):
        path = self.filepath.text()
        output_file = self.outfilepath.text()
        encrypted = None
        try:
            pdf = PdfFileReader(path,"rb")
            if pdf.getIsEncrypted():
                encrypted = True
                self.errmsg1 = QMessageBox()
                self.errmsg1.setIcon(QMessageBox.Warning)
                self.errmsg1.setWindowIcon(QtGui.QIcon("angry.png"))
                self.errmsg1.setWindowTitle("Warning !")
                self.errmsg1.setText("The file is already encrypted")
                self.errmsg1.setStandardButtons(QMessageBox.Ok)
                self.errmsg1.buttonClicked.connect(self.clearingvalues)
                self.errmsg1.exec_()     
            pdf_writer = PdfFileWriter()

        except Exception:
            self.errmsg.setText("Enter valid entries ! !")
            self.errmsg.exec_()
            
        if encrypted == False:
            try:
                for page in range(pdf.getNumPages()):
                    pdf_writer.addPage(pdf.getPage(page))
                        
                pdf_writer.encrypt(user_pwd=self.passwordname.text())
                with open(output_file,"wb") as out:
                    pdf_writer.write(out)
            except Exception:
                self.errmsg.setText("Enter valid entries ! !")
                self.errmsg.exec_()
    
    def clearingvalues(self,i):
        self.filepath.clear()
        self.outfilepath.clear()
        self.passwordname.clear()
              
    def encryptinggui(self):
        self.dockwindow.setWindowTitle("Encrypt Pdf")
        w1 = QWidget(self)
        layout = QFormLayout()

        row1 = QHBoxLayout()
        self.filepath = QLineEdit()
        self.filepath.setReadOnly(True)
        self.filebutton = QPushButton("Browse ")
        self.filebutton.clicked.connect(self.getnewfilepath)
        row1.addWidget(self.filepath)
        row1.addWidget(self.filebutton)
        layout.addRow(QLabel("New File Path"),row1)

        row2 = QHBoxLayout()
        self.outfilepath = QLineEdit()
        self.outfilepath.setReadOnly(True)
        self.outfilebutton = QPushButton("Browse ")
        self.outfilebutton.clicked.connect(self.getsavefilepath)
        row2.addWidget(self.outfilepath)
        row2.addWidget(self.outfilebutton)
        layout.addRow(QLabel("Output File Path"),row2)

        self.passwordname = QLineEdit()
        layout.addRow("Password :",self.passwordname)

        encryptbut = QPushButton("Encrypt")
        encryptbut.clicked.connect(self.encryptgui)
        layout.addRow(encryptbut)
        w1.setLayout(layout)
        self.dockwindow.setWidget(w1)
     
    def initUI(self):

        layout1 = QHBoxLayout()
        layout2 = QVBoxLayout()
        self.extractbutton = QPushButton("Extract PDF",self)
        self.extractbutton.setFixedSize(150,50)
        self.extractbutton.setIcon(QtGui.QIcon("paper.png"))
        self.extractbutton.setIconSize(QtCore.QSize(40,40))
        self.extractbutton.setToolTip("Extract PDF Pages in given range")
        self.extractbutton.clicked.connect(self.extractui)
        layout2.addWidget(self.extractbutton)

        self.splitbutton = QPushButton("Split PDF",self)
        self.splitbutton.setFixedSize(150,50)
        self.splitbutton.setIcon(QtGui.QIcon("split.png"))
        self.splitbutton.setIconSize(QtCore.QSize(40,40))
        self.splitbutton.setToolTip("Split PDF into parts")
        self.splitbutton.clicked.connect(self.splitui)
        layout2.addWidget(self.splitbutton)

        self.waterbutton = QPushButton("Add Watermark",self)
        self.waterbutton.setFixedSize(150,50)
        self.waterbutton.setIcon(QtGui.QIcon("watermark.png"))
        self.waterbutton.setIconSize(QtCore.QSize(40,40))
        self.waterbutton.setToolTip("Add WaterMark to PDF Pages")
        self.waterbutton.clicked.connect(self.addwatermark)
        layout2.addWidget(self.waterbutton)

        self.reorderbutton = QPushButton("Reorder PDF",self)
        self.reorderbutton.setFixedSize(150,50)
        self.reorderbutton.setIcon(QtGui.QIcon("number.png"))
        self.reorderbutton.setIconSize(QtCore.QSize(40,40))
        self.reorderbutton.setToolTip("Edit Pdf by reordering the pdf pages")
        self.reorderbutton.clicked.connect(self.reorderpdf)
        layout2.addWidget(self.reorderbutton)

        layout1.addLayout(layout2)
        
        layout3 = QVBoxLayout()
        self.mergebutton = QPushButton("Merge PDF",self)
        self.mergebutton.setFixedSize(150,50)
        self.mergebutton.setIcon(QtGui.QIcon("merge.png"))
        self.mergebutton.setIconSize(QtCore.QSize(40,40))
        self.mergebutton.setToolTip("Merge multiple PDF files")
        self.mergebutton.clicked.connect(self.mergeui)
        layout3.addWidget(self.mergebutton)

        self.rotatebutton = QPushButton("Rotate PDF",self)
        self.rotatebutton.setFixedSize(150,50)
        self.rotatebutton.setIcon(QtGui.QIcon("rotate.png"))
        self.rotatebutton.setIconSize(QtCore.QSize(40,40))
        self.rotatebutton.setToolTip("Rotate PDF Pages")
        self.rotatebutton.clicked.connect(self.rotatepdf)
        layout3.addWidget(self.rotatebutton)

        self.txtbutton = QPushButton(".pdf to .txt",self)
        self.txtbutton.setFixedSize(150,50)
        self.txtbutton.setIcon(QtGui.QIcon("pencil.png"))
        self.txtbutton.setIconSize(QtCore.QSize(40,40))
        self.txtbutton.setToolTip("Convert pdf to text\nApplicable if PDF contains text")
        self.txtbutton.clicked.connect(self.pdftotxt)
        layout3.addWidget(self.txtbutton)

        self.passbutton = QPushButton("  Encrypt PDF",self)
        self.passbutton.setFixedSize(150,50)
        self.passbutton.setIcon(QtGui.QIcon("read-only.png"))
        self.passbutton.setIconSize(QtCore.QSize(40,40))
        self.passbutton.setToolTip("Add Password to PDF")
        self.passbutton.clicked.connect(self.encryptinggui)
        layout3.addWidget(self.passbutton)

        layout1.addLayout(layout3)
        self.win1.setLayout(layout1)

def main():
    App = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(App.exec_())

if __name__ == "__main__":
    main()
