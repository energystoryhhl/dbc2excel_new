
from PyQt5 import  QtGui, QtWidgets, QtCore
# from PyQt5.QtWidgets import QFileDialog,QTreeWidgetItem,QTreeWidget
from PyQt5.QtWidgets import *
from PyQt5.QtCore import pyqtSignal
import PyQt5.QtCore
import sys
import main_wd
import os

# from canmatrix.xls_common import *
import canmatrix
import canmatrix.formats
import threading
import canmatrix.convert
import canmatrix.canmatrix as cm
import canmatrix.copy as cmcp
import random



class convert_exec_thread(QtCore.QThread):
    _signal = pyqtSignal(str)
    _file_btn_signal = pyqtSignal(bool)
    def __init__(self, out_put_file, option, dbs): 
        super(convert_exec_thread, self).__init__()
        self.out_put_file = out_put_file
        self.option = option
        self.dbs = dbs
    def run(self):
        options = self.option
        dbs = self.dbs
        outdbs = {}
        for name in dbs:
            db = None

            if 'ecus' in options and options['ecus'] is not None:
                ecuList = options['ecus'].split(',')
                db = cm.CanMatrix()
                for ecu in ecuList:
                    cmcp.copyBUwithFrames(ecu, dbs[name], db)
            if 'frames' in options and options['frames'] is not None:
                frameList = options['frames'].split(',')
                db = cm.CanMatrix()
                for frame in frameList:
                    cmcp.copyFrame(frame, dbs[name], db)
            if db is None:
                db = dbs[name]

            outdbs[name] = db
        out = canmatrix.formats.dumpp(outdbs, self.out_put_file,bcExportEncoding='iso-8859-1',)

        self._signal.emit("Convert test!")
        if out == -1:
            self._signal.emit("Convert Failed!")
            self.open_convert_done_file_btn.setEnabled(False)
        else:
            self._signal.emit("Convert Success!")
            self._signal.emit("Out put file: "  + self.out_put_file)
            self.convert_done_file_name = self.out_put_file
            self._file_btn_signal.emit(True)


class dbc2excel(main_wd.Ui_MainWindow):
    open_file_signal = pyqtSignal()
    main_tittle = "Dbc2Excel"
    random_tittle = [" - 极限之外无限地大 传说一直反覆", " - Do your best for your sake anytime", \
                     " - You can have your way in everything"]

    def tree_widget_chanaged(self,item, p_int):
        # print(item.parent())
        # print(item.columnCount())
        # print(p_int)
        # print(QTreeWidgetItem(item))

        if item.parent() == None:
            # print("root")
            if p_int >= 3:
                pass
            else:
                # pass
                print("root")
                ##doo

        else:
            if p_int <= 2:
                pass
            else:
                # pass
                print("child")
            #do

    def expand_all(self):
        if self.dbs != None:
            self.treeWidget.expandAll()
        else:
            pass

    def put_away_all(self):
        if self.dbs != None:
            self.treeWidget.collapseAll()

    def app_quit(self):
        exit(0)

    def close_btn_cb(self):
        if self.dbs != None:
            self.treeWidget.clear()
            self.dbs = None
            self.file_name = None


    def open_btn_cb(self):

        fname = QFileDialog.getOpenFileName(None, 'Open file', '.','Files(*.dbc *.dbf *.kcd *.arxml *.json *.xls *.sym *.xlsx)')
        if fname[0] == '':
            pass
        else:
            if self.dbs != None:
                self.treeWidget.itemChanged.disconnect()
            option = dict()
            name = fname[0]
            print(name)
            self.file_name = name
            self.dbs = canmatrix.formats.loadp(name, **option)
            self.textEdit.append("Open File: " + name)
            if self.dbs != None:
                self.treeWidget.clear()
                # self.treeWidget.setEditTriggers(QTreeWidget.clicked)
                total_num = 0
                for BU in self.dbs['']:
                    root = QTreeWidgetItem(self.treeWidget)
                    root.setText(0, str(hex(BU.arbitration_id.id)))
                    root.setText(1, BU.name)
                    root.setText(2, str(BU.cycle_time))
                    root.setFlags(PyQt5.QtCore.Qt.ItemIsEnabled | PyQt5.QtCore.Qt.ItemIsEditable)

                    # root.setFlags(root.flags())
                    # print(BU)
                    total_num += 1

                    for SG in BU.signals:
                        child =QTreeWidgetItem()
                        child.setText(3, SG.name)
                        child.setText(4, str(SG.start_bit))
                        child.setText(5, str(SG.size))
                        child.setText(6, str(SG.max))
                        child.setText(7, str(SG.min))
                        child.setText(8, str(SG.initial_value))
                        child.setText(9, str(SG.is_little_endian))
                        child.setText(10, str(SG.factor))
                        child.setText(11, str(SG.offset))
                        child.setText(12, str(SG.is_signed))
                        child.setText(13, str(SG.values))
                        child.setFlags(PyQt5.QtCore.Qt.ItemIsEnabled | PyQt5.QtCore.Qt.ItemIsEditable)
                        root.addChild(child)

                self.treeWidget.itemChanged.connect(self.tree_widget_chanaged)
                self.treeWidget.expandAll()
                self.textEdit.append("Total BU: " + str(total_num))
                # print(self.dbs)
            else:
                self.textEdit.append("Open File Failed" )

    def convert_exec(self,out_put_file,option):
        print("convert")
        print(out_put_file)
        options = self.option

        dbs = self.dbs
        outdbs = {}
        for name in dbs:
            db = None

            if 'ecus' in options and options['ecus'] is not None:
                ecuList = options['ecus'].split(',')
                db = cm.CanMatrix()
                for ecu in ecuList:
                    cmcp.copyBUwithFrames(ecu, dbs[name], db)
            if 'frames' in options and options['frames'] is not None:
                frameList = options['frames'].split(',')
                db = cm.CanMatrix()
                for frame in frameList:
                    cmcp.copyFrame(frame, dbs[name], db)
            if db is None:
                db = dbs[name]

            outdbs[name] = db

        out = canmatrix.formats.dumpp(outdbs, out_put_file)
        self.signal.emit("Convert test!")   
        if out == -1:
            self.textEdit.append("Convert Failed!")
            self.open_convert_done_file_btn.setEnabled(False)
        else:
            self.textEdit.append("Convert Success!")
            self.textEdit.append("Out put file: "  + out_put_file)
            self.convert_done_file_name = out_put_file
            self.open_convert_done_file_btn.setEnabled(True)


    def on_convert_cb(self):
        if self.dbs == None:
            self.textEdit.append("no in pust file!")
        else:
            out_file_type = "xlsx"
            if  str(self.file_name).split('.')[-1] == "dbc":
                pass
            else:
                out_file_type = "dbc"

            out_file_name = str(self.file_name).split(".")[0] + '.' +out_file_type

            file_path = QFileDialog.getSaveFileName(None, "output file", "./" + out_file_name ,
                                                    "Files(*.dbc *.dbf *.kcd *.arxml *.json *.xls *.sym *.xlsx)")
            file_name = file_path[0]
            if file_name == '':
                pass
            else:
                self.thread = convert_exec_thread(file_name,self.option,self.dbs)
                self.thread._signal.connect(self.append_text)
                self.thread._file_btn_signal.connect(self.set_file_btn_state)
                self.thread.start()

    def append_text(self,data):
        self.textEdit.append(data)

    def set_file_btn_state(self, state):
        self.open_convert_done_file_btn.setEnabled(state)

    def open_convert_done_file_cb(self):
        if self.convert_done_file_name != None:
            self.textEdit.append("Open File: " + self.convert_done_file_name)
            os.startfile(self.convert_done_file_name)



    def __init__(self, Dialog):
        super().setupUi(Dialog)  # 调用父类的setupUI函数
        Dialog.setWindowTitle(self.main_tittle + self.random_tittle[random.randint(1, len(self.random_tittle)) - 1])

        self.tree_widget_header = ['Id','FrameName','CycleTime','SignalName','StartBit','Length','Max','Min',\
                                   'Default','LittleEndian','Factor','Offset','isSigned','Value']
        self.treeWidget.setHeaderLabels(self.tree_widget_header)

        self.open_btn.clicked.connect(self.open_btn_cb)
        self.option = {'unsetFrameFd': None, 'deleteFrame': None, 'additionalAttributes': '', 'xlsMotorolaBitFormat': 'msbreverse',\
                       'deleteSignalAttributes': None, 'jsonMotorolaBitFormat': 'lsb', 'deleteZeroSignals': False, 'jsonAll': False, \
                       'arxmlIgnoreClusterInfo': False, 'deleteFrameAttributes': None, 'silent': False, 'setFrameFd': None, \
                       'cutLongFrames': None, 'changeFrameId': None, 'verbosity': 0, 'dbcExportEncoding': 'iso-8859-1', \
                       'symExportEncoding': 'iso-8859-1', 'dbfImportEncoding': 'iso-8859-1', 'force_output': None, 'dbfExportEncoding': 'iso-8859-1',\
                       'jsonCanard': False, 'skipLongDlc': None, 'addFrameReceiver': None, 'dbcImportCommentEncoding': 'iso-8859-1', 'ecus': None, \
                       'deleteEcu': None, 'dbcImportEncoding': 'iso-8859-1', 'frames': None, 'arVersion': '3.2.3', 'merge': None, 'symImportEncoding': 'iso-8859-1',\
                       'recalcDLC': False, 'renameFrame': None, 'renameEcu': None, 'renameSignal': None, 'additionalFrameAttributes': '', \
                       'dbcExportCommentEncoding': 'iso-8859-1', 'deleteSignal': None, 'frameIdIncrement': None, 'deleteObsoleteDefines': False}

        self.dbs = None
        self.file_name = None
        self.convert_done_file_name = None

        self.action_open.triggered.connect(self.open_btn_cb)
        self.action_close.triggered.connect(self.close_btn_cb)
        self.action_exit.triggered.connect(self.app_quit)
        self.actione_expand_all.triggered.connect(self.expand_all)
        self.action_put_away_all.triggered.connect(self.put_away_all)
        self.convert_btn.clicked.connect(self.on_convert_cb)
        self.open_convert_done_file_btn.clicked.connect(self.open_convert_done_file_cb)
        self.open_convert_done_file_btn.setEnabled(False)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = dbc2excel(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
