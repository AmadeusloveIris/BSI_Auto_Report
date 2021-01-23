import os
import sys
import autor_window
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog
from glob import glob
from report_generator import Report


class MainDialog(QDialog):
    def __init__(self, parent=None):
        super(QDialog, self).__init__(parent)
        self.ui = autor_window.Ui_Report_Auto_Generator()
        self.ui.setupUi(self)

    def get_input_path(self):
        input_pwd = QFileDialog.getExistingDirectory(self, "选取PEAKS生成报告文件夹")
        self.ui.Input_path.setText(input_pwd)

    def get_output_path(self):
        output_pwd, _ = QFileDialog.getSaveFileName(self, "选取最终报告位置及名称", filter="Docx Files (*.docx)")
        self.ui.Output_path.setText(output_pwd)

    def generate_report(self):
        self.ui.stautus.setText('Loading necessary information......')
        report = Report(self.ui.Input_path.text(), self.ui.onumber.text(),
                        self.ui.sname.text(), self.ui.hrmass.text(), self.ui.lrmass.text())
        self.ui.stautus.setText('Writing Report......')
        report.tpl.render(report.context, report.jinja_env)
        report.tpl.save(self.ui.Output_path.text())
        for f in glob(os.path.join('temp', '*')):
            os.remove(f)
        self.ui.stautus.setText('Complete')


if __name__ == '__main__':
    myapp = QApplication(sys.argv)
    myDlg = MainDialog()
    myDlg.show()
    sys.exit(myapp.exec_())
