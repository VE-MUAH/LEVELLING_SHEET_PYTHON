import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLineEdit, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QHBoxLayout,
    QGridLayout, QMenuBar, QMenu, QAction, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
import math


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(150, 150, 1000, 600)
        self.setWindowTitle('VICENTIA LEVELLING SHEET')
        # self.setStyleSheet('background-color: blue')

        self.tablewidget = QTableWidget()
        self.tablewidget.setRowCount(10000)
        self.tablewidget.setColumnCount(10000)
        self.tablewidget.setHorizontalHeaderItem(0, QTableWidgetItem('BS'))
        self.tablewidget.setHorizontalHeaderItem(1, QTableWidgetItem('IS'))
        self.tablewidget.setHorizontalHeaderItem(2, QTableWidgetItem('FS'))
        self.tablewidget.setHorizontalHeaderItem(3, QTableWidgetItem('HPC'))
        self.tablewidget.setHorizontalHeaderItem(4, QTableWidgetItem('IRL'))
        self.tablewidget.setHorizontalHeaderItem(5, QTableWidgetItem('ADJUST'))
        self.tablewidget.setHorizontalHeaderItem(6, QTableWidgetItem('FRL'))
        self.tablewidget.setHorizontalHeaderItem(7, QTableWidgetItem('REMARKS'))

        self.initialBMlabel = QLabel('INITIAL BM')
        self.initialBMinput = QLineEdit()
        self.finalBMlabel = QLabel('FINAL BM')
        self.finalBMinput = QLineEdit()
        self.klabel = QLabel('K Value')
        self.kinput = QLineEdit()
        self.calculatebutton = QPushButton('Calculate')
        self.clearbutton = QPushButton('Clear')
        self.clearbutton.clicked.connect(self.clear_table)
        self.calculatebutton.clicked.connect(self.calculate)

        self.arithmetic_check_label = QLabel('ARITHMETIC CHECK')
        self.sum_bs_label = QLabel('SUM(BS):')
        self.sum_fs_label = QLabel('SUM(FS):')
        self.last_irl_label = QLabel('LAST IRL:')
        self.first_irl_label = QLabel('FIRST IRL:')
        self.arithmetic_result_label = QLabel('RESULT:')
        self.sum_bs_value = QLabel('')
        self.sum_fs_value = QLabel('')
        self.last_irl_value = QLabel('')
        self.first_irl_value = QLabel('')
        self.arithmetic_result_value = QLabel('')
        self.allowable_misclose_label = QLabel('ALLOWABLE MISCLOSE:')
        self.allowable_misclose_value = QLabel('')
        self.misclose_label = QLabel('MISCLOSE:')
        self.misclose_value = QLabel('')
        self.acceptable_misclose_label = QLabel('ACCEPTABLE:')
        self.acceptable_misclose_value = QLabel('')

        input_layout = QHBoxLayout()
        input_layout.addWidget(self.initialBMlabel)
        input_layout.addWidget(self.initialBMinput)
        input_layout.addWidget(self.finalBMlabel)
        input_layout.addWidget(self.finalBMinput)
        input_layout.addWidget(self.klabel)
        input_layout.addWidget(self.kinput)
        input_layout.addWidget(self.calculatebutton)
        input_layout.addWidget(self.clearbutton)

        main_layout = QVBoxLayout()
        main_layout.addLayout(input_layout)
        main_layout.addWidget(self.tablewidget)

        arithmetic_layout = QGridLayout()
        arithmetic_layout.addWidget(self.arithmetic_check_label, 0, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        arithmetic_layout.addWidget(self.sum_bs_label, 1, 0)
        arithmetic_layout.addWidget(self.sum_bs_value, 1, 1)
        arithmetic_layout.addWidget(self.sum_fs_label, 2, 0)
        arithmetic_layout.addWidget(self.sum_fs_value, 2, 1)
        arithmetic_layout.addWidget(self.last_irl_label, 3, 0)
        arithmetic_layout.addWidget(self.last_irl_value, 3, 1)
        arithmetic_layout.addWidget(self.first_irl_label, 4, 0)
        arithmetic_layout.addWidget(self.first_irl_value, 4, 1)
        arithmetic_layout.addWidget(self.arithmetic_result_label, 5, 0)
        arithmetic_layout.addWidget(self.arithmetic_result_value, 5, 1)
        arithmetic_layout.addWidget(self.misclose_label, 6, 0)
        arithmetic_layout.addWidget(self.misclose_value, 6, 1)
        arithmetic_layout.addWidget(self.allowable_misclose_label, 7, 0)
        arithmetic_layout.addWidget(self.allowable_misclose_value, 7, 1)
        arithmetic_layout.addWidget(self.acceptable_misclose_label, 8, 0)
        arithmetic_layout.addWidget(self.acceptable_misclose_value, 8, 1)

        main_layout.addLayout(arithmetic_layout)

        menubar = QMenuBar(self)
        file_menu = menubar.addMenu('File')

        import_action = QAction('Import', self)
        import_action.triggered.connect(self.import_data)

        export_action = QAction('Export', self)
        export_action.triggered.connect(self.export_data)

        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about_dialog)

        file_menu.addAction(import_action)
        file_menu.addAction(export_action)
        file_menu.addAction(about_action)

        main_layout.setMenuBar(menubar)
        self.setLayout(main_layout)

    def import_data(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", "CSV Files (*.csv);;All Files (*)")
        if file_name:
            with open(file_name, 'r') as file:
                lines = file.readlines()
                for row_num, line in enumerate(lines):
                    items = line.strip().split(',')
                    for col_num, item in enumerate(items):
                        self.tablewidget.setItem(row_num, col_num, QTableWidgetItem(item))

    def export_data(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Save File", "", "CSV Files (*.csv);;All Files (*)")
        if file_name:
            with open(file_name, 'w') as file:
                for row in range(self.tablewidget.rowCount()):
                    row_data = []
                    for col in range(self.tablewidget.columnCount()):
                        item = self.tablewidget.item(row, col)
                        if item and item.text():
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    file.write(','.join(row_data) + '\n')

    def show_about_dialog(self):
        QMessageBox.about(self, "About", "This is a leveling sheet application developed by : Emuah Vicentia,  "
                                         "e-mail address :  vicentiaemuah21@gmail.com/ vicentiaemuah20896360@gmail.com:"

                                         "WhatsApp line :  0246218639/ 0200202164, Geomatic Engineering student.")

    def clear_table(self):
        # Clear all the items in the table
        self.tablewidget.clearContents()

        # Reset row count to 10000 to maintain the initial state
        self.tablewidget.setRowCount(10000)

        # Clear the values of the input fields
        self.initialBMinput.clear()
        self.finalBMinput.clear()
        self.kinput.clear()

        # Clear the arithmetic check labels
        self.sum_bs_value.setText('')
        self.sum_fs_value.setText('')
        self.last_irl_value.setText('')
        self.first_irl_value.setText('')
        self.arithmetic_result_value.setText('')
        self.allowable_misclose_value.setText('')
        self.misclose_value.setText('')
        self.acceptable_misclose_value.setText('')

    def calculate(self):
        try:
            initial_bm = float(self.initialBMinput.text())
            final_bm = float(self.finalBMinput.text())
            k_value = float(self.kinput.text())
        except ValueError:
            print("Invalid BM or k values")
            return

        # Set the first IRL to the initial BM
        self.tablewidget.setItem(0, 4, QTableWidgetItem(f"{initial_bm:.3f}"))
        self.tablewidget.setItem(0, 7, QTableWidgetItem(f"BM({initial_bm:.3f})"))

        # Calculate the first HPC
        bs_item = self.tablewidget.item(0, 0)
        bs = float(bs_item.text()) if bs_item and bs_item.text() else 0
        hpc = initial_bm + bs
        self.tablewidget.setItem(0, 3, QTableWidgetItem(f"{hpc:.3f}"))

        # using the previous hpc to calculte for the irl
        prev_hpc = hpc
        irl_values = [initial_bm]  # storing IRL values
        row_count = 1

        sum_bs = bs
        sum_fs = 0
        bs_count = 1 if bs != 0 else 0

        # Inputting in the rows with your data
        for row in range(1, self.tablewidget.rowCount()):
            bs_item = self.tablewidget.item(row, 0)
            fs_item = self.tablewidget.item(row, 2)
            is_item = self.tablewidget.item(row, 1)

            # Checking if the row is empty
            if (bs_item is None or bs_item.text() == '') and \
                    (fs_item is None or fs_item.text() == '') and \
                    (is_item is None or is_item.text() == ''):
                break
            # converting them to float
            bs = float(bs_item.text()) if bs_item and bs_item.text() else 0
            fs = float(fs_item.text()) if fs_item and fs_item.text() else 0
            is_ = float(is_item.text()) if is_item and is_item.text() else 0
            # summing up bs and fs
            sum_bs += bs
            sum_fs += fs
            if bs != 0:
                bs_count += 1

            # Calculating IRL if IS is present
            if is_ != 0:
                irl = prev_hpc - is_
                self.tablewidget.setItem(row, 4, QTableWidgetItem(f"{irl:.3f}"))
                irl_values.append(irl)
                row_count += 1
                continue

            # Calculate IRL and HPC only for change points
            if bs != 0 and fs != 0:
                # Change Point
                irl = prev_hpc - fs
                hpc = irl + bs
                self.tablewidget.setItem(row, 3, QTableWidgetItem(f"{hpc:.3f}"))
                self.tablewidget.setItem(row, 7, QTableWidgetItem('Change Point'))
                prev_hpc = hpc
            else:
                irl = prev_hpc - fs if fs != 0 else prev_hpc

            self.tablewidget.setItem(row, 4, QTableWidgetItem(f"{irl:.3f}"))

            irl_values.append(irl)
            row_count += 1

        # Calculate the last IRL using the previous HPC
        if row_count > 1:
            fs_item = self.tablewidget.item(row_count - 1, 2)
            fs = float(fs_item.text()) if fs_item and fs_item.text() else 0
            last_irl = prev_hpc - fs
            self.tablewidget.setItem(row_count - 1, 4, QTableWidgetItem(f"{last_irl:.3f}"))

        # Calculate adjustment
        if irl_values:
            adjustment = last_irl - final_bm
            adjustment = -1 * adjustment  # Invert the adjustment sign for the final IRL

            self.tablewidget.setItem(row_count - 1, 5, QTableWidgetItem(f"{adjustment:.3f}"))

            # Calculate adjustment per row and apply to all rows except the last one

            adjustment_per_row = adjustment / (row_count - 1) if (row_count - 1) != 0 else 0
            for row in range(1, row_count - 1):
                self.tablewidget.setItem(row, 5, QTableWidgetItem(f"{adjustment_per_row:.3f}"))

        # Calculate FRL by adding adjustment to IRL for each row
        for row in range(row_count):
            irl_item = self.tablewidget.item(row, 4)
            adjust_item = self.tablewidget.item(row, 5)

            if irl_item and adjust_item:
                irl = float(irl_item.text())
                adjust = float(adjust_item.text())
                frl = irl + adjust
                self.tablewidget.setItem(row, 6, QTableWidgetItem(f"{frl:.3f}"))

                # Perform arithmetic check
                first_irl_item = self.tablewidget.item(0, 4)
                last_irl_item = self.tablewidget.item(row_count - 1, 4)
                first_irl = float(first_irl_item.text()) if first_irl_item else 0
                last_irl = float(last_irl_item.text()) if last_irl_item else 0

                arithmetic_check = (sum_bs - sum_fs) == (last_irl - first_irl)
                self.sum_bs_value.setText(f"{sum_bs:.3f}")
                self.sum_fs_value.setText(f"{sum_fs:.3f}")
                self.last_irl_value.setText(f"{last_irl:.3f}")
                self.first_irl_value.setText(f"{first_irl:.3f}")
                self.arithmetic_result_value.setText("Pass" if arithmetic_check else "Fail")

                # Calculate misclose
                misclose = last_irl - final_bm
                self.misclose_value.setText(f"{misclose:.3f}")

                # Calculate allowable misclose
                n = bs_count  # Number of instrument stations (change points)
                allowable_misclose = k_value * math.sqrt(n)
                self.allowable_misclose_value.setText(f"±{allowable_misclose:.3f}")

                # Check if misclose is within acceptable limits
                acceptable_misclose = initial_bm - final_bm
                self.acceptable_misclose_value.setText(
                    "Yes" if -allowable_misclose <= misclose <= allowable_misclose else "No")

        # Display final BM with remark
        self.tablewidget.setItem(row_count - 1, 7, QTableWidgetItem(f"BM({final_bm:.3f})"))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec())
