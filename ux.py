import sys
import jiraWriter3  # Import your original Python script
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem, QHBoxLayout, QMessageBox, QRadioButton, QButtonGroup
import csv
import datetime
import os


class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.selected_file = None
        self.selected_worksheet = None
        self.columns_set_1 = []
        self.columns_set_2 = []
        self.data_preview = []
        self.flag = None

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel to Jira')
        self.setGeometry(100, 100, 800, 500)
        self.showFullScreen()

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # File Selection
        self.file_label = QLabel('Select Excel File:')
        self.file_line_edit = QLineEdit('No file selected')
        self.file_button = QPushButton('Browse...')
        self.file_button.clicked.connect(self.select_file)

        # Worksheet Selection
        self.worksheet_label = QLabel('Select Worksheet:')
        self.worksheet_combo_box = QComboBox()
        self.worksheet_combo_box.currentIndexChanged.connect(self.update_preview)

        # Worksheet Preview
        self.preview_label = QLabel('Worksheet Preview:')
        self.worksheet_preview = QTableWidget()

        # Column Set 1 Selection
        self.column_label_1 = QLabel('Select Columns (Values Wanted):')
        self.column_list_widget_1 = QListWidget()
        self.column_list_widget_1.setSelectionMode(QListWidget.MultiSelection)

        # Column Set 2 Selection
        self.column_label_2 = QLabel('Select Columns (Subtask Due Dates Wanted):')
        self.column_list_widget_2 = QListWidget()
        self.column_list_widget_2.setSelectionMode(QListWidget.MultiSelection)

        # Input Variable
        self.variable_label = QLabel('Input Variable (Comma-separated strings):')
        self.variable_line_edit = QLineEdit()

        #-----------------------------------------------------
        # Preview Table
        self.preview_table = QTableWidget()
        self.preview_table.setHorizontalHeaderLabels(["Column 1", "Column 2", "Column 3"])  # Set column headers

        # Save to CSV Button
        self.save_button = QPushButton('Save to CSV')
        self.save_button.clicked.connect(self.save_to_csv)
        self.save_button.setStyleSheet('background-color: #4CAF50; color: white;')

        # Cancel Button
        self.cancel_button = QPushButton('Cancel')
        self.cancel_button.clicked.connect(self.cancel_operation)
        self.cancel_button.setStyleSheet('background-color: #F44336; color: white;')
        #----------------------------------------------------------

        # Run Button
        self.run_button = QPushButton('Run Program')
        self.run_button.clicked.connect(self.run_program)
        self.run_button.setStyleSheet('background-color: #4CAF50; color: white;')

        #Quit Button
        self.quit_button = QPushButton('Quit')
        self.quit_button.clicked.connect(self.close_app)

        #Conversion/RIEF Choice
        self.choice_label = QLabel('Select your choice:')
        self.radio_conversion = QRadioButton('Conversion')
        self.radio_rief = QRadioButton('RIEF')
        self.radio_conversion.setChecked(True)  # Default selection
        self.radio_conversion.toggled.connect(lambda: self.set_flag("Conversion"))
        self.radio_rief.toggled.connect(lambda: self.set_flag("RIEF"))

        self.button_group = QButtonGroup()
        self.button_group.addButton(self.radio_conversion)
        self.button_group.addButton(self.radio_rief)

        choice_layout = QHBoxLayout()
        choice_layout.addWidget(self.choice_label)
        choice_layout.addWidget(self.radio_conversion)
        choice_layout.addWidget(self.radio_rief)


        layout = QHBoxLayout()
        input_layout = QVBoxLayout()
        csv_preview_layout = QVBoxLayout()
        excel_layout = QVBoxLayout()
        column_select_title_layout = QHBoxLayout()
        column_select_values_layout = QHBoxLayout()


        #----------------
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        csv_preview_layout.addWidget(self.preview_table)
        csv_preview_layout.addLayout(button_layout)
        #----------------


        input_layout.addWidget(self.file_label)
        input_layout.addWidget(self.file_line_edit)
        input_layout.addWidget(self.file_button)

        input_layout.addWidget(self.worksheet_label)
        input_layout.addWidget(self.worksheet_combo_box)

        excel_layout.addWidget(self.preview_label)
        excel_layout.addWidget(self.worksheet_preview)

        column_select_title_layout.addWidget(self.column_label_1)
        column_select_title_layout.addWidget(self.column_label_2)
        input_layout.addLayout(column_select_title_layout)

        column_select_values_layout.addWidget(self.column_list_widget_1)
        column_select_values_layout.addWidget(self.column_list_widget_2)
        input_layout.addLayout(column_select_values_layout)

        hbox_layout = QHBoxLayout()
        hbox_layout.addWidget(self.column_label_1)
        hbox_layout.addWidget(self.column_label_2)
        hbox_layout.addWidget(self.column_list_widget_1)
        hbox_layout.addWidget(self.column_list_widget_2)


        input_layout.addWidget(self.variable_label)
        input_layout.addWidget(self.variable_line_edit)
        input_layout.addLayout(choice_layout)

        input_layout.addWidget(self.run_button)
        input_layout.addWidget(self.quit_button)

        layout.addLayout(input_layout)
        layout.addLayout(excel_layout)
        layout.addLayout(csv_preview_layout)
        central_widget.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.selected_file = file_path
            self.file_line_edit.setText(file_path)
            self.update_worksheet_combo_box(file_path)

    def set_flag(self, flag):
        self.flag = flag

    def update_worksheet_combo_box(self, file_path):
        if not file_path:
            return

        try:
            sheets = pd.read_excel(file_path, sheet_name=None)
            self.worksheet_combo_box.clear()
            self.worksheet_combo_box.addItems(sheets.keys())
            self.worksheet_combo_box.setCurrentIndex(0)
        except Exception as e:
            print(f"Error: {e}")
            self.worksheet_combo_box.clear()
            self.worksheet_combo_box.addItem("Error: Could not load worksheets.")
            self.worksheet_combo_box.setCurrentIndex(0)

    def update_preview(self):
        if not self.selected_file:
            return

        worksheet_name = self.worksheet_combo_box.currentText()
        if not worksheet_name:
            return

        try:
            df = pd.read_excel(self.selected_file, sheet_name=worksheet_name)
            self.worksheet_preview.setColumnCount(len(df.columns))
            self.worksheet_preview.setRowCount(10)  # Show the first 10 rows as a preview

            # Set column headers
            self.worksheet_preview.setHorizontalHeaderLabels(df.columns)

            # Fill the preview with data
            for row_idx, row_data in df.head(10).iterrows():
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value))
                    self.worksheet_preview.setItem(row_idx, col_idx, item)

            self.selected_worksheet = worksheet_name
            self.columns_set_1 = []
            self.columns_set_2 = []

            # Update the column list widgets with available columns
            self.column_list_widget_1.clear()
            self.column_list_widget_1.addItems(df.columns)
            self.column_list_widget_2.clear()
            self.column_list_widget_2.addItems(df.columns)
        except Exception as e:
            print(f"Error: {e}")


    def save_to_csv(self):
            file_path, _ = QFileDialog.getSaveFileName(self, 'Save CSV File', '', 'CSV Files (*.csv)')
            if file_path:
                try:
                    with open(file_path, 'w', newline='') as csv_file:
                        writer = csv.writer(csv_file)
                        writer.writerows(self.data_preview)
                    QMessageBox.information(self, 'Success', 'Data saved to CSV successfully.', QMessageBox.Ok)
                except Exception as e:
                    QMessageBox.warning(self, 'Error', f'Error saving data to CSV: {str(e)}', QMessageBox.Ok)


    def cancel_operation(self):
        self.data_preview = []
        self.preview_table.clearContents()
        self.close()


    def run_program(self):
        if not (self.selected_file and self.selected_worksheet):
            QMessageBox.warning(self, 'Error', 'Please select an Excel file and a worksheet.', QMessageBox.Ok)
            return

        variable_value = self.variable_line_edit.text().split(',')
        columns_set_1 = [item.text() for item in self.column_list_widget_1.selectedItems()]
        columns_set_2 = [item.text() for item in self.column_list_widget_2.selectedItems()]

        if not columns_set_1 or not columns_set_2:
            QMessageBox.warning(self, 'Error', 'Please select columns for both sets.', QMessageBox.Ok)
            return

        self.columns_set_1 = columns_set_1
        self.columns_set_2 = columns_set_2

        if self.flag is None:
            QMessageBox.warning(self, 'Error', 'Please select your choice.', QMessageBox.Ok)
            return

        # Call your_script here and pass the necessary inputs
        self.data_preview = jiraWriter3.generate(self.selected_file, self.selected_worksheet, self.columns_set_1, self.columns_set_2, variable_value, self.flag)

        self.preview_table.setRowCount(len(self.data_preview))
        self.preview_table.setColumnCount(len(self.data_preview[0]))

        for row_idx, row_data in enumerate(self.data_preview):
            for col_idx, cell_value in enumerate(row_data):
                item = QTableWidgetItem(str(cell_value))
                self.preview_table.setItem(row_idx, col_idx, item)
        



    def close_app(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Apply a custom stylesheet to enhance the appearance of the application
    app.setStyleSheet('''
        QMainWindow {
            background-color: #f0f0f0;
        }
        QLabel {
            font-size: 14px;
        }
        QPushButton {
            padding: 10px;
            font-size: 14px;
            border: 2px solid #4CAF50;
            border-radius: 5px;
            min-width: 150px;
        }
        QPushButton:hover {
            background-color: #4CAF50;
            color: white;
        }
        QTableWidget {
            border: 1px solid #ccc;
        }
        QTableWidget QHeaderView::section {
            background-color: #f0f0f0;
            border: none;
            font-size: 12px;
        }
        QListWidget {
            border: 1px solid #ccc;
        }
        QListWidget QAbstractItemView {
            selection-background-color: #4CAF50;
            selection-color: white;
        }
    ''')

    window = AppWindow()
    window.show()
    sys.exit(app.exec_())

