import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAction, 
    QFileDialog, QVBoxLayout, QWidget, QToolBar, QStatusBar, QColorDialog,
    QInputDialog, QMessageBox, QMenu, QHeaderView, QAbstractItemView,
    QDialog, QFormLayout, QLabel, QLineEdit, QPushButton
)
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtCore import Qt, QSize

class FilterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filter Data")
        self.setGeometry(300, 300, 300, 150)
        
        self.layout = QFormLayout()
        self.setLayout(self.layout)
        
        self.column_label = QLabel("Column:")
        self.column_input = QLineEdit()
        self.value_label = QLabel("Value:")
        self.value_input = QLineEdit()
        
        self.layout.addRow(self.column_label, self.column_input)
        self.layout.addRow(self.value_label, self.value_input)
        
        self.button_box = QWidget()
        self.button_layout = QVBoxLayout()
        self.button_box.setLayout(self.button_layout)
        
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")
        
        self.button_layout.addWidget(self.ok_button)
        self.button_layout.addWidget(self.cancel_button)
        
        self.layout.addWidget(self.button_box)
        
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        
    def get_inputs(self):
        return self.column_input.text(), self.value_input.text()


class Grenda(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Grenda - Advanced Excel-like App")
        self.setGeometry(200, 200, 1200, 800)

        # Initialize UI Elements
        self.create_table()
        self.create_toolbar()
        self.create_status_bar()

        # Set the layout
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def create_table(self):
        """Create the main table widget."""
        self.table = QTableWidget(10, 10, self)
        self.table.setFont(QFont("Arial", 12))
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableWidget { gridline-color: black; }")
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)  # Use SelectRows instead of SelectAll
        self.table.setHorizontalHeaderLabels(["Column {}".format(i+1) for i in range(10)])

    def create_toolbar(self):
        """Create a toolbar with actions."""
        toolbar = QToolBar("Main Toolbar")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)

        save_action = QAction(QIcon("save_icon.png"), "Save", self)
        save_action.triggered.connect(self.save_to_csv)
        toolbar.addAction(save_action)

        load_action = QAction(QIcon("load_icon.png"), "Load", self)
        load_action.triggered.connect(self.load_from_csv)
        toolbar.addAction(load_action)

        save_excel_action = QAction(QIcon("save_excel_icon.png"), "Save as Excel", self)
        save_excel_action.triggered.connect(self.save_to_excel)
        toolbar.addAction(save_excel_action)

        load_excel_action = QAction(QIcon("load_excel_icon.png"), "Load Excel", self)
        load_excel_action.triggered.connect(self.load_from_excel)
        toolbar.addAction(load_excel_action)

        add_row_action = QAction(QIcon("add_row_icon.png"), "Add Row", self)
        add_row_action.triggered.connect(self.add_row)
        toolbar.addAction(add_row_action)

        add_column_action = QAction(QIcon("add_column_icon.png"), "Add Column", self)
        add_column_action.triggered.connect(self.add_column)
        toolbar.addAction(add_column_action)

        remove_row_action = QAction(QIcon("remove_row_icon.png"), "Remove Row", self)
        remove_row_action.triggered.connect(self.remove_row)
        toolbar.addAction(remove_row_action)

        remove_column_action = QAction(QIcon("remove_column_icon.png"), "Remove Column", self)
        remove_column_action.triggered.connect(self.remove_column)
        toolbar.addAction(remove_column_action)

        merge_action = QAction(QIcon("merge_icon.png"), "Merge Cells", self)
        merge_action.triggered.connect(self.merge_cells)
        toolbar.addAction(merge_action)

        format_action = QAction(QIcon("format_icon.png"), "Format Cell", self)
        format_action.triggered.connect(self.format_cell)
        toolbar.addAction(format_action)

        formula_action = QAction(QIcon("formula_icon.png"), "Insert Formula", self)
        formula_action.triggered.connect(self.insert_formula)
        toolbar.addAction(formula_action)

        filter_action = QAction(QIcon("filter_icon.png"), "Filter Data", self)
        filter_action.triggered.connect(self.show_filter_dialog)
        toolbar.addAction(filter_action)

        undo_action = QAction(QIcon("undo_icon.png"), "Undo", self)
        undo_action.triggered.connect(self.undo_last_action)
        toolbar.addAction(undo_action)

    def create_status_bar(self):
        """Create a status bar at the bottom."""
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

    def save_to_csv(self):
        """Save the content of the table to a CSV file."""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'w') as file:
                for row in range(self.table.rowCount()):
                    row_data = []
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        if item is not None:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    file.write(','.join(row_data) + '\n')
            self.statusBar.showMessage("File saved successfully", 2000)

    def load_from_csv(self):
        """Load content from a CSV file into the table."""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'r') as file:
                content = file.readlines()
                self.table.setRowCount(len(content))
                for row, line in enumerate(content):
                    columns = line.strip().split(',')
                    self.table.setColumnCount(len(columns))
                    for column, data in enumerate(columns):
                        self.table.setItem(row, column, QTableWidgetItem(data))
            self.statusBar.showMessage("File loaded successfully", 2000)

    def save_to_excel(self):
        """Save the content of the table to an Excel file."""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for column in range(self.table.columnCount()):
                    item = self.table.item(row, column)
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append('')
                data.append(row_data)
            df = pd.DataFrame(data)
            df.to_excel(file_name, index=False, header=False)
            self.statusBar.showMessage("File saved successfully as Excel", 2000)

    def load_from_excel(self):
        """Load content from an Excel file into the table."""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            df = pd.read_excel(file_name, header=None)
            self.table.setRowCount(df.shape[0])
            self.table.setColumnCount(df.shape[1])
            for row in range(df.shape[0]):
                for column in range(df.shape[1]):
                    self.table.setItem(row, column, QTableWidgetItem(str(df.iat[row, column])))
            self.statusBar.showMessage("File loaded successfully as Excel", 2000)

    def add_row(self):
        """Add a new row to the table."""
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        self.statusBar.showMessage("Row added", 2000)

    def add_column(self):
        """Add a new column to the table."""
        column_count = self.table.columnCount()
        self.table.insertColumn(column_count)
        self.statusBar.showMessage("Column added", 2000)

    def remove_row(self):
        """Remove the selected row from the table."""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)
            self.statusBar.showMessage("Row removed", 2000)

    def remove_column(self):
        """Remove the selected column from the table."""
        current_column = self.table.currentColumn()
        if current_column >= 0:
            self.table.removeColumn(current_column)
            self.statusBar.showMessage("Column removed", 2000)

    def merge_cells(self):
        """Merge the selected cells."""
        selected_range = self.table.selectedRanges()
        if len(selected_range) > 0:
            top_row = selected_range[0].topRow()
            bottom_row = selected_range[0].bottomRow()
            left_column = selected_range[0].leftColumn()
            right_column = selected_range[0].rightColumn()
            self.table.setSpan(top_row, left_column, bottom_row - top_row + 1, right_column - left_column + 1)
            self.statusBar.showMessage("Cells merged", 2000)

    def format_cell(self):
        """Format the selected cell."""
        selected_item = self.table.currentItem()
        if selected_item:
            color = QColorDialog.getColor()
            if color.isValid():
                selected_item.setBackground(color)
                self.statusBar.showMessage("Cell formatted", 2000)

    def insert_formula(self):
        """Insert a formula into the selected cell."""
        selected_item = self.table.currentItem()
        if selected_item:
            formula, ok = QInputDialog.getText(self, "Insert Formula", "Enter the formula:")
            if ok and formula:
                try:
                    # Basic formula evaluation (e.g., "=5+3")
                    if formula.startswith("="):
                        result = eval(formula[1:])
                        selected_item.setText(str(result))
                        self.statusBar.showMessage("Formula inserted", 2000)
                    else:
                        QMessageBox.warning(self, "Error", "Formula must start with '='")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Invalid formula: {e}")

    def show_filter_dialog(self):
        """Show dialog to filter data."""
        dialog = FilterDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            column_name, filter_value = dialog.get_inputs()
            column_index = self.table.horizontalHeader().visualIndexAt(self.table.columnViewportPosition(column_name))
            if column_index >= 0:
                self.filter_data(column_index, filter_value)

    def filter_data(self, column, value):
        """Filter table data based on column and value."""
        for row in range(self.table.rowCount()):
            item = self.table.item(row, column)
            if item and value not in item.text():
                self.table.setRowHidden(row, True)
            else:
                self.table.setRowHidden(row, False)

    def undo_last_action(self):
        """Undo the last action (requires implementing an undo stack)."""
        QMessageBox.information(self, "Undo", "Undo feature not yet implemented.")

    def clear_table(self):
        """Clear all content in the table."""
        self.table.clearContents()
        self.statusBar.showMessage("Table cleared", 2000)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = Grenda()
    window.show()
    sys.exit(app.exec_())
