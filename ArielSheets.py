import sys
import json
import re
import numpy as np
import logging
from datetime import datetime
from sympy import sympify, SympifyError
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTableWidget, 
                            QTableWidgetItem, QVBoxLayout, QWidget, 
                            QMenuBar, QMenu, QFileDialog, QToolBar,
                            QFontComboBox, QSpinBox, QPushButton,
                            QColorDialog, QLineEdit, QHBoxLayout,
                            QInputDialog, QMessageBox, QDialog, 
                            QLabel, QDialogButtonBox, QComboBox,
                            QHeaderView, QTabWidget)
from PyQt6.QtGui import QFont, QKeySequence, QColor, QAction, QShortcut, QIcon
from PyQt6.QtCore import Qt, QRegularExpression, QSize
import requests
import webbrowser
from packaging import version
import os
from pathlib import Path

# Create logs directory in AppData
app_data_path = os.path.join(os.getenv('APPDATA'), 'Ariel Sheets')
logs_path = os.path.join(app_data_path, 'logs')
os.makedirs(logs_path, exist_ok=True)

# Set up logging with the new path
logging.basicConfig(
    filename=os.path.join(logs_path, f'pysheet_{datetime.now().strftime("%Y%m%d")}.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class TableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Table")
        layout = QVBoxLayout(self)

        # Rows input
        layout.addWidget(QLabel("Number of rows:"))
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 100)
        self.rows_spin.setValue(5)
        layout.addWidget(self.rows_spin)

        # Columns input
        layout.addWidget(QLabel("Number of columns:"))
        self.cols_spin = QSpinBox()
        self.cols_spin.setRange(1, 26)
        self.cols_spin.setValue(5)
        layout.addWidget(self.cols_spin)

        # Style selection
        layout.addWidget(QLabel("Table Style:"))
        self.style_combo = QComboBox()
        self.style_combo.addItems(["Simple", "Striped", "Professional"])
        layout.addWidget(self.style_combo)

        # Dialog buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

class FormulaBar(QLineEdit):
    def __init__(self, spreadsheet):
        super().__init__()
        self.spreadsheet = spreadsheet
        self.setPlaceholderText("Enter formula or value")

class Sheet(QTableWidget):
    """Individual sheet class that inherits from QTableWidget"""
    def __init__(self, rows=50, cols=26, parent=None):
        super().__init__(rows, cols, parent)
        self.cell_validations = {}  # Initialize validations dict
        self.tables = []  # Initialize tables list
        self.setup_sheet()
        
    def setup_sheet(self):
        # Set headers
        headers = [chr(65 + i) for i in range(self.columnCount())]
        self.setHorizontalHeaderLabels(headers)
        self.setVerticalHeaderLabels([str(i+1) for i in range(self.rowCount())])
        
        # Enable selection and copying
        self.setSelectionMode(QTableWidget.SelectionMode.ContiguousSelection)

class ExcelClone(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ariel Sheets")
        self.resize(1000, 800)
        self.setWindowIcon(QIcon("icon.ico"))
        
        # Apply stylesheet
        self.setStyleSheet(Style.get_stylesheet())
        
        # Initialize variables
        self.clipboard = None
        self.sort_order = Qt.SortOrder.AscendingOrder
        self.cell_validations = {}
        self.current_file_path = None
        
        # Update table styles with new theme colors
        self.table_styles = {
            "Simple": {
                "header": QColor(Style.SECONDARY),
                "cells": QColor(Style.BACKGROUND),
                "border": True
            },
            "Striped": {
                "header": QColor(Style.PRIMARY),
                "cells": [QColor(Style.BACKGROUND), QColor(Style.ACCENT)],
                "border": True
            },
            "Professional": {
                "header": QColor(Style.PRIMARY),
                "cells": QColor(Style.BACKGROUND),
                "border": True
            }
        }
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Create formula bar
        self.formula_bar = FormulaBar(self)
        layout.addWidget(self.formula_bar)
        
        # Create tab widget for sheets
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_sheet)
        layout.addWidget(self.tab_widget)
        
        # Add first sheet
        self.add_sheet()
        
        # Create menu bar and toolbars
        self.create_menu()
        self.create_format_toolbar()
        
        # Connect formula bar to current sheet
        self.formula_bar.returnPressed.connect(self.formula_entered)
        self.tab_widget.currentChanged.connect(self.sheet_changed)
        
        # Check for updates on startup
        self.check_updates()

    def add_sheet(self):
        """Add a new sheet to the workbook"""
        sheet_count = self.tab_widget.count()
        new_sheet = Sheet()
        new_sheet.itemChanged.connect(self.cell_changed)
        new_sheet.cell_validations = {}  # Initialize validations for the new sheet
        
        # Get sheet name from user
        while True:
            name, ok = QInputDialog.getText(self, 'New Sheet', 
                                          'Enter sheet name:',
                                          text=f'Sheet{sheet_count + 1}')
            if not ok:
                if sheet_count == 0:  # Must have at least one sheet
                    continue
                return
            
            # Check if name is unique
            if not any(name == self.tab_widget.tabText(i) 
                      for i in range(self.tab_widget.count())):
                break
            QMessageBox.warning(self, 'Error', 'Sheet name must be unique')
        
        self.tab_widget.addTab(new_sheet, name)
        self.tab_widget.setCurrentWidget(new_sheet)

    def close_sheet(self, index):
        """Close a sheet tab"""
        if self.tab_widget.count() <= 1:
            QMessageBox.warning(self, 'Error', 
                              'Cannot close last sheet')
            return
        
        self.tab_widget.removeTab(index)

    def rename_sheet(self):
        """Rename the current sheet"""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return
            
        current_name = self.tab_widget.tabText(current_index)
        while True:
            name, ok = QInputDialog.getText(self, 'Rename Sheet', 
                                          'Enter new name:',
                                          text=current_name)
            if not ok:
                return
                
            # Check if name is unique
            if not any(name == self.tab_widget.tabText(i) 
                      for i in range(self.tab_widget.count())):
                break
            QMessageBox.warning(self, 'Error', 'Sheet name must be unique')
        
        self.tab_widget.setTabText(current_index, name)

    def sheet_changed(self, index):
        """Handle sheet tab changes"""
        if index >= 0:
            current_sheet = self.tab_widget.widget(index)
            current_item = current_sheet.currentItem()
            if current_item:
                self.formula_bar.setText(current_item.text())
            else:
                self.formula_bar.clear()

    @property
    def current_sheet(self):
        """Get the currently active sheet"""
        return self.tab_widget.currentWidget()

    def create_menu(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        new_action = QAction("New", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_file)
        file_menu.addAction(new_action)
        
        open_action = QAction("Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        
        save_action = QAction("Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save As", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_as_file)
        file_menu.addAction(save_as_action)
        
        # Edit menu
        edit_menu = menubar.addMenu("Edit")
        
        copy_action = QAction("Copy", self)
        copy_action.setShortcut(QKeySequence.StandardKey.Copy)
        copy_action.triggered.connect(self.copy_cells)
        edit_menu.addAction(copy_action)

        paste_action = QAction("Paste", self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(self.paste_cells)
        edit_menu.addAction(paste_action)

        cut_action = QAction("Cut", self)
        cut_action.setShortcut(QKeySequence.StandardKey.Cut)
        cut_action.triggered.connect(self.cut_cells)
        edit_menu.addAction(cut_action)

        # Table menu
        table_menu = menubar.addMenu("Table")
        
        insert_table_action = QAction("Insert Table", self)
        insert_table_action.triggered.connect(self.insert_table)
        table_menu.addAction(insert_table_action)

        sort_table_action = QAction("Sort Table", self)
        sort_table_action.triggered.connect(self.sort_table)
        table_menu.addAction(sort_table_action)

        # Data menu
        data_menu = menubar.addMenu("Data")
        
        validation_action = QAction("Data Validation", self)
        validation_action.triggered.connect(self.add_data_validation)
        data_menu.addAction(validation_action)

        # Help menu
        help_menu = menubar.addMenu("Help")
        
        check_update_action = QAction("Check for Updates", self)
        check_update_action.triggered.connect(self.check_updates)
        help_menu.addAction(check_update_action)

        # Sheet menu
        sheet_menu = menubar.addMenu("Sheet")
        
        new_sheet_action = QAction("New Sheet", self)
        new_sheet_action.setShortcut("Ctrl+Shift+N")
        new_sheet_action.triggered.connect(self.add_sheet)
        sheet_menu.addAction(new_sheet_action)
        
        rename_sheet_action = QAction("Rename Sheet", self)
        rename_sheet_action.triggered.connect(self.rename_sheet)
        sheet_menu.addAction(rename_sheet_action)

    def create_format_toolbar(self):
        format_toolbar = QToolBar()
        format_toolbar.setIconSize(QSize(16, 16))  # Smaller icons
        self.addToolBar(format_toolbar)

        # Font family
        self.font_combo = QFontComboBox()
        self.font_combo.setFixedWidth(150)  # Set fixed width
        self.font_combo.currentFontChanged.connect(self.change_font)
        format_toolbar.addWidget(self.font_combo)

        format_toolbar.addSeparator()

        # Font size
        self.font_size = QSpinBox()
        self.font_size.setRange(6, 72)
        self.font_size.setValue(11)
        self.font_size.setFixedWidth(50)  # Set fixed width
        self.font_size.valueChanged.connect(self.change_font_size)
        format_toolbar.addWidget(self.font_size)

        format_toolbar.addSeparator()

        # Bold
        bold_btn = QPushButton("B")
        bold_font = QFont("Arial", 10)
        bold_font.setBold(True)
        bold_btn.setFont(bold_font)
        bold_btn.setFixedSize(30, 30)  # Square button
        bold_btn.clicked.connect(self.format_bold)
        format_toolbar.addWidget(bold_btn)

        # Italic
        italic_btn = QPushButton("I")
        italic_font = QFont("Arial", 10)
        italic_font.setItalic(True)
        italic_btn.setFont(italic_font)
        italic_btn.setFixedSize(30, 30)  # Square button
        italic_btn.clicked.connect(self.format_italic)
        format_toolbar.addWidget(italic_btn)

        format_toolbar.addSeparator()

        # Cell color
        color_btn = QPushButton("Fill")
        color_btn.setFixedWidth(50)
        color_btn.clicked.connect(self.change_cell_color)
        format_toolbar.addWidget(color_btn)

    def change_font(self, font):
        if self.current_sheet:
            for item in self.current_sheet.selectedItems():
                item.setFont(font)

    def change_font_size(self, size):
        if self.current_sheet:
            for item in self.current_sheet.selectedItems():
                font = item.font()
                font.setPointSize(size)
                item.setFont(font)

    def format_bold(self):
        if self.current_sheet:
            for item in self.current_sheet.selectedItems():
                font = item.font()
                font.setBold(not font.bold())
                item.setFont(font)

    def format_italic(self):
        if self.current_sheet:
            for item in self.current_sheet.selectedItems():
                font = item.font()
                font.setItalic(not font.italic())
                item.setFont(font)

    def change_cell_color(self):
        if self.current_sheet:
            color = QColorDialog.getColor()
            if color.isValid():
                for item in self.current_sheet.selectedItems():
                    item.setBackground(color)

    def copy_cells(self):
        self.clipboard = []
        for item in self.spreadsheet.selectedItems():
            self.clipboard.append({
                'row': item.row(),
                'col': item.column(),
                'text': item.text(),
                'font': item.font(),
                'background': item.background()
            })

    def paste_cells(self):
        if not self.clipboard:
            return
            
        selected_items = self.spreadsheet.selectedItems()
        if not selected_items:
            current_row = self.spreadsheet.currentRow()
            current_col = self.spreadsheet.currentColumn()
            if current_row < 0 or current_col < 0:
                return
            base_row = current_row
            base_col = current_col
        else:
            top_left = selected_items[0]
            base_row = top_left.row()
            base_col = top_left.column()

        for cell in self.clipboard:
            row_offset = cell['row'] - self.clipboard[0]['row']
            col_offset = cell['col'] - self.clipboard[0]['col']
            new_row = base_row + row_offset
            new_col = base_col + col_offset
            
            if (new_row >= 0 and new_row < self.spreadsheet.rowCount() and 
                new_col >= 0 and new_col < self.spreadsheet.columnCount()):
                new_item = QTableWidgetItem(cell['text'])
                new_item.setFont(cell['font'])
                new_item.setBackground(cell['background'])
                self.spreadsheet.setItem(new_row, new_col, new_item)

    def cut_cells(self):
        self.copy_cells()
        for item in self.spreadsheet.selectedItems():
            item.setText("")
    def insert_table(self):
        dialog = TableDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            rows = dialog.rows_spin.value()
            cols = dialog.cols_spin.value()
            style = dialog.style_combo.currentText()
            
            start_row = self.current_sheet.currentRow()
            start_col = self.current_sheet.currentColumn()
            
            if start_row < 0 or start_col < 0:
                QMessageBox.warning(self, "Warning", "Please select a starting cell")
                return

            self.create_table(start_row, start_col, rows, cols, style)

    def create_table(self, start_row, start_col, rows, cols, style):
        style_dict = self.table_styles[style]
        
        # Create header row
        for col in range(cols):
            header_item = QTableWidgetItem(f"Header {col + 1}")
            header_item.setBackground(style_dict["header"])
            header_item.setForeground(QColor("white") if style == "Professional" else QColor("black"))
            font = header_item.font()
            font.setBold(True)
            header_item.setFont(font)
            self.current_sheet.setItem(start_row, start_col + col, header_item)

        # Create data cells
        for row in range(rows - 1):
            for col in range(cols):
                item = QTableWidgetItem(f"Data")
                if style == "Striped":
                    item.setBackground(style_dict["cells"][row % 2])
                else:
                    item.setBackground(style_dict["cells"])
                self.current_sheet.setItem(start_row + row + 1, start_col + col, item)

        # Store table information
        if not hasattr(self.current_sheet, 'tables'):
            self.current_sheet.tables = []
        
        self.current_sheet.tables.append({
            'start_row': start_row,
            'start_col': start_col,
            'rows': rows,
            'cols': cols
        })

    def sort_table(self):
        current_row = self.current_sheet.currentRow()
        current_col = self.current_sheet.currentColumn()
        
        if not hasattr(self.current_sheet, 'tables'):
            self.current_sheet.tables = []
        
        table = None
        for t in self.current_sheet.tables:
            if (t['start_row'] <= current_row < t['start_row'] + t['rows'] and
                t['start_col'] <= current_col < t['start_col'] + t['cols']):
                table = t
                break
        
        if not table:
            QMessageBox.warning(self, "Warning", "Please select a cell within a table")
            return

        col_to_sort = current_col - table['start_col']
        
        # Collect data from the table
        data = []
        for row in range(table['rows'] - 1):  # Exclude header row
            row_data = []
            for col in range(table['cols']):
                item = self.current_sheet.item(table['start_row'] + row + 1, 
                                           table['start_col'] + col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        # Sort data based on the selected column
        data.sort(key=lambda x: x[col_to_sort])

        # Update the table with sorted data
        for row, row_data in enumerate(data):
            for col, cell_data in enumerate(row_data):
                item = QTableWidgetItem(cell_data)
                self.current_sheet.setItem(table['start_row'] + row + 1, 
                                       table['start_col'] + col, item)

    def get_sort_key(self, value):
        """Handle different types of data for sorting"""
        try:
            return float(value)
        except ValueError:
            return str(value).lower()

    def update_table_data(self, table, data):
        for row, row_data in enumerate(data):
            for col, cell_data in enumerate(row_data):
                item = QTableWidgetItem(cell_data)
                self.spreadsheet.setItem(table['start_row'] + row + 1, 
                                       table['start_col'] + col, item)

    def find_table(self, current_row, current_col):
        for t in self.tables:
            if (t['start_row'] <= current_row < t['start_row'] + t['rows'] and
                t['start_col'] <= current_col < t['start_col'] + t['cols']):
                return t
        return None

    def get_table_data(self, table):
        data = []
        for row in range(table['rows'] - 1):  # Exclude header row
            row_data = []
            for col in range(table['cols']):
                item = self.spreadsheet.item(table['start_row'] + row + 1, 
                                          table['start_col'] + col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        return data

    def add_data_validation(self):
        dialog = DataValidationDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            validation_type = dialog.type_combo.currentText()
            
            for item in self.spreadsheet.selectedItems():
                cell_id = self.get_cell_id(item.row(), item.column())
                validation = {
                    'type': validation_type,
                    'min': dialog.min_value.text() if validation_type == "Custom Range" else None,
                    'max': dialog.max_value.text() if validation_type == "Custom Range" else None
                }
                self.cell_validations[cell_id] = validation

    def validate_cell_input(self, item):
        """Validate cell input based on sheet-specific validation rules"""
        if not item:
            return True
        
        sheet = self.tab_widget.currentWidget()
        if not sheet:
            return True
        
        if not hasattr(sheet, 'cell_validations'):
            sheet.cell_validations = {}
        
        cell_id = self.get_cell_id(item.row(), item.column())
        if cell_id not in sheet.cell_validations:
            return True

        validation = sheet.cell_validations[cell_id]
        value = item.text()

        if validation['type'] == "Number Only":
            try:
                float(value)
                return True
            except ValueError:
                return False
        elif validation['type'] == "Text Only":
            try:
                float(value)
                return False
            except ValueError:
                return True
        elif validation['type'] == "Custom Range":
            try:
                num_value = float(value)
                min_val = float(validation['min']) if validation['min'] else float('-inf')
                max_val = float(validation['max']) if validation['max'] else float('inf')
                return min_val <= num_value <= max_val
            except ValueError:
                return False

        return True

    def cell_changed(self, item):
        """Handle cell content changes"""
        if not item:
            return
        
        sheet = self.tab_widget.currentWidget()
        if not sheet:
            return

        if not self.validate_cell_input(item):
            QMessageBox.warning(self, "Invalid Input", 
                              "The entered value does not meet the validation criteria.")
            item.setText("")
            return

        text = item.text()
        if text.startswith('='):
            result = self.evaluate_formula(text)
            # Temporarily disconnect to prevent recursive signal
            sheet.itemChanged.disconnect(self.cell_changed)
            item.setText(result)
            sheet.itemChanged.connect(self.cell_changed)

    def new_file(self):
        """Create a new spreadsheet, prompting to save if there are unsaved changes"""
        if self.has_unsaved_changes():
            reply = QMessageBox.question(
                self, 'Ariel Sheets - Save Changes?',
                'Do you want to save your changes before creating a new file?',
                QMessageBox.StandardButton.Save | 
                QMessageBox.StandardButton.Discard | 
                QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Save
            )

            if reply == QMessageBox.StandardButton.Save:
                self.save_file()
            elif reply == QMessageBox.StandardButton.Cancel:
                return  # Don't create new file if user cancels

        # Disconnect the itemChanged signal temporarily to prevent triggering while clearing
        self.spreadsheet.itemChanged.disconnect(self.cell_changed)
        
        # Clear all content
        self.spreadsheet.clear()
        
        # Reset the spreadsheet
        self.setup_spreadsheet()
        
        # Clear the tables list
        self.tables.clear()
        
        # Clear the clipboard
        self.clipboard = None
        
        # Clear the formula bar
        self.formula_bar.clear()
        
        # Clear validations
        self.cell_validations.clear()
        
        # Reconnect the itemChanged signal
        self.spreadsheet.itemChanged.connect(self.cell_changed)

    def save_file(self):
        if hasattr(self, 'current_file_path') and self.current_file_path:
            filename = self.current_file_path
        else:
            filename, _ = QFileDialog.getSaveFileName(
                self, "Save Spreadsheet", "", 
                "Ariel Sheets Files (*.xlas);;All Files (*)"
            )
            
            if filename and not filename.endswith('.xlas'):
                filename += '.xlas'
        
        if filename:
            try:
                data = {
                    'sheets': {}
                }
                
                # Save each sheet
                for sheet_index in range(self.tab_widget.count()):
                    sheet = self.tab_widget.widget(sheet_index)
                    sheet_name = self.tab_widget.tabText(sheet_index)
                    sheet_data = {
                        'cells': {},
                        'tables': getattr(sheet, 'tables', []),
                        'validations': getattr(sheet, 'cell_validations', {})
                    }
                    
                    # Save cell contents and formatting
                    for row in range(sheet.rowCount()):
                        for col in range(sheet.columnCount()):
                            item = sheet.item(row, col)
                            if item and (item.text() or item.background().color().isValid()):
                                cell_id = self.get_cell_id(row, col)
                                sheet_data['cells'][cell_id] = {
                                    'text': item.text(),
                                    'background': item.background().color().name(),
                                    'foreground': item.foreground().color().name(),  # Add foreground color
                                    'font_family': item.font().family(),
                                    'font_size': item.font().pointSize(),
                                    'font_bold': item.font().bold(),
                                    'font_italic': item.font().italic()
                                }
                
                data['sheets'][sheet_name] = sheet_data
            
                with open(filename, 'w') as f:
                    json.dump(data, f)
                    
                self.current_file_path = filename  # Store the file path
                logging.info(f"File saved successfully: {filename}")
                
            except Exception as e:
                logging.error(f"Error saving file: {str(e)}")
                QMessageBox.warning(self, "Error", f"Could not save file: {str(e)}")

    def open_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open Spreadsheet", "", 
            "Ariel Sheets Files (*.xlas);;All Files (*)"
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    data = json.load(f)
                
                # Clear current sheets
                while self.tab_widget.count() > 0:
                    self.tab_widget.removeTab(0)
                
                # Load sheets
                for sheet_name, sheet_data in data['sheets'].items():
                    new_sheet = Sheet()
                    new_sheet.itemChanged.connect(self.cell_changed)
                    
                    # Restore cells
                    for cell_id, cell_data in sheet_data['cells'].items():
                        col = ord(cell_id[0]) - 65
                        row = int(cell_id[1:]) - 1
                        
                        item = QTableWidgetItem(cell_data['text'])
                        
                        # Restore formatting
                        font = QFont(cell_data['font_family'], 
                                   cell_data['font_size'])
                        font.setBold(cell_data['font_bold'])
                        font.setItalic(cell_data['font_italic'])
                        item.setFont(font)
                        
                        # Set background color
                        item.setBackground(QColor(cell_data['background']))
                        
                        # Set text color (foreground)
                        if 'foreground' in cell_data:
                            item.setForeground(QColor(cell_data['foreground']))
                        else:
                            item.setForeground(QColor("black"))  # Default text color
                        
                        new_sheet.setItem(row, col, item)
                    
                    # Restore tables and validations
                    new_sheet.tables = sheet_data.get('tables', [])
                    new_sheet.cell_validations = sheet_data.get('validations', {})
                    
                    self.tab_widget.addTab(new_sheet, sheet_name)
                
                # Select first sheet
                if self.tab_widget.count() > 0:
                    self.tab_widget.setCurrentIndex(0)
                
                self.current_file_path = filename  # Store the opened file path
                
            except Exception as e:
                logging.error(f"Error opening file: {str(e)}")
                QMessageBox.warning(self, "Error", f"Could not open file: {str(e)}")

    def evaluate_formula(self, formula):
        if not formula.startswith('='):
            return formula

        formula = formula[1:]  # Remove equals sign
        
        # Handle special functions
        if formula.startswith(('SUM(', 'AVERAGE(', 'MIN(', 'MAX(', 'COUNT(')):
            return self.handle_special_function(formula)
            
        # Replace cell references with values
        formula = self.replace_cell_references(formula)
        if formula.startswith('ERROR'):
            return formula
        
        # Safely evaluate the formula
        try:
            # Only allow safe operations
            allowed_chars = set('0123456789+-*/() .')
            if not all(c in allowed_chars for c in formula):
                return "ERROR: Invalid characters in formula"
            result = eval(formula)
            return str(result)
        except Exception as e:
            return f"ERROR: Invalid formula ({str(e)})"

    def handle_special_function(self, formula):
        function_match = re.match(r'([A-Z]+)\((.*)\)', formula)
        if not function_match:
            return "ERROR: Invalid function format"
            
        func_name, range_str = function_match.groups()
        try:
            cells = self.get_cells_in_range(range_str)
            values = [float(cell.text() or 0) for cell in cells]
            
            if not values:
                return "0"
                
            if func_name == "SUM":
                return str(sum(values))
            elif func_name == "AVERAGE":
                return str(sum(values) / len(values))
            elif func_name == "MIN":
                return str(min(values))
            elif func_name == "MAX":
                return str(max(values))
            elif func_name == "COUNT":
                return str(len([v for v in values if v != 0]))
                
        except Exception as e:
            logging.error(f"Function evaluation error: {str(e)}")
            return f"ERROR: Invalid {func_name} range"

    def get_cells_in_range(self, range_str):
        start, end = range_str.split(':')
        start_item = self.get_cell_from_id(start)
        end_item = self.get_cell_from_id(end)
        
        start_row = start_item.row()
        start_col = start_item.column()
        end_row = end_item.row()
        end_col = end_item.column()
        
        cells = []
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                item = self.spreadsheet.item(row, col)
                if item is None:
                    item = QTableWidgetItem("0")
                    self.spreadsheet.setItem(row, col, item)
                cells.append(item)
        return cells

    def formula_entered(self):
        current_item = self.spreadsheet.currentItem()
        if current_item:
            current_item.setText(self.formula_bar.text())

    def update_formula_bar(self, current, previous):
        if current:
            self.formula_bar.setText(current.text())

    def get_cell_id(self, row, col):
        return f"{chr(65 + col)}{row + 1}"

    def get_cell_from_id(self, cell_id):
        match = re.match(r"([A-Z])(\d+)", cell_id)
        if match:
            col = ord(match.group(1)) - 65
            row = int(match.group(2)) - 1
            item = self.spreadsheet.item(row, col)
            if item is None:
                item = QTableWidgetItem("")
                self.spreadsheet.setItem(row, col, item)
            return item
        return None

    def replace_cell_references(self, formula):
        """Replace cell references with their values"""
        cell_pattern = r'[A-Z]\d+'
        cell_references = re.findall(cell_pattern, formula)
        
        for cell_ref in cell_references:
            cell_item = self.get_cell_from_id(cell_ref)
            cell_value = cell_item.text() if cell_item else "0"
            try:
                cell_value = float(cell_value)
            except ValueError:
                return "ERROR: Invalid cell reference"
            formula = formula.replace(cell_ref, str(cell_value))
        
        return formula

    def has_unsaved_changes(self):
        """Check if there are any unsaved changes in any sheet"""
        for sheet_index in range(self.tab_widget.count()):
            sheet = self.tab_widget.widget(sheet_index)
            # Check if any cell in the sheet has content
            for row in range(sheet.rowCount()):
                for col in range(sheet.columnCount()):
                    item = sheet.item(row, col)
                    if item and item.text():
                        return True
        return False

    def closeEvent(self, event):
        """Handle the window close event"""
        if not self.has_unsaved_changes():
            event.accept()
            return

        reply = QMessageBox.question(
            self, 'Ariel Sheets - Save Changes?',
            'Do you want to save your changes before closing?',
            QMessageBox.StandardButton.Save | 
            QMessageBox.StandardButton.Discard | 
            QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save
        )

        if reply == QMessageBox.StandardButton.Save:
            self.save_file()
            event.accept()
        elif reply == QMessageBox.StandardButton.Discard:
            event.accept()
        else:
            event.ignore()

    def check_updates(self):
        checker = UpdateChecker()
        update_available, latest_version, download_url, changelog = checker.check_for_updates()
        
        if update_available:
            msg = QMessageBox()
            msg.setWindowTitle("Update Available")
            msg.setWindowIcon(self.windowIcon())
            msg.setIcon(QMessageBox.Icon.Information)
            
            update_text = (f"A new version of Ariel Sheets is available!\n\n"
                          f"Current version: {checker.current_version}\n"
                          f"Latest version: {latest_version}\n\n"
                          f"Changelog:\n{changelog}\n\n"
                          f"Would you like to download the update?")
            
            msg.setText(update_text)
            msg.setStandardButtons(
                QMessageBox.StandardButton.Yes | 
                QMessageBox.StandardButton.No
            )
            
            if msg.exec() == QMessageBox.StandardButton.Yes:
                webbrowser.open(download_url)

    def save_as_file(self):
        """Save the spreadsheet to a new file"""
        # Clear current file path to force Save As dialog
        self.current_file_path = None
        self.save_file()

class DataValidationDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data Validation")
        layout = QVBoxLayout(self)

        # Validation type
        layout.addWidget(QLabel("Validation Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Any Value", "Number Only", "Text Only", "Custom Range"])
        layout.addWidget(self.type_combo)

        # Range inputs (for custom range)
        self.range_widget = QWidget()
        range_layout = QHBoxLayout(self.range_widget)
        self.min_value = QLineEdit()
        self.max_value = QLineEdit()
        range_layout.addWidget(QLabel("Min:"))
        range_layout.addWidget(self.min_value)
        range_layout.addWidget(QLabel("Max:"))
        range_layout.addWidget(self.max_value)
        layout.addWidget(self.range_widget)
        self.range_widget.hide()

        # Connect type combo to show/hide range inputs
        self.type_combo.currentTextChanged.connect(self.on_type_changed)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def on_type_changed(self, text):
        self.range_widget.setVisible(text == "Custom Range")

class UpdateChecker:
    def __init__(self):
        self.current_version = "1.0.0"  # Your current version
        self.update_url = "https://raw.githubusercontent.com/YourUsername/ArielSheets/main/version.json"
        # Backup URL in case GitHub is down
        self.backup_url = "https://your-backup-domain.com/version.json"

    def check_for_updates(self):
        try:
            # Try primary URL
            response = requests.get(self.update_url, timeout=5)
            if response.status_code != 200:
                # Try backup URL if primary fails
                response = requests.get(self.backup_url, timeout=5)
            
            if response.status_code == 200:
                update_data = response.json()
                latest_version = update_data.get('version')
                download_url = update_data.get('download_url')
                changelog = update_data.get('changelog', '')

                if version.parse(latest_version) > version.parse(self.current_version):
                    return True, latest_version, download_url, changelog
                
            return False, None, None, None

        except Exception as e:
            logging.error(f"Update check failed: {str(e)}")
            return False, None, None, None

class Style:
    # Color scheme
    PRIMARY = "#2E7D32"  # Dark green
    SECONDARY = "#4CAF50"  # Medium green
    ACCENT = "#81C784"  # Light green
    BACKGROUND = "#FFFFFF"  # White
    SURFACE = "#F5F5F5"  # Light gray
    TEXT = "#212121"  # Almost black
    TEXT_SECONDARY = "#757575"  # Gray
    
    @staticmethod
    def get_stylesheet():
        return """
        QMainWindow {
            background-color: #F5F5F5;
        }
        
        QMenuBar {
            background-color: #2E7D32;
            color: white;
            padding: 4px;
        }
        
        QMenuBar::item:selected {
            background-color: #4CAF50;
        }
        
        QMenu {
            background-color: #FFFFFF;
            border: 1px solid #CCCCCC;
        }
        
        QMenu::item:selected {
            background-color: #81C784;
        }
        
        QToolBar {
            background-color: #FFFFFF;
            border-bottom: 1px solid #CCCCCC;
            padding: 2px;
        }
        
        QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 3px;
        }
        
        QPushButton:hover {
            background-color: #2E7D32;
        }
        
        QPushButton:pressed {
            background-color: #1B5E20;
        }
        
        QTabWidget::pane {
            border: 1px solid #CCCCCC;
            background-color: #FFFFFF;
        }
        
        QTabBar::tab {
            background-color: #F5F5F5;
            color: #212121;
            padding: 8px 16px;
            border: 1px solid #CCCCCC;
            border-bottom: none;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        
        QTabBar::tab:selected {
            background-color: #4CAF50;
            color: white;
        }
        
        QTableWidget {
            background-color: #FFFFFF;
            gridline-color: #E0E0E0;
        }
        
        QHeaderView::section {
            background-color: #4CAF50;
            color: white;
            padding: 5px;
            border: none;
        }
        
        QLineEdit {
            padding: 5px;
            border: 1px solid #CCCCCC;
            border-radius: 3px;
            background-color: #FFFFFF;
        }
        
        QSpinBox, QFontComboBox {
            padding: 4px;
            border: 1px solid #CCCCCC;
            border-radius: 3px;
            background-color: #FFFFFF;
        }
        
        QMessageBox {
            background-color: #FFFFFF;
        }
        
        QMessageBox QPushButton {
            min-width: 80px;
        }
        
        QScrollBar:vertical {
            border: none;
            background-color: #F5F5F5;
            width: 10px;
            margin: 0px;
        }

        QScrollBar::handle:vertical {
            background-color: #4CAF50;
            border-radius: 5px;
            min-height: 20px;
        }

        QScrollBar::handle:vertical:hover {
            background-color: #2E7D32;
        }
        """

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Set application icon (shows in taskbar)
    app.setWindowIcon(QIcon("icon.ico"))
    
    window = ExcelClone()
    window.show()
    sys.exit(app.exec())