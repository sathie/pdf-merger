#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jul 16 16:30:52 2022

@author: sarah
"""

import sys, subprocess, os, platform
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import(
    QWidget,
    QApplication,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QLabel,
    QPushButton,
    QFrame,
    QListWidget,
    QFileDialog,
    QAbstractItemView,
    QMessageBox
)
from PyQt5.QtCore import pyqtSlot
from PyPDF2 import PdfMerger
from PyPDF2.errors import PdfReadError

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("PDF Merger")
        
        outer_layout = QVBoxLayout()
        outer_layout.setSpacing(10)
        
        # Buttons on top
        button_layout = QHBoxLayout()
        button_addfiles = QPushButton('Add files', self)
        button_addfiles.clicked.connect(self.select_files)
        button_layout.addWidget(button_addfiles)
        
        button_rmfiles = QPushButton('Remove files', self)
        button_rmfiles.clicked.connect(self.remove_file)
        button_layout.addWidget(button_rmfiles)
        
        outer_layout.addLayout(button_layout)
        
        # List Widget
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        outer_layout.addWidget(self.list_widget)
        label_hint = QLabel("Drag & drop to reorder items")
        label_hint.setAlignment(Qt.AlignCenter)
        label_hint.setStyleSheet("QLabel {font-style: italic;}")
        outer_layout.addWidget(label_hint)
        
        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setLineWidth(1)
        outer_layout.addWidget(sep)
        
        # Metadata form
        form_layout = QFormLayout()
        label_metadata = QLabel("Metadata")
        label_metadata.setAlignment(Qt.AlignCenter)
        label_metadata.setStyleSheet("QLabel {font-size: 15px; font-weight: bold;}")
        form_layout.addRow(label_metadata)
        
        self.entry_title = QLineEdit(self)
        self.entry_author = QLineEdit(self)
        form_layout.addRow("Title:", self.entry_title)
        form_layout.addRow("Author:", self.entry_author)
        
        outer_layout.addLayout(form_layout)
        
        # Merge Button
        button_merge = QPushButton("Merge files", self)
        button_merge.setStyleSheet("QPushButton {font-size: 14px;}")
        button_merge.clicked.connect(self.merge_files)
        outer_layout.addWidget(button_merge)
        button_merge.adjustSize()
        
        self.setLayout(outer_layout)
        
    def pdf_merge(self, files, outfile):
        merger = PdfMerger()

        for pdf in files:
            merger.append(pdf, import_bookmarks = True)
        
        metadata = {"/Producer": "Python"}
        
        title = self.entry_title.text()
        if title:
            metadata["/Title"] = title

        author = self.entry_author.text()
        if author:
            metadata["/Author"] = author
        
        merger.add_metadata(metadata)
        merger.write(outfile)
        merger.close()
        
        return True
    
    def error_message(self, title, message, icon):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(icon)
        msg.exec()
        
    @pyqtSlot()
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, 
            "QFileDialog.getOpenFileNames()",
            "",
            "PDF Files (*.pdf);;All Files (*)")
        
        if files:
            for num, file in enumerate(files):
                self.list_widget.addItem(file)
    
    @pyqtSlot()
    def remove_file(self):
        selected_items = self.list_widget.selectedItems()
        if selected_items:
            for item in selected_items:
                self.list_widget.takeItem(self.list_widget.row(item))
                
    @pyqtSlot()
    def merge_files(self):
        infiles = []
        for index in range(self.list_widget.count()):
            infiles.append(self.list_widget.item(index).text())
        
        if not infiles:
            self.error_message("Error!",
                               "No files selected",
                               QMessageBox.Warning)
            return False
        
        outfile, _ = QFileDialog.getSaveFileName(self, 
            "QFileDialog.getSaveFileName()",
            "",
            "PDF Files (*.pdf);;All Files (*)")
        
        try:
            self.pdf_merge(infiles, outfile)
        except PdfReadError:
            self.error_message("Error!",
                               "Could not read files.",
                               QMessageBox.Critical)
        except:
            self.error_message("Error!",
                "Something unexpected happened.\nCould not merge files.",
                QMessageBox.Critical)
        else:
            if platform.system() == "Darwin": #macOS
                subprocess.call(("open", outfile))
            elif platform.system() == "Windows": #Windows
                os.startfile(outfile)
            else:
                subprocess.call(("xdg-open", outfile))
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())