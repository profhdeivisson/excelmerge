from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, QMessageBox, QFileDialog)
from columnSelector import ColumnSelector
import pandas as pd

class ExcelCombiner(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle('ExcelMerge')
        self.setMinimumWidth(300)
        layout = QVBoxLayout()
        self.listWidget = QListWidget()
        layout.addWidget(self.listWidget)
        self.addButton = QPushButton('Selecione os arquivos')
        self.addButton.clicked.connect(self.addFiles)
        layout.addWidget(self.addButton)
        self.removeButton = QPushButton('Remover arquivo selecionado')
        self.removeButton.clicked.connect(self.removeSelectedFile)
        layout.addWidget(self.removeButton)
        self.clearAllButton = QPushButton('Limpar Tudo')
        self.clearAllButton.clicked.connect(self.clearFileList)
        layout.addWidget(self.clearAllButton)
        self.selectColumnsButton = QPushButton('Selecionar Colunas')
        self.selectColumnsButton.clicked.connect(self.openColumnSelector)
        layout.addWidget(self.selectColumnsButton)
        self.setLayout(layout)

    def clearFileList(self):
        self.selected_files.clear()
        self.listWidget.clear()

    def addFiles(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos", "", "Excel Files (*.xlsx)", options=options)
        if files:
            self.selected_files.extend(files)
            self.updateFileList()

    def removeSelectedFile(self):
        list_items = self.listWidget.selectedItems()
        if not list_items: return
        for item in list_items:
            self.selected_files.remove(item.text())
            self.listWidget.takeItem(self.listWidget.row(item))

    def updateFileList(self):
        self.listWidget.clear()
        for file in self.selected_files:
            QListWidgetItem(file, self.listWidget)

    def openColumnSelector(self):
        if not self.selected_files:
            QMessageBox.warning(self, 'Aviso', "Nenhum arquivo foi selecionado.")
            return
        elif len(self.selected_files) == 1:
            QMessageBox.warning(self, 'Aviso', "Selecione dois ou mais arquivos para combinar.")
            return
        try:
            columns_list = [pd.read_excel(file).columns.tolist() for file in self.selected_files]
            if not all(columns_list[0] == columns for columns in columns_list[1:]):
                QMessageBox.warning(self, 'Inconsistência', "As planilhas selecionadas têm colunas inconsistentes.")
                return
        except Exception as e:
            QMessageBox.critical(self, 'Erro', f"Ocorreu um erro ao verificar as colunas: {e}")
            return

        column_selector = ColumnSelector(columns_list[0], self.selected_files, self)
        column_selector.exec_()
