from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QPushButton, QCheckBox, QHBoxLayout, QScrollArea, QWidget, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt
import pandas as pd

class ColumnSelector(QDialog):
    def __init__(self, columns, selected_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Selecionar Colunas')
        self.columns = columns
        self.selected_files = selected_files
        self.selected_columns = columns.copy()  # Inicializar com todas as colunas
        self.checkboxes = []  # Lista para armazenar as caixas de seleção
        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout(self)
        self.setMinimumWidth(285)
        # Área de rolagem para as caixas de seleção
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scrollContent = QWidget(scroll)
        scrollLayout = QVBoxLayout(scrollContent)
        scrollContent.setLayout(scrollLayout)

        # Botão 'Selecionar Tudo/Deselecionar Tudo'
        self.toggleSelectButton = QPushButton('Desmarcar Tudo', self)
        self.toggleSelectButton.clicked.connect(self.toggleSelection)
        scrollLayout.addWidget(self.toggleSelectButton)

        # Adicionando as caixas de seleção ao layout de rolagem
        for column in self.columns:
            cb = QCheckBox(column)
            cb.setChecked(True)
            cb.stateChanged.connect(self.updateSelection)
            self.checkboxes.append(cb)  # Adicionar à lista de caixas de seleção
            scrollLayout.addWidget(cb)

        # Definindo o widget de rolagem como o widget central
        scroll.setWidget(scrollContent)
        main_layout.addWidget(scroll)

        # Layout horizontal para os botões
        buttons_layout = QHBoxLayout()

        # Botão 'Gerar Planilha'
        self.generateButton = QPushButton('Gerar Planilha', self)
        self.generateButton.clicked.connect(self.generateSheet)
        buttons_layout.addWidget(self.generateButton)

        # Botão 'Cancelar'
        self.cancelButton = QPushButton('Cancelar', self)
        self.cancelButton.clicked.connect(self.reject)
        buttons_layout.addWidget(self.cancelButton)

        # Adicionando o layout dos botões ao layout principal
        main_layout.addLayout(buttons_layout)

        self.setLayout(main_layout)

    def toggleSelection(self):
        # Verifica se todas as caixas estão marcadas para ajustar o texto do botão e a ação
        if all(cb.isChecked() for cb in self.checkboxes):
            for cb in self.checkboxes:
                cb.setChecked(False)
            self.toggleSelectButton.setText('Selecionar Tudo')
        else:
            for cb in self.checkboxes:
                cb.setChecked(True)
            self.toggleSelectButton.setText('Desmarcar Tudo')

    def updateSelection(self, state):
        sender = self.sender()
        column_text = sender.text()
        if state == Qt.Checked and column_text not in self.selected_columns:
            self.selected_columns.append(column_text)
        elif state != Qt.Checked and column_text in self.selected_columns:
            self.selected_columns.remove(column_text)
            self.toggleSelectButton.setText('Selecionar Tudo')

    def generateSheet(self):
        try:
            if not self.selected_columns:
                QMessageBox.warning(self, 'Aviso', "Nenhuma coluna foi selecionada.")
                return

            dataframes = [pd.read_excel(file, usecols=self.selected_columns) for file in self.selected_files]
            all_data = pd.concat(dataframes, ignore_index=True)
            options = QFileDialog.Options()
            save_path, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo combinado", "", "Excel Files (*.xlsx)", options=options)
            
            if save_path:
                all_data.to_excel(save_path, index=False)
                QMessageBox.information(self, 'Sucesso', f"Todos os dados foram combinados com sucesso em '{save_path}'.")
                self.parent().clearFileList()  # Limpar a lista de arquivos na interface principal
                self.accept()  # Fechar a janela após salvar
            else:
                QMessageBox.warning(self, 'Aviso', "A operação de salvar foi cancelada pelo usuário.")
        except Exception as e:
            QMessageBox.critical(self, 'Erro', f"Ocorreu um erro ao gerar a planilha: {e}")
