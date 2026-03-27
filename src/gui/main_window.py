from utils.oc_loader import load_oc, generate_purchase_order_bulk, validate_file, FileValidationError
from m8.purchase_order import PurchaseOrder
from m8 import M8, BadRequestException
import shutil
from pathlib import Path
import sys
import os

import markdown

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QTableWidget, QTableWidgetItem,
    QProgressBar, QMessageBox, QTextBrowser,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


MODELO_ARQUIVO = Path(resource_path(
    os.path.join("static", "OCs de Preceptores.xlsx")))
ABOUT_ARQUIVO = Path(resource_path(os.path.join("static", "about.md")))

TENANT = "prozeducacao"
COMPANY_ID = 31

COL_CHECKBOX = 0
COL_EMPRESA = 1
COL_FORNECEDOR = 2
COL_VALOR = 3
COL_VENCIMENTO = 4
COL_OBSERVACAO = 5
COL_STATUS = 6
COL_RESULTADO = 7

COLUNAS = ["", "Empresa", "Tipo", "Fornecedor", "Valor",
           "Vencimento", "Observação", "Status", "Resultado"]

COR_SUCESSO = QColor("#c8e6c9")
COR_ERRO = QColor("#ffcdd2")


class CredenciaisDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Credenciais")

        layout = QFormLayout(self)

        self.usuario_input = QLineEdit()
        layout.addRow("Usuário:", self.usuario_input)

        self.senha_input = QLineEdit()
        self.senha_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addRow("Senha:", self.senha_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def valores(self) -> tuple[str, str]:
        return self.usuario_input.text(), self.senha_input.text()


class SobreDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sobre")
        self.resize(480, 340)

        layout = QVBoxLayout(self)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        html = markdown.markdown(ABOUT_ARQUIVO.read_text(encoding="utf-8"))
        browser.setHtml(html)
        layout.addWidget(browser)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class CriarOCWorker(QThread):
    progresso = pyqtSignal(int, bool, str)  # row, sucesso, mensagem
    concluido = pyqtSignal()

    def __init__(self, m8: M8, pos: list[PurchaseOrder]):
        super().__init__()
        self._m8 = m8
        self._pos = pos

    def run(self):
        for i, po in enumerate(self._pos):
            try:
                self._m8.create_purchase_order(po, full=True)
                self.progresso.emit(i, True, "Criado")
            except (BadRequestException, Exception) as e:
                self.progresso.emit(i, False, str(e))
        self.concluido.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ordem de Compra — M8")

        menu_bar = self.menuBar()
        menu_file = menu_bar.addMenu("Arquivo")
        action_save_model = menu_file.addAction(
            "Salvar planilha modelo...")
        action_save_model.triggered.connect(self._salvar_modelo)

        menu_configs = menu_bar.addMenu("Configurações")
        action_credentials = menu_configs.addAction("Credenciais")
        action_credentials.triggered.connect(self._abrir_credenciais)

        action_about = menu_bar.addAction("Sobre")
        action_about.triggered.connect(lambda: SobreDialog(self).exec())

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)

        buttons = QHBoxLayout()
        self.import_button = QPushButton("Importar OC")
        self.import_button.clicked.connect(self._choose_file)
        buttons.addWidget(self.import_button)

        self.create_button = QPushButton("Criar OCs no M8")
        self.create_button.clicked.connect(self._criar_ocs)
        self.create_button.setEnabled(False)
        buttons.addWidget(self.create_button)

        layout.addLayout(buttons)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.table = QTableWidget(0, len(COLUNAS))
        self.table.setHorizontalHeaderLabels(COLUNAS)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setColumnWidth(COL_CHECKBOX, 30)
        layout.addWidget(self.table)

        self.delete_button = QPushButton("Apagar OCs temporárias")
        self.delete_button.clicked.connect(self._apagar_temporarias)
        self.delete_button.setEnabled(False)
        layout.addWidget(self.delete_button)

        self._usuario: str | None = None
        self._senha: str | None = None
        self._m8 = M8()
        self._pos: list[PurchaseOrder] = []
        self._worker: CriarOCWorker | None = None

    def _salvar_modelo(self):
        dest, _ = QFileDialog.getSaveFileName(
            self, "Salvar planilha modelo", MODELO_ARQUIVO.name, "Excel (*.xlsx)"
        )
        if not dest:
            return
        shutil.copy2(MODELO_ARQUIVO, dest)

    def _abrir_credenciais(self) -> bool:
        dialog = CredenciaisDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._usuario, self._senha = dialog.valores()
            self._m8.set_credentials(
                username=self._usuario,
                password=self._senha,
                tenant=TENANT,
                company_id=COMPANY_ID,
            )
            return True
        return False

    def _choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "Excel (*.xlsx)")
        if not path:
            return

        try:
            validate_file(path)
        except FileValidationError as e:
            QMessageBox.critical(self, "Arquivo inválido", str(e))
            return

        items = load_oc(path)
        self._pos = generate_purchase_order_bulk(items)
        self._preencher_table(self._pos)
        self.create_button.setEnabled(bool(self._pos))
        self.delete_button.setEnabled(bool(self._pos))

    def _preencher_table(self, pos: list[PurchaseOrder]):
        self.table.setRowCount(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        for po in pos:
            row = self.table.rowCount()
            self.table.insertRow(row)

            checkbox_item = QTableWidgetItem()
            checkbox_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            checkbox_item.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(row, COL_CHECKBOX, checkbox_item)

            vencimento = po.installments[0].vencimento if po.installments else ""
            valor = po.installments[0].valor if po.installments else ""
            valores = [
                str(po.empresaId),
                str(po.tipoOC),
                str(po.fornecedorId),
                str(valor),
                vencimento,
                po.observacao,
                po.status,
                "",
            ]
            for col, texto in enumerate(valores):
                item = QTableWidgetItem(texto)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(row, col + 1, item)
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(COL_CHECKBOX, 30)

    def _criar_ocs(self):
        if not self._pos:
            return

        if self._usuario is None:
            if not self._abrir_credenciais():
                return

        self.create_button.setEnabled(False)
        self.import_button.setEnabled(False)

        self.progress_bar.setMaximum(len(self._pos))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        self._worker = CriarOCWorker(self._m8, self._pos)
        self._worker.progresso.connect(self._ao_progresso)
        self._worker.concluido.connect(self._ao_concluir)
        self._worker.start()

    def _ao_progresso(self, row: int, sucesso: bool, mensagem: str):
        cor = COR_SUCESSO if sucesso else COR_ERRO
        item = QTableWidgetItem(mensagem)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        item.setBackground(cor)
        self.table.setItem(row, COL_RESULTADO, item)
        for col in range(1, COL_RESULTADO):
            cell = self.table.item(row, col)
            if cell:
                cell.setBackground(cor)
        checkbox = self.table.item(row, COL_CHECKBOX)
        if checkbox:
            checkbox.setFlags(Qt.ItemFlag.ItemIsUserCheckable)
        self.progress_bar.setValue(row + 1)

    def _apagar_temporarias(self):
        rows_to_delete = []
        for row in range(self.table.rowCount()):
            checkbox = self.table.item(row, COL_CHECKBOX)
            resultado = self.table.item(row, COL_RESULTADO)
            is_checked = checkbox and checkbox.checkState() == Qt.CheckState.Checked
            is_unsent = not resultado or not resultado.text()
            if is_checked and is_unsent:
                rows_to_delete.append(row)

        if not rows_to_delete:
            return

        reply = QMessageBox.question(
            self,
            "Confirmar exclusão",
            f"Deseja apagar {len(rows_to_delete)} OC(s) selecionada(s)?\n\n"
            "Esta ação remove as OCs somente desta aplicação. "
            "Nenhum registro será excluído no M8.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        for row in reversed(rows_to_delete):
            self.table.removeRow(row)
            del self._pos[row]

        self.create_button.setEnabled(bool(self._pos))
        self.delete_button.setEnabled(bool(self._pos))

    def _ao_concluir(self):
        self.import_button.setEnabled(True)
        self.create_button.setEnabled(True)
        self.table.resizeColumnsToContents()
