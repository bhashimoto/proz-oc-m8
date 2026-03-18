from utils.oc_loader import load_oc, generate_purchase_order_bulk, validar_arquivo, ErroValidacaoArquivo
from m8.purchase_order import PurchaseOrder
from m8 import M8, BadRequestException
import shutil
from pathlib import Path
import sys
import os

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QTableWidget, QTableWidgetItem,
    QProgressBar, QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


MODELO_ARQUIVO = Path(resource_path(os.path.join("static", "OCs de Preceptores.xlsx")))

TENANT = "prozeducacao"
COMPANY_ID = 31

COL_EMPRESA = 0
COL_FORNECEDOR = 1
COL_VALOR = 2
COL_VENCIMENTO = 3
COL_OBSERVACAO = 4
COL_STATUS = 5
COL_RESULTADO = 6

COLUNAS = ["Empresa", "Fornecedor", "Valor",
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
        menu_arquivo = menu_bar.addMenu("Arquivo")
        acao_salvar_modelo = menu_arquivo.addAction(
            "Salvar planilha modelo...")
        acao_salvar_modelo.triggered.connect(self._salvar_modelo)

        menu_configuracoes = menu_bar.addMenu("Configurações")
        acao_credenciais = menu_configuracoes.addAction("Credenciais")
        acao_credenciais.triggered.connect(self._abrir_credenciais)

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)

        botoes = QHBoxLayout()
        self.import_button = QPushButton("Importar OC")
        self.import_button.clicked.connect(self._choose_file)
        botoes.addWidget(self.import_button)

        self.criar_button = QPushButton("Criar OCs no M8")
        self.criar_button.clicked.connect(self._criar_ocs)
        self.criar_button.setEnabled(False)
        botoes.addWidget(self.criar_button)

        layout.addLayout(botoes)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.tabela = QTableWidget(0, len(COLUNAS))
        self.tabela.setHorizontalHeaderLabels(COLUNAS)
        self.tabela.horizontalHeader().setStretchLastSection(True)
        self.tabela.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tabela)

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
            validar_arquivo(path)
        except ErroValidacaoArquivo as e:
            QMessageBox.critical(self, "Arquivo inválido", str(e))
            return

        items = load_oc(path)
        self._pos = generate_purchase_order_bulk(items)
        self._preencher_tabela(self._pos)
        self.criar_button.setEnabled(bool(self._pos))

    def _preencher_tabela(self, pos: list[PurchaseOrder]):
        self.tabela.setRowCount(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        for po in pos:
            row = self.tabela.rowCount()
            self.tabela.insertRow(row)
            vencimento = po.installments[0].vencimento if po.installments else ""
            valor = po.installments[0].valor if po.installments else ""
            valores = [
                str(po.empresaId),
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
                self.tabela.setItem(row, col, item)
        self.tabela.resizeColumnsToContents()

    def _criar_ocs(self):
        if not self._pos:
            return

        if self._usuario is None:
            if not self._abrir_credenciais():
                return

        self.criar_button.setEnabled(False)
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
        self.tabela.setItem(row, COL_RESULTADO, item)
        for col in range(COL_RESULTADO):
            cell = self.tabela.item(row, col)
            if cell:
                cell.setBackground(cor)
        self.progress_bar.setValue(row + 1)

    def _ao_concluir(self):
        self.import_button.setEnabled(True)
        self.criar_button.setEnabled(True)
        self.tabela.resizeColumnsToContents()
