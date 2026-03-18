from __future__ import annotations

from pathlib import Path
from typing import Any
import datetime as dt

from utils.excel_helper import ExcelHelper
from m8.purchase_order import PurchaseOrder, PurchaseOrderInstallment, PurchaseOrderItem

OC_SHEET = "OC"
CADASTRO_SHEET = "Cadastro"
VERSOES_SUPORTADAS = {"v1"}


class ErroValidacaoArquivo(Exception):
    pass


def validar_arquivo(file_path: str | Path) -> None:
    helper = ExcelHelper(file_path)

    if CADASTRO_SHEET not in helper.sheet_names():
        raise ErroValidacaoArquivo(
            f"Planilha inválida. Utilizar a planilha modelo fornecida."
        )

    versao = helper.read_cell(row=1, col=1, sheet=CADASTRO_SHEET)
    if versao not in VERSOES_SUPORTADAS:
        versoes = ", ".join(sorted(VERSOES_SUPORTADAS))
        raise ErroValidacaoArquivo(
            f"Versão do arquivo inválida: '{
                versao}'. Versões suportadas: {versoes}."
        )


def generate_purchase_order_bulk(input: list[dict[str, Any]]) -> list[PurchaseOrder]:
    pos = []
    for item in input:
        unidade = item["Unidade"]
        fornecedorId = item["ID Fornecedor"]
        valor = item["Valor"]
        vencimento = item["Vencimento"].strftime("%Y-%m-%dT%H:%M:%SZ")
        obs = item["Observação"]

        pos.append(generate_purchase_order(
            unidade=unidade, fornecedorId=fornecedorId, valor=valor, vencimento=vencimento, obs=obs))

    return pos


def load_oc(file_path: str | Path) -> list[dict[str, Any]]:
    """Load the 'OC' sheet from an Excel file and return its table as a list of dicts.

    Row 1 is treated as the header. Rows where all values are None are skipped.
    """
    helper = ExcelHelper(file_path)
    rows = helper.read_all(sheet=OC_SHEET)
    return [row for row in rows if any(v is not None for v in row.values())]


def generate_purchase_order(unidade: str, fornecedorId: int,
                            valor: float, vencimento: str, obs: str) -> PurchaseOrder:
    EMPRESA = {
        'Carapicuíba': 1,
        'Diadema': 1,
        'Grajaú': 1,
        'GUAIANASES': 1,
        'Guarulhos': 1,
        'Itaquera': 1,
        'Jabaquara': 1,
        'Mauá': 1,
        'Sacomã': 1,
        'Santo Amaro': 1,
        'São Miguel': 1,
        'Belo Horizonte': 31,
        'Venda Nova': 31,
        'Contagem': 31,
        'Divinópolis': 31,
        'Ipatinga': 31,
        'Juiz de Fora': 31,
        'Montes Claros': 31,
        'Uberlândia': 31,
    }

    FUNCIONARIO: dict[int, int] = {
        1:  1411093179,
        31: 1411091897,
    }

    CENTRO_CUSTO: dict[str, int] = {
        'Carapicuíba': 0,
        'Diadema': 0,
        'Grajaú': 0,
        'GUAIANASES': 0,
        'Guarulhos': 0,
        'Itaquera': 0,
        'Jabaquara': 0,
        'Mauá': 0,
        'Sacomã': 0,
        'Santo Amaro': 0,
        'São Miguel': 0,
        'Belo Horizonte': 287,
        'Venda Nova': 290,
        'Contagem': 286,
        'Divinópolis': 288,
        'Ipatinga': 292,
        'Juiz de Fora': 293,
        'Montes Claros': 289,
        'Uberlândia': 291,
    }

    OP_FISCAL: dict[int, int] = {
        1: 1,  # TODO: atualizar com valor certo
        31: 250,
    }

    """
    MONTHS = {
        1: "JANEIRO",
        2: "FEVEREIRO",
        3: "MARÇO",
        4: "ABRIL",
        5: "MAIO",
        6: "JUNHO",
        7: "JULHO",
        8: "AGOSTO",
        9: "SETEMBRO",
        10: "OUTUBRO",
        11: "NOVEMBRO",
        12: "DEZEMBRO",
    }
    """

    STATUS = "Aprovado"
    FRETE_ID = 9
    CONDICAO_PAGAMENTO_ID = 15
    TIPO_ORDEM_COMPRA_ID = 1
    PRODUTO_ID = 2307
    UNIDADE_ID = 5
    QUANTIDADE = 1
    QUANTIDADE_COMPRADA = 1

    dt_emissao = dt.datetime.now()
    emissao = dt_emissao.strftime("%Y-%m-%dT%H:%M:%SZ")

    empresaId = EMPRESA[unidade]
    funcionarioId = FUNCIONARIO[empresaId]
    centroCustoId = CENTRO_CUSTO[unidade]
    operacaoFiscalId = OP_FISCAL[empresaId]
    observacao = obs

    po_installment = PurchaseOrderInstallment(
        vencimento=vencimento, valor=valor)

    po_item = PurchaseOrderItem(
        produtoId=PRODUTO_ID,
        unidadeId=UNIDADE_ID,
        quantidade=QUANTIDADE,
        quantidadeComprada=QUANTIDADE_COMPRADA,
        valorUnitario=valor,
        centroCustoId=centroCustoId,
        operacaoFiscalId=operacaoFiscalId,
    )

    po = PurchaseOrder(
        empresaId=empresaId,
        status=STATUS,
        emissao=emissao,
        fornecedorId=fornecedorId,
        funcionarioId=funcionarioId,
        freteId=FRETE_ID,
        condicaoPagamentoId=CONDICAO_PAGAMENTO_ID,
        observacao=observacao,
        items=[po_item],
        installments=[po_installment],
        tipoOrdemCompraId=TIPO_ORDEM_COMPRA_ID
    )

    return po
