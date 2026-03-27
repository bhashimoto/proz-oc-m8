from __future__ import annotations

from pathlib import Path
from typing import Any
import datetime as dt

from utils.excel_helper import ExcelHelper
from m8.purchase_order import PurchaseOrder, PurchaseOrderInstallment, PurchaseOrderItem

OC_SHEET = "OC"
CADASTRO_SHEET = "Cadastro"
SUPPORTED_VERSIONS = {"v2"}


class FileValidationError(Exception):
    pass


def validate_file(file_path: str | Path) -> None:
    helper = ExcelHelper(file_path)

    if CADASTRO_SHEET not in helper.sheet_names():
        raise FileValidationError(
            "Planilha inválida. Utilizar a planilha modelo fornecida."
        )

    version = helper.read_cell(row=1, col=1, sheet=CADASTRO_SHEET)
    if version not in SUPPORTED_VERSIONS:
        supported = ", ".join(sorted(SUPPORTED_VERSIONS))
        raise FileValidationError(
            f"Versão do arquivo inválida: '{
                version}'. Versões suportadas: {supported}."
        )


def load_oc(file_path: str | Path) -> list[dict[str, Any]]:
    """Load the 'OC' sheet from an Excel file and return its data as a list of dicts.

    Row 1 is treated as the header. Rows where all values are None are skipped.
    """
    helper = ExcelHelper(file_path)
    rows = helper.read_all(sheet=OC_SHEET)
    return [row for row in rows if any(v is not None for v in row.values())]


def generate_purchase_order(unidade: str, tipo_oc: str, fornecedorId: int,
                            valor: float, vencimento: str, obs: str) -> PurchaseOrder:
    TIPO_OC_VALIDO = [
        "professor",
        "preceptor",
    ]

    if tipo_oc not in TIPO_OC_VALIDO:
        raise ValueError(f"tipo de OC inválido: {tipo_oc}")

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

    if unidade not in EMPRESA:
        raise ValueError(f"Unidade inválida: {unidade}")

    EMPLOYEE: dict[int, int] = {
        1:  1411093179,
        31: 1411091897,
    }

    COST_CENTER: dict[str, int] = {
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

    FISCAL_OP: dict[int, dict[str, int]] = {
        1: {
            "professor": 1,
            "preceptor": 1,
        },  # TODO: update with correct value
        31: {
            "professor": 252,
            "preceptor": 250,
        }
    }

    STATUS = "Aprovado"
    FREIGHT_ID = 9
    PAYMENT_CONDITION_ID = 15
    PURCHASE_ORDER_TYPE_ID = 1
    PRODUCT_ID = 2307
    UNIT_ID = 5
    QUANTITY = 1
    PURCHASED_QUANTITY = 1

    issued_at = dt.datetime.now()
    emissao = issued_at.strftime("%Y-%m-%dT%H:%M:%SZ")

    empresaId = EMPRESA[unidade]
    funcionarioId = EMPLOYEE[empresaId]
    centroCustoId = COST_CENTER[unidade]
    operacaoFiscalId = FISCAL_OP[empresaId][tipo_oc]

    vencimento = vencimento.replace("T00", "T04")

    po_installment = PurchaseOrderInstallment(
        vencimento=vencimento, valor=valor)

    po_item = PurchaseOrderItem(
        produtoId=PRODUCT_ID,
        unidadeId=UNIT_ID,
        quantidade=QUANTITY,
        quantidadeComprada=PURCHASED_QUANTITY,
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
        freteId=FREIGHT_ID,
        condicaoPagamentoId=PAYMENT_CONDITION_ID,
        observacao=obs,
        items=[po_item],
        installments=[po_installment],
        tipoOrdemCompraId=PURCHASE_ORDER_TYPE_ID,
        tipoOC=tipo_oc
    )

    return po


def generate_purchase_order_bulk(input: list[dict[str, Any]]) -> list[PurchaseOrder]:
    pos = []
    for item in input:
        unidade = item["Unidade"]
        fornecedorId = item["ID Fornecedor"]
        valor = item["Valor"]
        vencimento = item["Vencimento"].strftime("%Y-%m-%dT%H:%M:%SZ")
        obs = item["Observação"]
        tipo_oc = item["Tipo"].lower()

        pos.append(generate_purchase_order(
            unidade=unidade, fornecedorId=fornecedorId, valor=valor, vencimento=vencimento, obs=obs, tipo_oc=tipo_oc))

    return pos
