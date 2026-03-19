from m8.purchase_order import PurchaseOrder, PurchaseOrderInstallment, PurchaseOrderItem


# ---------------------------------------------------------------------------
# PurchaseOrderItem
# ---------------------------------------------------------------------------

def test_purchase_order_item_to_dict_contains_all_fields():
    item = PurchaseOrderItem(
        produtoId=2307,
        unidadeId=5,
        quantidade=1,
        quantidadeComprada=1,
        valorUnitario=500.0,
        centroCustoId=287,
        operacaoFiscalId=250,
    )
    d = item.to_dict()
    assert d["produtoId"] == 2307
    assert d["unidadeId"] == 5
    assert d["quantidade"] == 1
    assert d["quantidadeComprada"] == 1
    assert d["valorUnitario"] == 500.0
    assert d["centroCustoId"] == 287
    assert d["operacaoFiscalId"] == 250


# ---------------------------------------------------------------------------
# PurchaseOrderInstallment
# ---------------------------------------------------------------------------

def test_purchase_order_installment_to_dict_contains_due_date_and_value():
    installment = PurchaseOrderInstallment(
        vencimento="2026-04-30T00:00:00Z",
        valor=1500.0,
    )
    d = installment.to_dict()
    assert "vencimento" in d
    assert "valor" in d
    assert d["vencimento"] == "2026-04-30T00:00:00Z"
    assert d["valor"] == 1500.0


# ---------------------------------------------------------------------------
# PurchaseOrder
# ---------------------------------------------------------------------------

def _sample_po():
    item = PurchaseOrderItem(
        produtoId=2307,
        unidadeId=5,
        quantidade=1,
        quantidadeComprada=1,
        valorUnitario=800.0,
        centroCustoId=286,
        operacaoFiscalId=250,
    )
    installment = PurchaseOrderInstallment(
        vencimento="2026-05-01T00:00:00Z",
        valor=800.0,
    )
    return PurchaseOrder(
        empresaId=31,
        status="Aprovado",
        emissao="2026-03-19T10:00:00Z",
        fornecedorId=42,
        funcionarioId=1411091897,
        freteId=9,
        condicaoPagamentoId=15,
        observacao="Observação de teste",
        items=[item],
        installments=[installment],
        tipoOrdemCompraId=1,
    )


def test_purchase_order_to_dict_without_full_excludes_items_and_installments():
    d = _sample_po().to_dict(full=False)
    assert "items" not in d
    assert "installments" not in d
    assert d["empresaId"] == 31
    assert d["status"] == "Aprovado"


def test_purchase_order_to_dict_with_full_includes_items_and_installments():
    d = _sample_po().to_dict(full=True)
    assert "items" in d
    assert "installments" in d
    assert isinstance(d["items"], list)
    assert isinstance(d["installments"], list)
    assert len(d["items"]) == 1
    assert len(d["installments"]) == 1
    # Verify sub-dicts are correctly serialized
    assert d["items"][0]["produtoId"] == 2307
    assert d["installments"][0]["valor"] == 800.0
