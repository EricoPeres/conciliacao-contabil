import pandas as pd
from ofxparse import OfxParser
from datetime import datetime
from collections import defaultdict

def carregar_saldos_ofx(caminho_ofx):
    with open(caminho_ofx, 'r', encoding='utf-8') as arquivo:
        ofx = OfxParser.parse(arquivo)

    saldos_por_dia = defaultdict(float)

    for transacao in ofx.account.statement.transactions:
        data = transacao.date.date()
        saldos_por_dia[data] += transacao.amount

    saldo_inicial = ofx.account.statement.balance - sum(saldos_por_dia.values())
    saldo_acumulado = saldo_inicial
    saldo_diario = {}

    for data in sorted(saldos_por_dia.keys()):
        saldo_acumulado += saldos_por_dia[data]
        saldo_diario[data] = round(saldo_acumulado, 2)

    return saldo_diario

def carregar_saldos_excel(caminho_excel, coluna_data='Data', coluna_saldo='Saldo'):
    df = pd.read_excel(caminho_excel)
    df[coluna_data] = pd.to_datetime(df[coluna_data]).dt.date
    return dict(zip(df[coluna_data], df[coluna_saldo]))

def comparar_saldos(saldos_ofx, saldos_excel, caminho_saida):
    print(f"{'Data':<12} {'OFX':>12} {'Excel':>12} {'Diferença':>12}")
    print("-" * 50)

    datas = sorted(set(saldos_ofx.keys()).union(saldos_excel.keys()))
    linhas_relatorio = []

    for data in datas:
        saldo_ofx = saldos_ofx.get(data, None)
        saldo_excel = saldos_excel.get(data, None)

        if saldo_ofx is None:
            status = "Ausente OFX"
        elif saldo_excel is None:
            status = "Ausente Excel"
        else:
            diff = round(saldo_ofx - saldo_excel, 2)
            status = diff

        print(f"{data} {saldo_ofx or '---':>12} {saldo_excel or '---':>12} {status:>12}")

        if (
            saldo_ofx != saldo_excel
            or saldo_ofx is None
            or saldo_excel is None
        ):
            linhas_relatorio.append({
                'Data': data,
                'Saldo OFX': saldo_ofx,
                'Saldo Excel': saldo_excel,
                'Diferença': status
            })

    # Exportar diferenças para Excel
    if linhas_relatorio:
        df_diferencas = pd.DataFrame(linhas_relatorio)
        df_diferencas.to_excel(caminho_saida, index=False)
        print(f"\n📄 Relatório exportado para: {caminho_saida}")
    else:
        print("\n✅ Todos os saldos conferem. Nenhuma diferença encontrada.")

# 🛠️ USO
CAMINHO_OFX = 'extrato.ofx'
CAMINHO_EXCEL = 'saldo_diario.xlsx'
CAMINHO_SAIDA = 'relatorio_diferencas.xlsx'

saldos_ofx = carregar_saldos_ofx(CAMINHO_OFX)
saldos_excel = carregar_saldos_excel(CAMINHO_EXCEL)

comparar_saldos(saldos_ofx, saldos_excel, CAMINHO_SAIDA)