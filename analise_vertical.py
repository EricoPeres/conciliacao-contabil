import pandas as pd
from pathlib import Path

class AnaliseFinanceira:
    def __init__(self, arquivo_entrada):
        self.arquivo_entrada = Path(arquivo_entrada)
        self.df_balanco = None
        self.resultados = {}

    def carregar_dados(self):
        """Carrega os dados do Excel."""
        try:
            self.df_balanco = pd.read_excel(self.arquivo_entrada)
            if "Saldo Atual" not in self.df_balanco.columns:
                raise ValueError("Coluna 'Saldo Atual' não encontrada.")
        except Exception as e:
            raise RuntimeError(f"Erro ao carregar o arquivo: {e}")

    def analise_vertical(self):
        """Executa a análise vertical do balanço patrimonial."""
        df = self.df_balanco.copy()
        total = df["Saldo Atual"].sum()
        df["% do Total"] = (df["Saldo Atual"] / total) * 100
        self.resultados['Analise Vertical'] = df

    def salvar_resultados(self, arquivo_saida="analise_vertical.xlsx"):
        """Salva os resultados da análise em um novo arquivo Excel."""
        if not self.resultados:
            raise RuntimeError("Nenhuma análise foi executada.")
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            for nome_aba, df in self.resultados.items():
                df.to_excel(writer, sheet_name=nome_aba, index=False)

    def executar(self):
        """Executa todas as etapas da análise."""
        self.carregar_dados()
        self.analise_vertical()
        self.salvar_resultados()


if __name__ == "__main__":
    import sys

    # Substitua 'balanco.xlsx' pelo nome do seu arquivo de entrada
    entrada = "balanco.xlsx"

    analise = AnaliseFinanceira(entrada)
    try:
        analise.executar()
        print("Análise vertical concluída e salva em 'analise_vertical.xlsx'.")
    except Exception as e:
        print(f"Erro: {e}")
