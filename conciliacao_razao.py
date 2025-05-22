import pandas as pd

def importar_razao_excel(caminho_arquivo):
    """Importa o arquivo Excel contendo o razão contábil."""
    try:
        df = pd.read_excel(caminho_arquivo)
        print("Arquivo importado com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao importar o arquivo: {e}")
        return None

def conciliacao_por_valor(df):
    """Realiza a conciliação entre débitos e créditos de mesmo valor."""
    
    # Certifique-se de que os valores estão numéricos
    df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
    df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)

    # Cria colunas auxiliares
    df['Valor'] = df['Débito'] - df['Crédito']
    df['Conciliado'] = False

    conciliacoes = []

    for i, linha in df.iterrows():
        if df.at[i, 'Conciliado'] or df.at[i, 'Valor'] == 0:
            continue

        for j in df.index[i+1:]:
            if df.at[j, 'Conciliado']:
                continue

            # Verifica se os valores se anulam (Débito-Crédito)
            if abs(df.at[i, 'Valor'] + df.at[j, 'Valor']) < 0.01:
                conciliacoes.append((i, j))
                df.at[i, 'Conciliado'] = True
                df.at[j, 'Conciliado'] = True
                break

    print(f"{len(conciliacoes)} pares conciliados.")

    return conciliacoes, df

def exibir_resultado(conciliacoes, df):
    """Exibe os pares conciliados."""
    for i, j in conciliacoes:
        print("\n--- Conciliação encontrada ---")
        print("Lançamento 1:")
        print(df.loc[i])
        print("Lançamento 2:")
        print(df.loc[j])

    nao_conciliados = df[~df['Conciliado']]
    if not nao_conciliados.empty:
        print("\n--- Lançamentos não conciliados ---")
        print(nao_conciliados[['Débito', 'Crédito', 'Valor']])
    else:
        print("\nTodos os lançamentos foram conciliados!")

if __name__ == "__main__":
    caminho = input("Digite o caminho do arquivo Excel: ")
    df_razao = importar_razao_excel(caminho)
    
    if df_razao is not None:
        conciliacoes, df_resultado = conciliacao_por_valor(df_razao)
        exibir_resultado(conciliacoes, df_resultado)
