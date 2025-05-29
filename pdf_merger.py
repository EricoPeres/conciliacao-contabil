import os
from PyPDF2 import PdfMerger

def mesclar_pdfs_em_pasta(caminho_pasta, nome_arquivo_saida):
    merger = PdfMerger()

    # Verifica se a pasta existe
    if not os.path.isdir(caminho_pasta):
        print("Erro: Caminho da pasta inválido.")
        return

    # Lista e ordena os arquivos PDF
    arquivos = sorted(f for f in os.listdir(caminho_pasta) if f.lower().endswith('.pdf'))

    if not arquivos:
        print("Nenhum arquivo PDF encontrado na pasta.")
        return

    # Adiciona cada PDF
    for arquivo in arquivos:
        caminho_completo = os.path.join(caminho_pasta, arquivo)
        merger.append(caminho_completo)
        print(f"Adicionado: {arquivo}")

    # Adiciona extensão .pdf se não estiver no nome
    if not nome_arquivo_saida.lower().endswith('.pdf'):
        nome_arquivo_saida += '.pdf'

    caminho_saida = os.path.join(caminho_pasta, nome_arquivo_saida)
    merger.write(caminho_saida)
    merger.close()

    print(f"\nPDF mesclado salvo como: {caminho_saida}")

# ===== Entrada via terminal =====
if __name__ == "__main__":
    caminho = input("Digite o caminho da pasta com os PDFs: ").strip()
    nome_saida = input("Digite o nome do arquivo PDF de saída (ex: final.pdf): ").strip()
    mesclar_pdfs_em_pasta(caminho, nome_saida)

