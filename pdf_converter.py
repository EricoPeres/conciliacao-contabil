import os
import pdfplumber
import pandas as pd
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Configure o caminho para o Tesseract, se necessário (ex: no Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extrair_tabelas_ou_texto(pdf_path):
    resultados = []
    texto_ocr_completo = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()

            if tables and any(tables[0]):
                for table_index, table in enumerate(tables, start=1):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    sheet_name = f"Pag{page_number}_Tab{table_index}"
                    resultados.append((sheet_name, df))
            else:
                # OCR da imagem da página
                image = page.to_image(resolution=300).original
                text = pytesseract.image_to_string(image, lang='por')
                texto_ocr_completo.append(f"\n--- Página {page_number} ---\n{text.strip()}")
                # Também salvar no Excel
                df = pd.DataFrame([[line] for line in text.strip().split("\n") if line.strip()], columns=["Texto OCR"])
                sheet_name = f"Pag{page_number}_OCR"
                resultados.append((sheet_name, df))

    return resultados, "\n".join(texto_ocr_completo)

def processar_pdf(pdf_path, output_dir):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    excel_path = os.path.join(output_dir, f"{base_name}.xlsx")
    ocr_txt_path = os.path.join(output_dir, f"{base_name}_ocr.txt")

    print(f"🔄 Processando: {os.path.basename(pdf_path)}")

    try:
        resultados, texto_ocr = extrair_tabelas_ou_texto(pdf_path)

        if not resultados:
            print("⚠️ Nenhum conteúdo extraído.")
            return

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for sheet_name, df in resultados:
                writer_sheet = sheet_name[:31]
                df.to_excel(writer, sheet_name=writer_sheet, index=False)

        if texto_ocr.strip():
            with open(ocr_txt_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(texto_ocr)

        print(f"✅ Arquivos salvos: {os.path.basename(excel_path)}, {os.path.basename(ocr_txt_path)}")

    except Exception as e:
        print(f"❌ Erro ao processar {os.path.basename(pdf_path)}: {e}")

def main():
    pasta = input("Digite o caminho da pasta com os arquivos PDF: ").strip()

    if not os.path.isdir(pasta):
        print("❌ Pasta inválida. Verifique o caminho.")
        return

    pdfs = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
    if not pdfs:
        print("⚠️ Nenhum arquivo PDF encontrado na pasta.")
        return

    for pdf in pdfs:
        caminho_pdf = os.path.join(pasta, pdf)
        processar_pdf(caminho_pdf, pasta)

    print("\n🏁 Processamento concluído para todos os arquivos.")

if __name__ == "__main__":
    main()
