import pandas as pd

arquivos_teste = ["12603.pdf", "12673.pdf", "12751.pdf"]

df_ocr = pd.read_excel("saida_mercadorias.xlsx")
df_pym = pd.read_excel("saida_pymupdf.xlsx")

for arq in arquivos_teste:
    print(f"\n===== {arq} =====")
    print("--- Tesseract ---")
    print(df_ocr[df_ocr["arquivo"] == arq][["produto", "data", "mes"]].to_string(index=False))
    print("--- PyMuPDF ---")
    print(df_pym[df_pym["arquivo"] == arq][["produto", "data", "mes"]].to_string(index=False))