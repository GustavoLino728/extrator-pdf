import fitz  # PyMuPDF
import pytesseract
from pytesseract import Output
import pandas as pd
import re
import cv2
import numpy as np
import os
import json
import hashlib
from difflib import SequenceMatcher

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

PASTA_PDFS = "PDFs"                       # Pasta com os PDFs individuais
ARQUIVO_SAIDA_XLSX = "saida_mercadorias.xlsx"
ARQUIVO_SAIDA_PDF = "pdfs_unificados.pdf"
PASTA_CACHE = "cache_ocr"
DPI = 300

PRODUTOS_ESPERADOS = [
    "ABOBORA",
    "BANANA",
    "BATATA DOCE",
    "BATATA INGLESA",
    "CARA",
    "CEBOLA",
    "CENOURA",
    "CHUCHU",
    "COENTRO",
    "MAMAO",
    "MELANCIA",
    "MACAXEIRA"
    "MELAO",
    "PIMENTAO",
    "TOMATE",
    "OVOS"
]

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def get_cache_path(pdf_hash, num_pagina):
    os.makedirs(PASTA_CACHE, exist_ok=True)
    return os.path.join(PASTA_CACHE, f"{pdf_hash}_page_{num_pagina:04d}.json")


def agrupar_por_linhas(dados_ocr, tolerancia_y=15):
    linhas = {}
    for i in range(len(dados_ocr['text'])):
        texto = dados_ocr['text'][i].strip()
        if not texto:
            continue

        y = dados_ocr['top'][i]
        encontrou = False
        for y_existente in linhas.keys():
            if abs(y_existente - y) < tolerancia_y:
                linhas[y_existente].append({
                    'texto': texto,
                    'left': dados_ocr['left'][i],
                    'top': y
                })
                encontrou = True
                break

        if not encontrou:
            linhas[y] = [{
                'texto': texto,
                'left': dados_ocr['left'][i],
                'top': y
            }]

    return dict(sorted(linhas.items()))


def encontrar_produto(texto_ocr, produtos_esperados, threshold=0.55):
    if not texto_ocr:
        return None

    texto_upper = texto_ocr.upper()
    texto_limpo = re.sub(r'[^A-ZÀ-Ý\s]', ' ', texto_upper)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()

    melhor_prod = None
    melhor_score = 0

    for prod in produtos_esperados:
        if prod in texto_limpo:
            return prod

        palavras_prod = prod.split()
        if all(re.search(r'\b' + re.escape(p) + r'\b', texto_limpo) for p in palavras_prod):
            return prod

        score = SequenceMatcher(None, prod, texto_limpo).ratio()
        if score > melhor_score:
            melhor_score = score
            melhor_prod = prod

    if melhor_score >= threshold:
        return melhor_prod
    return None


def ordenar_numericamente(lista_arquivos):
    """Ordena arquivos cujos nomes são numéricos (ex: 12500.pdf)."""
    def chave(arquivo):
        nome = os.path.splitext(arquivo)[0]
        try:
            return (0, int(nome))
        except ValueError:
            return (1, nome)  # Não-numéricos vão para o fim, em ordem alfabética
    return sorted(lista_arquivos, key=chave)


# ==========================================
# EXTRAÇÃO DE UM PDF
# ==========================================
def extrair_dados_pdf(caminho_pdf, nome_arquivo):
    """Extrai dados de um PDF. Retorna lista de dicionários."""
    doc = fitz.open(caminho_pdf)
    pdf_bytes = open(caminho_pdf, 'rb').read()
    pdf_hash = hashlib.md5(pdf_bytes).hexdigest()[:10]

    dados_extraidos = []
    variacoes_kg = {"KG", "K6", "K.G", "K.G.", "K G", "KS", "K G."}

    for num_pagina, pagina in enumerate(doc, start=1):
        cache_file = get_cache_path(pdf_hash, num_pagina)

        # --- 1. CACHE / OCR ---
        if os.path.exists(cache_file):
            with open(cache_file, 'r', encoding='utf-8') as f:
                dados_ocr = json.load(f)
            texto_completo = " ".join([t for t in dados_ocr['text'] if t.strip()])
        else:
            pix = pagina.get_pixmap(dpi=DPI)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            img_cinza = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

            dados_ocr = pytesseract.image_to_data(
                img_cinza, lang='por', output_type=Output.DICT, config='--psm 6'
            )
            texto_completo = " ".join([t for t in dados_ocr['text'] if t.strip()])

            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(dados_ocr, f, ensure_ascii=False)

        # --- 2. PULAR CHECKLIST ---
        if "CHECKLIST" in texto_completo.upper():
            print(f"    [pág {num_pagina}] CHECKLIST detectado. Pulando...")
            continue

        # --- 3. DATA E MÊS ---
        data_pedido = None
        mes = None
        match_data = re.search(r'\bData:\s*(\d{2}/\d{2}/\d{4})', texto_completo)
        if match_data:
            data_pedido = match_data.group(1)
            mes = data_pedido.split('/')[1]

        # --- 4. LUGAR ---
        lugar = "NÃO IDENTIFICADO"
        linhas = agrupar_por_linhas(dados_ocr)

        for y, palavras in linhas.items():
            palavras_ordenadas = sorted(palavras, key=lambda p: p['left'])
            texto_linha = [p['texto'] for p in palavras_ordenadas]

            if "Nome:" in texto_linha:
                idx = texto_linha.index("Nome:")
                nome_partes = []
                for p in texto_linha[idx + 1:]:
                    if any(label in p for label in
                           ["L.F:", "Fone:", "Cnpj:", "Cep:", "U.F:", "L.F", "Fone",
                            "L.F.", "IE:", "CNPJ:", "C.N.P.J", "Endereço:"]):
                        break
                    nome_partes.append(p)
                lugar = " ".join(nome_partes).strip()
                break

        # --- 5. CABEÇALHO DA TABELA ---
        y_cabecalho = None
        for y, palavras in linhas.items():
            textos = [p['texto'].upper().strip('.,;:') for p in palavras]
            if "QTD" in textos:
                y_cabecalho = y
                break

        if y_cabecalho is None:
            print(f"    [pág {num_pagina}] Cabeçalho 'QTD' não encontrado. Pulando...")
            continue

        # --- 6. FIM DA TABELA ---
        y_fim = float('inf')
        for y, palavras in linhas.items():
            if y > y_cabecalho:
                texto_linha = " ".join([p['texto'] for p in palavras])
                if "Totais" in texto_linha or "Observação" in texto_linha:
                    y_fim = y
                    break

        # --- 7. PRODUTOS ---
        for y, palavras in linhas.items():
            if y <= y_cabecalho or y >= y_fim:
                continue

            palavras_ordenadas = sorted(palavras, key=lambda p: p['left'])
            texto_linha_completo = " ".join([p['texto'] for p in palavras_ordenadas])

            if any(h in texto_linha_completo for h in
                   ["Descrição", "Totais", "Peso", "U.M.", "Código", "QTD"]):
                continue

            # Primeiro "KG" da linha
            idx_kg = -1
            for i, p in enumerate(palavras_ordenadas):
                token = p['texto'].upper().strip('.,;:')
                if token in variacoes_kg:
                    idx_kg = i
                    break

            if idx_kg == -1:
                continue

            texto_candidato = " ".join([p['texto'] for p in palavras_ordenadas[:idx_kg + 1]])
            produto_canonico = encontrar_produto(texto_candidato, PRODUTOS_ESPERADOS)

            if produto_canonico is None:
                continue

            dados_extraidos.append({
                "arquivo": nome_arquivo,
                "pagina": num_pagina,
                "lugar": lugar,
                "produto": f"{produto_canonico} KG",
                "data": data_pedido,
                "mes": mes
            })

    doc.close()
    return dados_extraidos


# ==========================================
# UNIR PDFs
# ==========================================
def unir_pdfs(lista_caminhos, caminho_saida):
    """Une vários PDFs em um único arquivo."""
    doc_final = fitz.open()
    for caminho in lista_caminhos:
        doc_temp = fitz.open(caminho)
        doc_final.insert_pdf(doc_temp)
        doc_temp.close()
    doc_final.save(caminho_saida)
    doc_final.close()
    print(f"\n📎 PDF unificado salvo em: {caminho_saida}")


# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def main():
    # 1) Listar PDFs da pasta
    if not os.path.isdir(PASTA_PDFS):
        print(f"❌ Pasta '{PASTA_PDFS}' não encontrada.")
        return

    arquivos = [f for f in os.listdir(PASTA_PDFS) if f.lower().endswith('.pdf')]
    if not arquivos:
        print(f"❌ Nenhum PDF encontrado em '{PASTA_PDFS}'.")
        return

    arquivos = ordenar_numericamente(arquivos)
    print(f"📂 {len(arquivos)} PDFs encontrados (em ordem):")
    for a in arquivos:
        print(f"   - {a}")

    # 2) Extrair dados de cada PDF
    todos_dados = []
    caminhos_completos = []

    for idx, nome_arquivo in enumerate(arquivos, start=1):
        caminho = os.path.join(PASTA_PDFS, nome_arquivo)
        caminhos_completos.append(caminho)
        print(f"\n[{idx}/{len(arquivos)}] Processando '{nome_arquivo}'...")

        try:
            dados = extrair_dados_pdf(caminho, nome_arquivo)
            todos_dados.extend(dados)
            print(f"  -> {len(dados)} produtos extraídos.")
        except Exception as e:
            print(f"  ❌ Erro ao processar '{nome_arquivo}': {e}")

    # 3) Gerar planilha consolidada
    if todos_dados:
        df = pd.DataFrame(todos_dados)
        df = df[["arquivo", "pagina", "lugar", "produto", "data", "mes"]]
        df.to_excel(ARQUIVO_SAIDA_XLSX, index=False)
        print(f"\n✅ Planilha '{ARQUIVO_SAIDA_XLSX}' gerada com {len(df)} registros!")
        print("\nPrévia:")
        print(df.head(20).to_string(index=False))
    else:
        print("\n⚠️ Nenhum dado foi extraído dos PDFs.")

    # 4) Unir PDFs
    unir_pdfs(caminhos_completos, ARQUIVO_SAIDA_PDF)


if __name__ == "__main__":
    main()