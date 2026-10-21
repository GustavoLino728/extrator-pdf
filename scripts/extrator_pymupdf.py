import pymupdf
import pandas as pd
import re
import os
from difflib import SequenceMatcher

# ==========================================
# CONFIGURAÇÕES
# ==========================================
PASTA_PDFS = "PDFs"
ARQUIVO_SAIDA_XLSX = "saida_pymupdf.xlsx"
ARQUIVO_SAIDA_PDF = "pdfs_unificados_pymupdf.pdf"

PRODUTOS_ESPERADOS = [
    "ABOBORA", "BANANA", "BATATA DOCE", "BATATA INGLESA", "CARA",
    "CEBOLA", "CENOURA", "CHUCHU", "COENTRO", "MACAXEIRA",
    "MAMAO", "MELANCIA", "MELAO", "OVOS", "PIMENTAO", "TOMATE",
]

# Produtos que podem aparecer SEM o sufixo "KG" (ex: OVOS)
PRODUTOS_SEM_KG = {"OVOS", "OVO"}

VARIACOES_KG = {"KG", "K6", "K.G", "K.G.", "K G", "KS"}
TOLERANCIA_LINHA = 3.0
HEADER_KEYWORDS = {
    "QTD", "QTDE", "DESCRICAO", "DESCRIÇÃO", "CODIGO", "CÓDIGO",
    "U.M.", "U.M", "UM", "PESO", "PRODUTO",
}

DEBUG = False  # Mude para True para depurar linha-a-linha


# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def extrair_palavras_pagina(pagina):
    return [
        {"texto": w[4], "x0": w[0], "y0": w[1], "x1": w[2], "y1": w[3]}
        for w in pagina.get_text("words")
    ]


def agrupar_por_linhas(palavras, tolerancia=TOLERANCIA_LINHA):
    linhas = {}
    for p in palavras:
        y = p["y0"]
        encontrou = False
        for y_existente in linhas.keys():
            if abs(y_existente - y) < tolerancia:
                linhas[y_existente].append(p)
                encontrou = True
                break
        if not encontrou:
            linhas[y] = [p]
    return dict(sorted(linhas.items()))


def parsear_data(texto):
    """Extrai DD/MM/AAAA de várias formas, inclusive OCR degradado."""
    if not texto:
        return None

    # 1) Com separadores: 29/11/2025, 29-11-2025, 29.11.2025
    for m in re.finditer(r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})', texto):
        d, mo, a = m.groups()
        if len(a) == 2:
            a = "20" + a
        try:
            if 1 <= int(d) <= 31 and 1 <= int(mo) <= 12 and 2020 <= int(a) <= 2100:
                return f"{int(d):02d}/{int(mo):02d}/{a}"
        except ValueError:
            continue

    # 2) Sem separadores: remove TUDO que não é dígito
    #    "2911A2025" -> "29112025" -> janela de 8 -> 29/11/2025
    digitos = re.sub(r'\D', '', texto)
    for i in range(len(digitos) - 7):
        s = digitos[i:i + 8]
        d, mo, a = s[:2], s[2:4], s[4:]
        try:
            if 1 <= int(d) <= 31 and 1 <= int(mo) <= 12 and 2020 <= int(a) <= 2100:
                return f"{d}/{mo}/{a}"
        except ValueError:
            continue

    return None


def eh_label_data(token):
    """Reconhece 'Data:' mesmo com OCR degradado (Dala:, Dsta:, DATA)."""
    t = token.upper().strip(":.,; ")
    if "DATA" in t or "DALA" in t or "DSTA" in t:
        return True
    # Padrão D?T? (4 letras começando com D)
    if len(t) == 4 and t.startswith("D"):
        return True
    return False


def encontrar_data(palavras):
    """Localiza a data com estratégia em camadas."""

    # 1) Achar o rótulo "Data:" (tolerante a OCR)
    idx_data = None
    for i, p in enumerate(palavras):
        if eh_label_data(p["texto"]):
            idx_data = i
            break

    if idx_data is not None:
        p_data = palavras[idx_data]
        x_data, y_data = p_data["x0"], p_data["y0"]

        # Candidatos ordenados por proximidade
        candidatos = []
        for p in palavras:
            if p is p_data:
                continue
            dy = p["y0"] - y_data
            dx = p["x0"] - x_data

            # Mesma linha (±20 pt) e à direita
            if abs(dy) < 20 and dx > 0:
                candidatos.append((abs(dy) + dx * 0.01, p["texto"]))
            # Linha de baixo próxima (±60 pt), mesma região X (±500 pt)
            elif 0 < dy < 60 and abs(dx) < 500:
                candidatos.append((dy + abs(dx) * 0.01, p["texto"]))
            # Mesmo bloco, mais abaixo (OCR empilhado)
            elif 0 < dy < 300 and abs(dx) < 100:
                candidatos.append((dy * 2 + abs(dx) * 0.01, p["texto"]))

        candidatos.sort()
        for _, texto in candidatos[:15]:
            d = parsear_data(texto)
            if d:
                return d

        # Junta os 8 primeiros candidatos e tenta
        texto_junto = " ".join(t for _, t in candidatos[:8])
        d = parsear_data(texto_junto)
        if d:
            return d

    # 2) Fallback global: primeira data plausível na página
    for p in palavras:
        d = parsear_data(p["texto"])
        if d:
            return d

    return None


def encontrar_produto(texto, produtos, threshold=0.75, min_word_score=0.70):
    """Fuzzy token-a-token contra a whitelist."""
    if not texto:
        return None
    texto_limpo = re.sub(r'[^A-ZÀ-Ý\s]', ' ', texto.upper())
    tokens = [t for t in texto_limpo.split() if t]
    if not tokens:
        return None

    texto_junto = " ".join(tokens)
    melhor, melhor_score = None, 0

    for prod in produtos:
        # 1. Match exato de substring
        if prod in texto_junto:
            return prod

        # 2. Token-a-token
        palavras_prod = prod.split()
        scores = []
        for pp in palavras_prod:
            melhor_local = max(
                (SequenceMatcher(None, pp, t).ratio() for t in tokens),
                default=0,
            )
            scores.append(melhor_local)

        # Cada palavra do produto precisa bater individualmente
        if min(scores) < min_word_score:
            continue

        score_medio = sum(scores) / len(scores)
        if score_medio > melhor_score:
            melhor_score, melhor = score_medio, prod

    return melhor if melhor_score >= threshold else None


def separar_produtos_colados(produto, produtos):
    nome = produto.upper().replace(" KG", "").strip()
    for prod in produtos:
        if nome.startswith(prod) and nome != prod:
            resto = nome[len(prod):].strip()
            if resto in produtos:
                return [f"{prod} KG", f"{resto} KG"]
    return [produto]


def encontrar_cabecalho(linhas):
    melhor_y, melhor_score = None, 0
    for y, palavras_linha in linhas.items():
        textos = [p["texto"].upper().strip('.,;:') for p in palavras_linha]
        score = sum(1 for t in textos if any(k in t for k in HEADER_KEYWORDS))
        if score >= 2 and score > melhor_score:
            melhor_score = score
            melhor_y = y
    return melhor_y


def ordenar_numericamente(lista):
    def chave(a):
        nome = os.path.splitext(a)[0]
        try:
            return (0, int(nome))
        except ValueError:
            return (1, nome)
    return sorted(lista, key=chave)


# ==========================================
# EXTRAÇÃO DE UM PDF
# ==========================================
def extrair_dados_pdf(caminho_pdf, nome_arquivo):
    doc = pymupdf.open(caminho_pdf)
    dados = []

    for num_pagina, pagina in enumerate(doc, start=1):
        palavras = extrair_palavras_pagina(pagina)
        if not palavras:
            continue

        linhas = agrupar_por_linhas(palavras)
        texto_pagina = " ".join(p["texto"] for p in palavras)

        if "CHECKLIST" in texto_pagina.upper():
            print(f"    [pág {num_pagina}] CHECKLIST. Pulando.")
            continue

        # --- DATA ---
        data_pedido = encontrar_data(palavras)
        mes = data_pedido.split('/')[1] if data_pedido else None

        # --- LUGAR ---
        lugar = "NÃO IDENTIFICADO"
        for y, palavras_linha in linhas.items():
            ordenadas = sorted(palavras_linha, key=lambda p: p["x0"])
            textos = [p["texto"] for p in ordenadas]
            if "Nome:" in textos:
                idx = textos.index("Nome:")
                partes = textos[idx + 1:]
                if not partes:
                    ys_maiores = [yy for yy in linhas if yy > y]
                    if ys_maiores:
                        proxima = sorted(linhas[min(ys_maiores)], key=lambda p: p["x0"])
                        partes = [p["texto"] for p in proxima]
                nome_partes = []
                for t in partes:
                    if any(lbl in t for lbl in [
                        "L.F:", "Fone:", "Cnpj:", "Cep:", "U.F:", "L.F",
                        "Fone", "L.F.", "IE:", "CNPJ:", "C.N.P.J", "Endereço:"
                    ]):
                        break
                    nome_partes.append(t)
                lugar = " ".join(nome_partes).strip()
                break

        # --- CABEÇALHO ---
        y_cabecalho = encontrar_cabecalho(linhas)
        modo = "cabeçalho" if y_cabecalho is not None else "varredura"

        # --- FIM DA TABELA ---
        y_fim = float('inf')
        if y_cabecalho is not None:
            for y, palavras_linha in linhas.items():
                if y > y_cabecalho:
                    txt = " ".join(p["texto"] for p in palavras_linha)
                    if "Totais" in txt or "Observação" in txt:
                        y_fim = y
                        break

        # --- PRODUTOS ---
        produtos_pagina = []

        for y, palavras_linha in linhas.items():
            # Se temos cabeçalho: começa DEPOIS dele, termina NO/Antes do y_fim
            # (a linha y_fim em si pode conter OVOS junto com "Totais")
            if y_cabecalho is not None:
                if y <= y_cabecalho:
                    continue
                if y > y_fim:
                    continue

            ordenadas = sorted(palavras_linha, key=lambda p: p["x0"])
            texto_linha = " ".join(p["texto"] for p in ordenadas)

            # Pular linhas que claramente são cabeçalho
            if any(h in texto_linha for h in
                   ["Descrição", "U.M.", "Código", "CÓDIGO", "QTD"]):
                continue

            # --- 1. Tentar via "KG" (produto normal) ---
            idx_kg = -1
            for i, p in enumerate(ordenadas):
                token = p["texto"].upper().strip('.,;:')
                if token in VARIACOES_KG:
                    idx_kg = i
                    break

            if idx_kg != -1:
                texto_candidato = " ".join(p["texto"] for p in ordenadas[:idx_kg + 1])
                prod = encontrar_produto(texto_candidato, PRODUTOS_ESPERADOS)

                if DEBUG:
                    print(f"      [KG] {texto_candidato!r} -> {prod}")

                if prod is not None:
                    for p_final in separar_produtos_colados(f"{prod} KG", PRODUTOS_ESPERADOS):
                        if p_final not in produtos_pagina:
                            produtos_pagina.append(p_final)
                continue

            # --- 2. Tentar produtos SEM KG (OVOS, etc.) ---
            for prod_sem_kg in PRODUTOS_SEM_KG:
                if re.search(r'\b' + re.escape(prod_sem_kg) + r'\b', texto_linha.upper()):
                    if DEBUG:
                        print(f"      [SEM KG] {texto_linha!r} -> {prod_sem_kg}")
                    if prod_sem_kg not in produtos_pagina:
                        produtos_pagina.append(prod_sem_kg)
                    break

        print(f"    [pág {num_pagina}] modo={modo}, produtos={len(produtos_pagina)}, data={data_pedido}")

        for prod in produtos_pagina:
            dados.append({
                "arquivo": nome_arquivo,
                "pagina": num_pagina,
                "lugar": lugar,
                "produto": prod,
                "data": data_pedido,
                "mes": mes,
            })

    doc.close()
    return dados


# ==========================================
# UNIR PDFs
# ==========================================
def unir_pdfs(lista_caminhos, caminho_saida):
    doc_final = pymupdf.open()
    for caminho in lista_caminhos:
        doc_temp = pymupdf.open(caminho)
        doc_final.insert_pdf(doc_temp)
        doc_temp.close()
    doc_final.save(caminho_saida)
    doc_final.close()
    print(f"\n📎 PDF unificado salvo em: {caminho_saida}")


# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def main():
    if not os.path.isdir(PASTA_PDFS):
        print(f"❌ Pasta '{PASTA_PDFS}' não encontrada.")
        return

    arquivos = [f for f in os.listdir(PASTA_PDFS) if f.lower().endswith('.pdf')]
    if not arquivos:
        print(f"❌ Nenhum PDF em '{PASTA_PDFS}'.")
        return

    arquivos = ordenar_numericamente(arquivos)
    print(f"📂 {len(arquivos)} PDFs encontrados.")

    todos_dados = []
    caminhos_completos = []

    for idx, nome_arquivo in enumerate(arquivos, start=1):
        caminho = os.path.join(PASTA_PDFS, nome_arquivo)
        caminhos_completos.append(caminho)
        print(f"\n[{idx}/{len(arquivos)}] '{nome_arquivo}'")

        try:
            dados = extrair_dados_pdf(caminho, nome_arquivo)
            todos_dados.extend(dados)
            print(f"  -> {len(dados)} produtos.")
        except Exception as e:
            print(f"  ❌ Erro: {e}")

    if todos_dados:
        df = pd.DataFrame(todos_dados)
        df = df[["Arquivo", "Pagina", "Lugar", "Produto", "Data", "Mês"]]
        df.to_excel(ARQUIVO_SAIDA_XLSX, index=False)
        print(f"\n✅ '{ARQUIVO_SAIDA_XLSX}' com {len(df)} registros!")
        print(f"   Datas preenchidas: {df['data'].notna().sum()} / {len(df)}")
        print(f"   Lugar identificado: {(df['lugar'] != 'NÃO IDENTIFICADO').sum()} / {len(df)}")
        print(f"   OVOS capturados: {(df['produto'] == 'OVOS').sum()}")
    else:
        print("\n⚠️ Nenhum dado extraído.")

    unir_pdfs(caminhos_completos, ARQUIVO_SAIDA_PDF)


if __name__ == "__main__":
    main()