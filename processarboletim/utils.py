import os
import pdfplumber
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from io import BytesIO

def gerar_relatorio(file_obj, opcao="1"):
    dados, imagens = [], []
    nome_pdf = "relatorio_saida.pdf"

    colunas = {"1": 6, "2": 8, "3": 10, "4": 12}
    coluna_nota = colunas.get(opcao, 6)
    nota_col_name = f"Nota {opcao} Bim"
    cabecalho_buffer = None

    with pdfplumber.open(file_obj) as pdf:
        cabecalho_bbox = (0, 0, 841, 140)
        if pdf.pages:
            page0 = pdf.pages[0]
            try:
                cabecalho_img = page0.crop(cabecalho_bbox).to_image(resolution=300)
                cabecalho_buffer = BytesIO()
                cabecalho_img.save(cabecalho_buffer, format="PNG")
                cabecalho_buffer.seek(0)
            except Exception:
                cabecalho_buffer = None

        for page in pdf.pages:
            texto = page.extract_text() or ""
            for linha in texto.split("\n"):
                if linha.startswith("Diário:"):
                    nome_pdf = linha.replace("Diário:", "").split("- Médio")[0].strip() + ".pdf"
            tables = page.extract_tables() or []
            if not tables:
                continue
            table = tables[0]
            if table:
                for row in table:
                    if row and len(row) >= 7 and row[1] and str(row[1]).isdigit() and len(str(row[1])) == 12:
                        if row[5] and str(row[5]).lower() != "cancelado":
                            dados.append([row[1], row[2], row[coluna_nota]])

    df = pd.DataFrame(dados, columns=["Matrícula", "Nome", nota_col_name]) if dados else pd.DataFrame(columns=["Matrícula", "Nome", nota_col_name])
    df[nota_col_name] = pd.to_numeric(df[nota_col_name], errors="coerce")

    if not df.empty:
        media = df[nota_col_name].mean()
        mediana = df[nota_col_name].median()
        desvio = df[nota_col_name].std()
        nota_min = df[nota_col_name].min()
        nota_max = df[nota_col_name].max()
    else:
        media = mediana = desvio = nota_min = nota_max = 0.0

    stats_data = [
        ["Estatística", "Valor"],
        ["Média", f"{media:.2f}"],
        ["Mediana", f"{mediana:.2f}"],
        ["Desvio Padrão", f"{desvio:.2f}"],
        ["Nota Mínima", f"{nota_min:.2f}"],
        ["Nota Máxima", f"{nota_max:.2f}"]
    ]
    tabela_estatisticas = Table(stats_data, colWidths=[200, 100])
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.lightslategray),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ])
    tabela_estatisticas.setStyle(estilo_tabela)

    # HISTOGRAMA
    plt.figure()
    plt.hist(df[nota_col_name].dropna(), bins=10)
    plt.title("Distribuição de Notas")
    plt.xlabel("Nota")
    plt.ylabel("Quantidade de Alunos")
    plt.grid(True)
    buffer = BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight')
    plt.close()
    buffer.seek(0)
    imagens.append(buffer)

    # BARRAS ORDENADAS
    df_sorted = df.sort_values(nota_col_name) if not df.empty else df
    plt.figure(figsize=(14, 6))
    if not df_sorted.empty:
        cores_barras = df_sorted[nota_col_name].apply(lambda x: 'green' if x >= 70 else 'red')
        plt.bar(df_sorted["Nome"], df_sorted[nota_col_name], color=cores_barras)
    plt.title("Notas dos Alunos (Ordenadas)")
    plt.xticks(rotation=90)
    plt.tight_layout()
    buffer = BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight')
    plt.close()
    buffer.seek(0)
    imagens.append(buffer)

    # PIZZA
    situacao = {
        "Ruim (0-39)": int(df[df[nota_col_name] < 40].shape[0]) if not df.empty else 0,
        "Mediana (40-69)": int(df[(df[nota_col_name] >= 40) & (df[nota_col_name] < 70)].shape[0]) if not df.empty else 0,
        "Bom (70-85)": int(df[(df[nota_col_name] >= 70) & (df[nota_col_name] <= 85)].shape[0]) if not df.empty else 0,
        "Excelente (86-100)": int(df[df[nota_col_name] > 85].shape[0]) if not df.empty else 0
    }
    labels = list(situacao.keys())
    sizes = list(situacao.values())
    plt.figure(figsize=(8, 8))
    plt.pie(sizes, labels=labels, startangle=90, wedgeprops=dict(width=0.4))
    plt.axis('equal')
    buffer = BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight')
    plt.close()
    buffer.seek(0)
    imagens.append(buffer)

    abaixo_70_df = df[df[nota_col_name] < 70].copy() if not df.empty else df.copy()
    abaixo_40 = abaixo_70_df[abaixo_70_df[nota_col_name] < 40].copy() if not abaixo_70_df.empty else abaixo_70_df.copy()
    entre_40_70 = abaixo_70_df[(abaixo_70_df[nota_col_name] >= 40) & (abaixo_70_df[nota_col_name] < 70)].copy() if not abaixo_70_df.empty else abaixo_70_df.copy()
    if not abaixo_40.empty:
        abaixo_40.insert(0, "#", range(1, len(abaixo_40)+1))
    if not entre_40_70.empty:
        entre_40_70.insert(0, "#", range(1, len(entre_40_70)+1))

    tabela_abaixo_40 = Table([["#", "Matrícula", "Nome", "Nota"]] + (abaixo_40.values.tolist() if not abaixo_40.empty else []), colWidths=[30, 90, 140, 70])
    tabela_entre_40_70 = Table([["#", "Matrícula", "Nome", "Nota"]] + (entre_40_70.values.tolist() if not entre_40_70.empty else []), colWidths=[30, 90, 140, 70])
    tabela_abaixo_40.setStyle(estilo_tabela)
    tabela_entre_40_70.setStyle(estilo_tabela)

    final_pdf = BytesIO()
    pdf = canvas.Canvas(final_pdf, pagesize=landscape(A4))
    largura, altura = landscape(A4)

    try:
        if cabecalho_buffer is not None:
            cabecalho_buffer.seek(0)
            pdf.drawImage(ImageReader(cabecalho_buffer), 0, altura - 140, 841, 140)
    except Exception:
        pass

    try:
        if os.path.exists("verdes.jpg"):
            pdf.drawImage("verdes.jpg", 450, 10, 380, 380)
    except Exception:
        pass

    pdf.rect(30, altura - 200, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    titulo_infogerais = f"INFORMAÇÕES GERAIS DA TURMA {opcao}º BIM"
    pdf.drawString(60, altura - 197, titulo_infogerais)
    tabela_estatisticas.wrapOn(pdf, 300, 120)
    tabela_estatisticas.drawOn(pdf, 30, altura - 330)
    pdf.showPage()

    pdf.rect(30, altura - 45, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(145, altura - 43, "ALUNOS ENTRE 40 E 70")
    w = tabela_entre_40_70.wrapOn(pdf, 381, 120)[1] if tabela_entre_40_70 else 0
    tabela_entre_40_70.drawOn(pdf, 30, altura-55-w)

    pdf.rect(431, altura - 45, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(largura/2 + 135, altura - 43, "ALUNOS ABAIXO DE 40")
    w2 = tabela_abaixo_40.wrapOn(pdf, 381, 120)[1] if tabela_abaixo_40 else 0
    tabela_abaixo_40.drawOn(pdf, 431, altura-55-w2)

    pdf.showPage()

    imagens_por_pagina = 2
    largura_img = largura * 0.8
    altura_img = altura / imagens_por_pagina * 0.8

    for i in range(0, len(imagens), imagens_por_pagina):
        imagens_pagina = imagens[i:i+imagens_por_pagina]
        for j, img in enumerate(imagens_pagina):
            try:
                img.seek(0)
                imagem = ImageReader(img)
                iw, ih = imagem.getSize()
                escala = min(largura_img / iw, altura_img / ih)
                iw *= escala
                ih *= escala
                x = (largura - iw) / 2
                y = altura - (j + 1) * (altura / imagens_por_pagina) + ((altura / imagens_por_pagina - ih) / 2)
                pdf.drawImage(imagem, x, y, width=iw, height=ih)
            except Exception:
                continue
        pdf.showPage()

    pdf.save()
    final_pdf.seek(0)
    return final_pdf.getvalue(), nome_pdf
