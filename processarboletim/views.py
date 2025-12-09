from django.shortcuts import render
from django.http import HttpResponse
from .utils import gerar_relatorio
from io import BytesIO
import base64

def upload_pdf(request):
    if request.method == "POST":
        arquivo = request.FILES.get("arquivo_pdf")
        bimestre = request.POST.get("bimestre", "1")

        if not arquivo:
            return render(request, "upload.html", {
                "erro": "Nenhum arquivo foi enviado."
            })

        try:
            # Lê o PDF
            file_bytes = arquivo.read()
            file_obj = BytesIO(file_bytes)

            # Tenta gerar o relatório
            pdf_bytes, nome_pdf = gerar_relatorio(file_obj, opcao=bimestre)

            # Codifica para exibir no iframe
            pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

            return render(request, "result.html", {
                "pdf_b64": pdf_b64,
                "nome_pdf": nome_pdf
            })

        except Exception as e:
            # QUALQUER erro que acontecer vem parar aqui
            return render(request, "upload.html", {
                "erro": "O PDF enviado não está no formato esperado. Envie o boletim oficial do IFPB."
            })

    return render(request, "upload.html")



def home(request):
    return render(request, 'home.html')
