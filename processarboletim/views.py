from django.shortcuts import render
from django.http import HttpResponse
from django.core.mail import send_mail
from .utils import gerar_relatorio
from io import BytesIO
import base64
import json
from django.http import JsonResponse, HttpResponseBadRequest
from django.conf import settings
from django.core.mail import EmailMultiAlternatives

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
            pdf_bytes, nome_pdf, alunos_em_risco = gerar_relatorio(file_obj, opcao=bimestre)
            pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

            return render(request, "result.html", {
                "pdf_b64": pdf_b64,
                "nome_pdf": nome_pdf,
                "alunos_em_risco": alunos_em_risco,
            })

        except Exception as e:
            # QUALQUER erro que acontecer vem parar aqui
            return render(request, "upload.html", {
                "erro": "O PDF enviado não está no formato esperado. Envie o boletim oficial do IFPB."
            })

    return render(request, "upload.html")



def home(request):
    return render(request, 'home.html')

def enviar_email_coped(request):
    """
    Recebe POST JSON:
    { "alunos": [...], "pdf_b64": "...", "nome_pdf": "..." }
    Envia apenas texto com a lista de alunos em situação de risco (sem PDF).
    """
    if request.method != "POST":
        return JsonResponse({"ok": False, "error": "Método não permitido."}, status=405)

    try:
        payload = json.loads(request.body.decode('utf-8'))
    except Exception:
        return HttpResponseBadRequest("JSON inválido")

    alunos = payload.get('alunos', [])
    nome_pdf = payload.get('nome_pdf', 'relatorio_saida.pdf')

    if not alunos:
        return JsonResponse({"ok": False, "error": "Nenhum aluno em risco."}, status=400)

    # monta as linhas no formato: "1 201927621032 Aline Santana Ramos de Jesus 50"
    linhas = []
    for i, a in enumerate(alunos):
        matricula = a.get('matricula', '').strip()
        nome = a.get('nome', '').strip()
        nota = a.get('nota', '')
        linhas.append(f"{i+1} {matricula} {nome} {nota}")
    alunos_texto = "\n".join(linhas)

    assunto = f"Relatório: Alunos em Situação de Risco — {nome_pdf}"

    corpo = f"""Olá, equipe COPED,

Segue o relatório \"{nome_pdf}\".
Resumo: Foram identificados {len(alunos)} aluno(s) com notas abaixo do esperado que necessitam de acompanhamento pedagógico.

Lista de alunos em situação de risco:
{alunos_texto}

Observações:
- Notas <= 40: situação crítica
- Notas < 70 (e > 40): atenção

Atenciosamente,
[sistema automatizado] — Gerador de boletins
"""


    remetente = getattr(settings, 'DEFAULT_FROM_EMAIL', None) or getattr(settings, 'EMAIL_HOST_USER', None)

    try:
        send_mail(
            subject=assunto,
            message=corpo,
            from_email=remetente,
            recipient_list=[settings.EMAIL_OFICIAL],
            fail_silently=False,
        )
        return JsonResponse({"ok": True, "message": "E-mail enviado (modo teste: ver terminal)."})
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=500)