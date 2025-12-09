from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('upload/', views.upload_pdf, name='upload'),
    # se quiser manter a rota anterior com outro nome, pode deixar:
    path('processar-boletim/', views.upload_pdf, name='processar_boletim'),
]
