# frontend/views.py
from django.shortcuts import render
from django.http import HttpResponse
import subprocess
import os

def index(request):
    return render(request, 'index.html')

def run_automation(request):
    # Caminho absoluto para o script principal
    script_path = os.path.join(os.path.dirname(__file__), '../../sistemaOca.py')

    # Executar o script com subprocess
    try:
        subprocess.run(['python', script_path], check=True)
        return HttpResponse("Automação executada com sucesso!")
    except subprocess.CalledProcessError as e:
        return HttpResponse(f"Erro ao executar a automação: {e}")
