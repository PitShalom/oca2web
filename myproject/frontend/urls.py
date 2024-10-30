# frontend/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),  # Página inicial
    path('run-automation/', views.run_automation, name='run_automation'),  # URL da automação
]
