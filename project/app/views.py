from django.shortcuts import render
from django.http import HttpResponse


def index(request):
    return HttpResponse("Olá, mundo. Você está na página de index do app.")
