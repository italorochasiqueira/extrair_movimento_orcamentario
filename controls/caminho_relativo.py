import sys
import os
from pathlib import Path

def caminho_relativo(caminho_para_figuras):
    try:
        base_path = sys._MEIPASS  # Quando rodar do executável PyInstaller
    except Exception:
        base_path = os.path.abspath(".")  # Quando rodar no ambiente dev

    return os.path.join(base_path, caminho_para_figuras)

def caminho_diretorio():
    base_diretorio = Path(__file__).resolve().parent