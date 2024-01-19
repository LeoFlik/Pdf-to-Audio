import os
from PyPDF2 import PdfReader
import docx2txt
import pyttsx3
from customtkinter import filedialog
from customtkinter import *


def extrair_texto_pdf(caminho_pdf):
    texto = ""
    with open(caminho_pdf, "rb") as arquivo:
        leitor_pdf = PdfReader(arquivo)
        for pagina_num in range(len(leitor_pdf.pages)):
            pagina = leitor_pdf.pages[pagina_num]
            texto += pagina.extract_text()
    return texto


def extrair_texto_docx(caminho_docx):
    texto = docx2txt.process(caminho_docx)
    return texto


def converter_arquivo_para_audio(caminho_arquivo):
    if caminho_arquivo.endswith(".pdf"):
        texto = extrair_texto_pdf(caminho_arquivo)
    elif caminho_arquivo.endswith(".docx"):
        texto = extrair_texto_docx(caminho_arquivo)
    else:
        mensagem.configure(text="Formato de arquivo não suportado.")
        return

    nome_arquivo, _ = os.path.splitext(caminho_arquivo)
    caminho_saida = f"{nome_arquivo}.mp3"

    engine = pyttsx3.init()
    engine.setProperty("rate", 220)
    engine.save_to_file(texto, caminho_saida)
    engine.runAndWait()
    mensagem.configure(
        text=f"Conversão concluída.\n O arquivo de áudio foi salvo como '{caminho_saida}'."
    )


def carregar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("PDF files", "*.pdf"), ("DOCX files", "*.docx")],
    )
    if caminho_arquivo:
        converter_arquivo_para_audio(caminho_arquivo)


# Configurar a janela principal
root = CTk()
root.title("App")
root.geometry("800X400")
root.state("iconic")
set_appearance_mode("dark")
set_default_color_theme("dark-blue")

my_font = CTkFont("Sans", weight="bold", slant="roman", size=20)
my_font_a = CTkFont("Sans", weight="bold", slant="roman", size=15)

title_label = CTkLabel(root, text="Conversor de PDF/DOCX para Áudio", font=my_font)
title_label.pack_configure(pady=(10, 10), fill="both")
title_label.pack()

info = "Clique no botão para converter o arquivo.\n Alguns arquivos podem conter pequenas alterações na hora da conversão"
info_label = CTkLabel(root, text=info, font=my_font_a)
info_label.pack_configure(pady=(10, 10), fill="both", expand=True)
info_label.pack()

# Botão para carregar arquivo
botao_carregar = CTkButton(
    root, text="Carregar Arquivo", command=carregar_arquivo, border_width=3
)
botao_carregar.pack(pady=20)

# Mensagem de status
mensagem = CTkLabel(root, text="")
mensagem.pack_configure(pady=(10, 10), fill="both", expand=True)
mensagem.pack()

# Iniciar a interface gráfica
root.mainloop()
