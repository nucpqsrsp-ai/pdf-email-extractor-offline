import os
import re
import sys
import csv
import traceback
from datetime import datetime
from tkinter import (
    Tk, Button, Label, Text, END, filedialog, Scrollbar, RIGHT, Y, LEFT, BOTH,
    Checkbutton, IntVar, DISABLED, NORMAL, messagebox
)
from pypdf import PdfReader
from docx import Document

APP_VERSION = "1.1.0"
REGEX_EMAIL = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')

def log(txt_widget: Text, msg: str):
    """Escreve no log (área de texto), rolando automaticamente."""
    txt_widget.config(state=NORMAL)
    txt_widget.insert(END, msg + "\n")
    txt_widget.see(END)
    txt_widget.config(state=DISABLED)
    txt_widget.update_idletasks()

def extrair_emails_de_pdf(caminho_pdf: str) -> list:
    emails = []
    with open(caminho_pdf, "rb") as f:
        reader = PdfReader(f)
        texto = []
        for i, page in enumerate(reader.pages, start=1):
            try:
                t = page.extract_text() or ""
            except Exception:
                t = ""
            texto.append(t)
        texto = "\n".join(texto)
    emails = REGEX_EMAIL.findall(texto)
    return emails

def salvar_docx(emails_unicos: list, saida_docx: str):
    doc = Document()
    doc.add_heading('Lista de E-mails Encontrados', level=1)
    doc.add_paragraph(f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}')
    doc.add_paragraph('')
    if emails_unicos:
        for e in sorted(emails_unicos, key=str.lower):
            doc.add_paragraph(e)
    else:
        doc.add_paragraph('Nenhum e-mail encontrado.')
    doc.save(saida_docx)

def salvar_csv(emails_unicos: list, saida_csv: str):
    with open(saida_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["email"])
        for e in sorted(emails_unicos, key=str.lower):
            w.writerow([e])

def processar(btn_processar, txt_log, salvar_csv_var):
    try:
        btn_processar.config(state=DISABLED)

        # limpar log
        txt_log.config(state=NORMAL)
        txt_log.delete(1.0, END)
        txt_log.config(state=DISABLED)

        # selecionar PDFs
        arquivos = filedialog.askopenfilenames(
            title="Selecione um ou mais PDFs",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not arquivos:
            log(txt_log, "⚠️ Nenhum PDF selecionado.")
            return

        log(txt_log, f"📄 PDFs selecionados: {len(arquivos)}")

        # extrair e-mails
        todos_emails = []
        for idx, pdf in enumerate(arquivos, start=1):
            log(txt_log, f"• ({idx}/{len(arquivos)}) Lendo: {os.path.basename(pdf)}")
            try:
                emails = extrair_emails_de_pdf(pdf)
                log(txt_log, f"  → {len(emails)} e-mail(s) encontrado(s)")
                todos_emails.extend(emails)
            except Exception as e:
                log(txt_log, f"  ✗ Erro ao processar {os.path.basename(pdf)}: {e}")

        unicos = sorted(set(todos_emails), key=str.lower)
        log(txt_log, f"\n📬 Total extraídos: {len(todos_emails)}  |  Únicos: {len(unicos)}")

        # pasta de saída = pasta do primeiro PDF selecionado
        pasta_saida = os.path.dirname(arquivos[0]) if arquivos else os.getcwd()

        # salvar DOCX (com fallback se não puder gravar nessa pasta)
        saida_docx = os.path.join(pasta_saida, "emails_encontrados.docx")
        try:
            salvar_docx(unicos, saida_docx)
        except Exception:
            log(txt_log, "⚠️ Não foi possível gravar o DOCX na pasta dos PDFs. Escolha outro local…")
            alt_docx = filedialog.asksaveasfilename(
                title="Salvar como",
                defaultextension=".docx",
                filetypes=[("Documento Word", "*.docx")],
                initialfile="emails_encontrados.docx"
            )
            if alt_docx:
                salvar_docx(unicos, alt_docx)
                saida_docx = alt_docx
            else:
                saida_docx = None

        # CSV opcional
        saida_csv = None
        if salvar_csv_var.get() == 1:
            try:
                saida_csv = os.path.join(pasta_saida, "emails_encontrados.csv")
                salvar_csv(unicos, saida_csv)
            except Exception:
                log(txt_log, "⚠️ Não foi possível gravar o CSV na pasta dos PDFs. Escolha outro local…")
                alt_csv = filedialog.asksaveasfilename(
                    title="Salvar CSV como",
                    defaultextension=".csv",
                    filetypes=[("CSV", "*.csv")],
                    initialfile="emails_encontrados.csv"
                )
                if alt_csv:
                    salvar_csv(unicos, alt_csv)
                    saida_csv = alt_csv

        # mensagens finais
        if saida_docx:
            log(txt_log, f"✅ DOCX gerado: {saida_docx}")
        if saida_csv:
            log(txt_log, f"✅ CSV gerado:  {saida_csv}")

        if not unicos:
            log(txt_log, "ℹ️ Dica: se o PDF for escaneado (imagem), é preciso OCR. Posso incluir OCR depois.")
        else:
            log(txt_log, "💾 Para compartilhar: use os arquivos gerados.")

        # aviso visual de conclusão
        msg_ok = "Processo concluído."
        if saida_docx:
            msg_ok += f"\nDOCX: {saida_docx}"
        if saida_csv:
            msg_ok += f"\nCSV:  {saida_csv}"
        messagebox.showinfo("Concluído", msg_ok)

    except Exception as e:
        # mostra erro num popup e no log
        log(txt_log, "❌ Falha inesperada. Detalhes no traceback abaixo:")
        log(txt_log, traceback.format_exc())
        messagebox.showerror("Erro", str(e))
    finally:
        btn_processar.config(state=NORMAL)

def main():
    root = Tk()
    root.title(f"Extrator de E-mails de PDFs — v{APP_VERSION}")
    root.geometry("700x460")

    Label(root, text="Selecione seus PDFs e clique em Processar para gerar a lista de e-mails.").pack(pady=8)

    salvar_csv_var = IntVar(value=1)
    chk_csv = Checkbutton(root, text="Gerar também CSV", variable=salvar_csv_var)
    chk_csv.pack()

    # botão principal
    btn_processar = Button(root, text="Selecionar PDFs e Processar", width=32,
                           command=lambda: processar(btn_processar, txt_log, salvar_csv_var))
    btn_processar.pack(pady=10)

    # log
    Label(root, text="Log:").pack(anchor="w", padx=8)
    txt_log = Text(root, height=16, state=DISABLED, wrap="word")
    scroll = Scrollbar(root, command=txt_log.yview)
    txt_log.configure(yscrollcommand=scroll.set)
    txt_log.pack(side=LEFT, fill=BOTH, expand=True, padx=(8,0), pady=(0,8))
    scroll.pack(side=RIGHT, fill=Y, pady=(0,8))

    log(txt_log, "ℹ️ PDFs escaneados (imagem) não possuem texto extraível. "
                 "Se precisar, posso adicionar OCR (Tesseract) em uma versão futura.")
    root.mainloop()

if __name__ == "__main__":
    # Ajuste de diretório quando empacotado (PyInstaller)
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        os.chdir(os.path.dirname(sys.executable))
    main()
