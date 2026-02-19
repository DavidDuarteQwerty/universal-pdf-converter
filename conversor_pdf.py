import os
import comtypes.client
from pathlib import Path
from PIL import Image
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from docx2pdf import convert as word_to_pdf

# DETECTA AUTOMATICAMENTE A PASTA DE DOWNLOADS DE QUALQUER USUÁRIO
DOWNLOADS = Path(os.path.join(os.path.expanduser("~"), "Downloads"))

# ---------------- FUNÇÕES DE CONVERSÃO ----------------

def excel_to_pdf(input_path, output_path):
    excel = comtypes.client.CreateObject("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(str(input_path))
        wb.ExportAsFixedFormat(0, str(output_path))
        wb.Close()
    finally:
        excel.Quit()

def ppt_to_pdf(input_path, output_path):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    try:
        ppt = powerpoint.Presentations.Open(str(input_path), WithWindow=False)
        ppt.SaveAs(str(output_path), 32)
        ppt.Close()
    finally:
        powerpoint.Quit()

# ---------------- INTERFACE E LÓGICA ----------------

def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecionar Documentos ou Imagens",
        filetypes=[
            ("Todos suportados", "*.docx;*.xlsx;*.pptx;*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.tiff"),
            ("Documentos Office", "*.docx;*.xlsx;*.pptx"),
            ("Imagens", "*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.tiff")
        ]
    )
    if arquivos:
        for f in arquivos:
            if f not in lista_arquivos.get(0, tk.END):
                lista_arquivos.insert(tk.END, f)
        status_var.set(f" {lista_arquivos.size()} arquivo(s) na fila")

def processar_conversao():
    arquivos = lista_arquivos.get(0, tk.END)
    if not arquivos:
        messagebox.showwarning("Atenção", "Adicione arquivos antes de converter.")
        return

    status_var.set("⏳ Processando... Aguarde.")
    root.update()

    sucessos = []
    try:
        imagens_pil = []
        for caminho in arquivos:
            p = Path(caminho)
            ext = p.suffix.lower()
            saida = DOWNLOADS / f"{p.stem}.pdf"

            if ext == '.docx':
                word_to_pdf(str(p), str(saida))
                sucessos.append(saida.name)
            elif ext == '.xlsx':
                excel_to_pdf(p, saida)
                sucessos.append(saida.name)
            elif ext == '.pptx':
                ppt_to_pdf(p, saida)
                sucessos.append(saida.name)
            elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tiff']:
                img = Image.open(caminho).convert("RGB")
                imagens_pil.append(img)

        if imagens_pil:
            nome_album = nome_pdf_var.get().strip() or "Imagens_Unidas"
            caminho_album = DOWNLOADS / f"{nome_album}.pdf"
            imagens_pil[0].save(caminho_album, save_all=True, append_images=imagens_pil[1:])
            sucessos.append(caminho_album.name)

        status_var.set("✅ Concluído!")
        messagebox.showinfo("Sucesso", f"Conversão finalizada!\n{len(sucessos)} PDFs gerados em Downloads.")
        lista_arquivos.delete(0, tk.END)
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha: {e}")
        status_var.set("❌ Erro no processo.")

# ---------------- UI TKINTER ----------------

root = tk.Tk()
root.title("Conversor Universal PDF")
root.geometry("600x520")
root.configure(bg="#f0f0f0")

status_var = tk.StringVar(value=" Pronto")
nome_pdf_var = tk.StringVar(value="Arquivo_de_Imagens")

# Header
header = tk.Frame(root, bg="#333")
header.pack(fill="x")
tk.Label(header, text="CONVERSOR OFFICE & IMAGENS", bg="#333", fg="white", font=("Arial", 12, "bold")).pack(pady=10)

# Main
main = tk.Frame(root, bg="#f0f0f0", padx=20, pady=10)
main.pack(fill="both", expand=True)

tk.Label(main, text="Nome do PDF final (para fotos):", bg="#f0f0f0").pack(anchor="w")
ttk.Entry(main, textvariable=nome_pdf_var).pack(fill="x", pady=5)

frame_lista = tk.Frame(main)
frame_lista.pack(fill="both", expand=True, pady=10)
scroll = ttk.Scrollbar(frame_lista)
scroll.pack(side="right", fill="y")
lista_arquivos = tk.Listbox(frame_lista, font=("Segoe UI", 9), borderwidth=1, relief="solid")
lista_arquivos.pack(fill="both", expand=True)
lista_arquivos.config(yscrollcommand=scroll.set)
scroll.config(command=lista_arquivos.yview)

btn_frame = tk.Frame(main, bg="#f0f0f0")
btn_frame.pack(fill="x")
ttk.Button(btn_frame, text="+ Adicionar", command=selecionar_arquivos).pack(side="left", padx=2)
ttk.Button(btn_frame, text="Limpar", command=lambda: lista_arquivos.delete(0, tk.END)).pack(side="left", padx=2)

tk.Button(main, text="CONVERTER PARA PDF", bg="#0078d7", fg="white", font=("Arial", 10, "bold"), 
          relief="flat", command=processar_conversao).pack(fill="x", pady=15, ipady=5)

footer = tk.Frame(root, bg="#ddd")
footer.pack(fill="x", side="bottom")
tk.Label(footer, textvariable=status_var, bg="#ddd", font=("Arial", 8)).pack(side="left")

root.mainloop()
