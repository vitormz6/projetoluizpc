import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttkb
import pandas as pd
import os
from datetime import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from PIL import Image, ImageTk  # Import para lidar com imagens (logo)

# -------- CONFIGURAÇÕES GERAIS --------

FILE_NAME = "controle_inversores.xlsx"  # Excel local
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
CREDENTIALS_FILE = "credentials.json"
SPREADSHEET_NAME = "Controle de Inversores"

def autenticar_google_sheets():
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME).sheet1

def enviar_para_google_sheets(nova_linha):
    try:
        sheet = autenticar_google_sheets()
        sheet.append_row(nova_linha, value_input_option="RAW")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar para Google Sheets: {str(e)}")

def inicializar_planilha_local():
    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(
            columns=[
                "Data",
                "Tipo",
                "Código de Barras Inversor",
                "QR Code",
                "Responsável",
                "Status",
                "Cliente Destino",
                "Marca",
                "Observações",
            ]
        )
        df.to_excel(FILE_NAME, index=False)

def atualizar_tabela():
    for row in tabela.get_children():
        tabela.delete(row)
    df = pd.read_excel(FILE_NAME)
    for _, row_data in df.iterrows():
        tabela.insert("", "end", values=row_data.tolist())

def registrar_movimento(
    tipo,
    entry_codigo_barras,
    entry_qr_code,
    combo_responsavel,
    combo_status,
    combo_marca,
    entry_observacoes
):
    codigo_barras = entry_codigo_barras.get().strip()
    qr_code = entry_qr_code.get().strip() or "N/A"
    responsavel = combo_responsavel.get().strip()
    status = combo_status.get().strip()
    marca = combo_marca.get().strip()
    observacoes = entry_observacoes.get().strip()
    cliente_destino = ""  

    if not codigo_barras or not responsavel:
        messagebox.showwarning("Aviso", "Preencha os campos obrigatórios (cód. barras, responsável)!")
        return

    df = pd.read_excel(FILE_NAME)

    data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nova_linha = [
        data_atual,
        tipo,
        codigo_barras,
        qr_code,
        responsavel,
        status,
        cliente_destino,
        marca,
        observacoes,
    ]
    df.loc[len(df)] = nova_linha
    df.to_excel(FILE_NAME, index=False)

    enviar_para_google_sheets(nova_linha)

    entry_codigo_barras.delete(0, tk.END)
    entry_qr_code.delete(0, tk.END)
    entry_observacoes.delete(0, tk.END)

    atualizar_tabela()
    messagebox.showinfo("Sucesso", f"{tipo} registrada com sucesso!")


# ========== CRIAÇÃO DA JANELA PRINCIPAL ==========

tela = ttkb.Window(themename="flatly")
tela.title("Controle de Inversores")
tela.geometry("1000x600")
tela.iconbitmap("icon.ico")

# -------- CARREGANDO LOGO --------
try:
    # Abre a imagem (logo.png, por exemplo)
    img = Image.open("luiz.jpg")
    # Ajuste o tamanho conforme quiser (ex: 120x120)
    img = img.resize((120, 120), Image.Resampling.LANCZOS)
    # Transforma em PhotoImage para usar no Tkinter
    logo_tk = ImageTk.PhotoImage(img)

    # Cria um label para exibir o logo
    logo_label = ttkb.Label(tela, image=logo_tk)
    # Precisa manter a referência em uma variável (aqui, atribuída como atributo do label)
    logo_label.image = logo_tk
    # Posiciona no topo
    logo_label.pack(pady=5)
except Exception as e:
    # Se der erro (arquivo não encontrado ou etc.), só ignora
    print(f"Falha ao carregar logo: {e}")


inicializar_planilha_local()

# ---------- NOTEBOOK (ABAS) ----------
notebook = ttkb.Notebook(tela, bootstyle="primary")
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# ========== ABA 1: FORMULÁRIO ==========

aba_form = ttkb.Frame(notebook)
notebook.add(aba_form, text="Formulário")

titulo_label = ttkb.Label(
    aba_form,
    text="Registrar Movimento (Entrada / Saída)",
    font=("Helvetica", 18, "bold")
)
titulo_label.pack(pady=10)

form_frame = ttkb.Labelframe(aba_form, text="Preencha os dados", padding=20, bootstyle="info")
form_frame.pack(fill="x", padx=20, pady=10)

lbl_cod = ttkb.Label(form_frame, text="Código de Barras Inversor:", font=("Arial", 12))
lbl_cod.grid(row=0, column=0, sticky="w", padx=5, pady=5)
entry_codigo_barras = ttkb.Entry(form_frame, font=("Arial", 12), width=40)
entry_codigo_barras.grid(row=0, column=1, padx=5, pady=5)

lbl_qr = ttkb.Label(form_frame, text="QR Code (Opcional):", font=("Arial", 12))
lbl_qr.grid(row=1, column=0, sticky="w", padx=5, pady=5)
entry_qr_code = ttkb.Entry(form_frame, font=("Arial", 12), width=40)
entry_qr_code.grid(row=1, column=1, padx=5, pady=5)

lbl_resp = ttkb.Label(form_frame, text="Responsável:", font=("Arial", 12))
lbl_resp.grid(row=2, column=0, sticky="w", padx=5, pady=5)
combo_responsavel = ttkb.Combobox(form_frame, font=("Arial", 12), values=["Luiz", "Maria", "João"], width=37)
combo_responsavel.current(0)
combo_responsavel.grid(row=2, column=1, padx=5, pady=5)

lbl_status = ttkb.Label(form_frame, text="Status:", font=("Arial", 12))
lbl_status.grid(row=3, column=0, sticky="w", padx=5, pady=5)
combo_status = ttkb.Combobox(
    form_frame, font=("Arial", 12), 
    values=[
        "APTO - Inversor testado e pronto para utilização em campo",
        "INAPTO - Inversor com defeito elétrico",
        "INAPTO - Inversor com defeito antena wifi"
    ],
    width=37
)
combo_status.current(0)
combo_status.grid(row=3, column=1, padx=5, pady=5)

lbl_marca = ttkb.Label(form_frame, text="Marca:", font=("Arial", 12))
lbl_marca.grid(row=4, column=0, sticky="w", padx=5, pady=5)
combo_marca = ttkb.Combobox(form_frame, font=("Arial", 12), values=["Huawei", "Chint", "Goodwe", "SMA", "Foxess"], width=37)
combo_marca.current(0)
combo_marca.grid(row=4, column=1, padx=5, pady=5)

lbl_obs = ttkb.Label(form_frame, text="Observações:", font=("Arial", 12))
lbl_obs.grid(row=5, column=0, sticky="w", padx=5, pady=5)
entry_observacoes = ttkb.Entry(form_frame, font=("Arial", 12), width=40)
entry_observacoes.grid(row=5, column=1, padx=5, pady=5)

botoes_frame = ttkb.Frame(aba_form)
botoes_frame.pack(pady=10)

btn_entrada = ttkb.Button(
    botoes_frame,
    text="Registrar Entrada",
    bootstyle="success",
    command=lambda: registrar_movimento(
        "Entrada",
        entry_codigo_barras,
        entry_qr_code,
        combo_responsavel,
        combo_status,
        combo_marca,
        entry_observacoes
    )
)
btn_entrada.grid(row=0, column=0, padx=10)

btn_saida = ttkb.Button(
    botoes_frame,
    text="Registrar Saída",
    bootstyle="danger",
    command=lambda: registrar_movimento(
        "Saída",
        entry_codigo_barras,
        entry_qr_code,
        combo_responsavel,
        combo_status,
        combo_marca,
        entry_observacoes
    )
)
btn_saida.grid(row=0, column=1, padx=10)

# ========== ABA 2: RELATÓRIO ==========
aba_relatorio = ttkb.Frame(notebook)
notebook.add(aba_relatorio, text="Relatório")

titulo_label2 = ttkb.Label(
    aba_relatorio,
    text="Relatório de Inversores",
    font=("Helvetica", 18, "bold")
)
titulo_label2.pack(pady=10)

relatorio_btn_frame = ttkb.Frame(aba_relatorio)
relatorio_btn_frame.pack(pady=5)

btn_atualizar = ttkb.Button(
    relatorio_btn_frame,
    text="Atualizar Relatório",
    bootstyle="info",
    command=atualizar_tabela
)
btn_atualizar.pack()

frame_tabela = ttkb.Frame(aba_relatorio)
frame_tabela.pack(pady=10, fill="both", expand=True)

scrollbar_y = ttkb.Scrollbar(frame_tabela, orient="vertical")
scrollbar_y.pack(side="right", fill="y")

scrollbar_x = ttkb.Scrollbar(frame_tabela, orient="horizontal")
scrollbar_x.pack(side="bottom", fill="x")

colunas = (
    "Data",
    "Tipo",
    "Código de Barras Inversor",
    "QR Code",
    "Responsável",
    "Status",
    "Cliente Destino",
    "Marca",
    "Observações",
)

global tabela
tabela = ttkb.Treeview(
    frame_tabela,
    columns=colunas,
    show="headings",
    yscrollcommand=scrollbar_y.set,
    xscrollcommand=scrollbar_x.set,
    bootstyle="info"
)
tabela.pack(side="left", fill="both", expand=True)

scrollbar_y.config(command=tabela.yview)
scrollbar_x.config(command=tabela.xview)

tabela.heading("Data", text="Data")
tabela.heading("Tipo", text="Tipo")
tabela.heading("Código de Barras Inversor", text="Código de Barras Inversor")
tabela.heading("QR Code", text="QR Code")
tabela.heading("Responsável", text="Responsável")
tabela.heading("Status", text="Status")
tabela.heading("Cliente Destino", text="Cliente Destino")
tabela.heading("Marca", text="Marca")
tabela.heading("Observações", text="Observações")

tabela.column("Data", stretch=True, minwidth=120, width=150)
tabela.column("Tipo", stretch=True, minwidth=80, width=100)
tabela.column("Código de Barras Inversor", stretch=True, minwidth=120, width=180)
tabela.column("QR Code", stretch=True, minwidth=80, width=120)
tabela.column("Responsável", stretch=True, minwidth=80, width=120)
tabela.column("Status", stretch=True, minwidth=120, width=250)
tabela.column("Cliente Destino", stretch=True, minwidth=80, width=150)
tabela.column("Marca", stretch=True, minwidth=80, width=100)
tabela.column("Observações", stretch=True, minwidth=120, width=220)

# Carrega a tabela inicial
atualizar_tabela()

tela.mainloop()
