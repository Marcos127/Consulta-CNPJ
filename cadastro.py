import tkinter as tk
from tkinter import ttk, messagebox
import requests
import sqlite3
import csv
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

def criar_banco():
    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS empresas (
            cnpj TEXT PRIMARY KEY,
            nome_fantasia TEXT,
            razao_social TEXT,
            ie TEXT,
            endereco TEXT,
            telefone TEXT,
            email TEXT,
            cnae TEXT,
            natureza TEXT,
            situacao TEXT,
            mei TEXT,
            simples TEXT,
            inicio TEXT
        )
    """)
    conn.commit()
    conn.close()

def consultar_cnpj():
    cnpj = entry_cnpj.get().strip()
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    response = requests.get(url)
    if response.status_code == 200:
        dados = response.json()
        for campo in campos:
            campo.delete(0, tk.END)

        entry_nome.insert(0, dados.get("nome_fantasia") or "Não disponível")
        entry_razao.insert(0, dados.get("razao_social") or "Não disponível")
        entry_ie.insert(0, "Não disponível")

        endereco = f"{dados.get('logradouro', '')}, {dados.get('numero', '')} {dados.get('complemento', '')}, {dados.get('bairro', '')}, {dados.get('municipio', '')} - {dados.get('uf', '')}, CEP: {dados.get('cep', '')}"
        entry_endereco.insert(0, endereco.strip())

        entry_telefone.insert(0, dados.get("ddd_telefone_1") or "Não disponível")
        entry_email.insert(0, dados.get("email") or "Não disponível")

        cnae_codigo = dados.get("cnae_fiscal", "")
        cnae_descricao = dados.get("cnae_fiscal_descricao", "")
        entry_cnae.insert(0, f"{cnae_codigo} - {cnae_descricao}" if cnae_descricao else "Não disponível")

        entry_natureza.insert(0, dados.get("natureza_juridica") or "Não disponível")
        entry_situacao.insert(0, dados.get("descricao_situacao_cadastral") or "Não disponível")
        entry_mei.insert(0, "Sim" if dados.get("opcao_pelo_mei") else "Não")
        entry_simples.insert(0, "Sim" if dados.get("opcao_pelo_simples") else "Não")
        entry_inicio.insert(0, dados.get("data_inicio_atividade") or "Não disponível")
    else:
        messagebox.showerror("Erro", f"CNPJ não encontrado ou API indisponível. Código: {response.status_code}")

def salvar_empresa():
    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO empresas VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        entry_cnpj.get(),
        entry_nome.get(),
        entry_razao.get(),
        entry_ie.get(),
        entry_endereco.get(),
        entry_telefone.get(),
        entry_email.get(),
        entry_cnae.get(),
        entry_natureza.get(),
        entry_situacao.get(),
        entry_mei.get(),
        entry_simples.get(),
        entry_inicio.get()
    ))
    conn.commit()
    conn.close()
    messagebox.showinfo("Sucesso", "Empresa salva com sucesso!")

def excluir_empresa():
    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM empresas WHERE cnpj = ?", (entry_cnpj.get(),))
    conn.commit()
    conn.close()
    messagebox.showinfo("Sucesso", "Empresa excluída com sucesso!")

def exportar_csv():
    caminho = "empresas_exportadas.csv"
    
    # Verifica se o arquivo está aberto
    if os.path.exists(caminho):
        try:
            os.remove(caminho)
        except PermissionError:
            messagebox.showerror("Erro", "Feche o arquivo CSV antes de exportar.")
            return

    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM empresas")
    registros = cursor.fetchall()
    conn.close()

    # Usando ponto e vírgula como separador e codificação adequada para o Excel
    with open(caminho, "w", newline="", encoding="latin-1") as f:  # ou encoding="cp1252"
        writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_ALL)
        writer.writerow([
            "CNPJ", "Nome Fantasia", "Razão Social", "IE", "Endereço", "Telefone", "Email",
            "CNAE", "Natureza Jurídica", "Situação", "MEI", "Simples", "Início Atividade"
        ])
        writer.writerows(registros)

    messagebox.showinfo("Exportação", "Arquivo CSV gerado com sucesso!\n\nDica: Ao abrir no Excel, vá em:\nDados > De um Arquivo CSV e escolha ';' como separador.")

def gerar_pdf():
    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM empresas")
    registros = cursor.fetchall()
    conn.close()

    c = canvas.Canvas("relatorio_empresas.pdf", pagesize=A4)
    largura, altura = A4
    y = altura - 50

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, y, "Relatório de Empresas")
    y -= 30

    c.setFont("Helvetica", 10)
    for empresa in registros:
        campos = [
            f"CNPJ: {empresa[0]}",
            f"Nome Fantasia: {empresa[1]}",
            f"Razão Social: {empresa[2]}",
            f"IE: {empresa[3]}",
            f"Endereço: {empresa[4]}",
            f"Telefone: {empresa[5]}",
            f"Email: {empresa[6]}",
            f"CNAE: {empresa[7]}",
            f"Natureza Jurídica: {empresa[8]}",
            f"Situação: {empresa[9]}",
            f"MEI: {empresa[10]}",
            f"Simples: {empresa[11]}",
            f"Início Atividade: {empresa[12]}"
        ]
        for campo in campos:
            c.drawString(50, y, campo)
            y -= 15
            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = altura - 50
        y -= 10

    c.save()
    messagebox.showinfo("PDF", "Relatório PDF gerado com sucesso!")

def listar_empresas():
    conn = sqlite3.connect("empresas.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM empresas")
    registros = cursor.fetchall()
    conn.close()

    janela_lista = tk.Toplevel(root)
    janela_lista.title("Empresas Salvas")
    janela_lista.geometry("1200x500")

    colunas = [
        "CNPJ", "Nome Fantasia", "Razão Social", "IE", "Endereço", "Telefone", "Email",
        "CNAE", "Natureza Jurídica", "Situação", "MEI", "Simples", "Início Atividade"
    ]

    frame = tk.Frame(janela_lista)
    frame.pack(fill="both", expand=True)

    tree = ttk.Treeview(frame, columns=colunas, show="headings")
    tree.pack(side="left", fill="both", expand=True)

    for col in colunas:
        tree.heading(col, text=col)
        tree.column(col, width=200, anchor="w")

    for empresa in registros:
        tree.insert("", "end", values=empresa)

    scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scroll_y.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scroll_y.set)

    scroll_x = ttk.Scrollbar(janela_lista, orient="horizontal", command=tree.xview)
    scroll_x.pack(side="bottom", fill="x")
    tree.configure(xscrollcommand=scroll_x.set)

    def carregar_dados(event):
        item = tree.selection()
        if item:
            valores = tree.item(item, "values")
            entry_cnpj.delete(0, tk.END)
            entry_cnpj.insert(0, valores[0])
            entry_nome.delete(0, tk.END)
            entry_nome.insert(0, valores[1])
            entry_razao.delete(0, tk.END)
            entry_razao.insert(0, valores[2])
            entry_ie.delete(0, tk.END)
            entry_ie.insert(0, valores[3])
            entry_endereco.delete(0, tk.END)
            entry_endereco.insert(0, valores[4])
            entry_telefone.delete(0, tk.END)
            entry_telefone.insert(0, valores[5])
            entry_email.delete(0, tk.END)
            entry_email.insert(0, valores[6])
            entry_cnae.delete(0, tk.END)
            entry_cnae.insert(0, valores[7])
            entry_natureza.delete(0, tk.END)
            entry_natureza.insert(0, valores[8])
            entry_situacao.delete(0, tk.END)
            entry_situacao.insert(0, valores[9])
            entry_mei.delete(0, tk.END)
            entry_mei.insert(0, valores[10])
            entry_simples.delete(0, tk.END)
            entry_simples.insert(0, valores[11])
            entry_inicio.delete(0, tk.END)
            entry_inicio.insert(0, valores[12])

    tree.bind("<Double-1>", carregar_dados)

# 🖥️ Interface principal
criar_banco()
root = tk.Tk()
root.title("Consulta de CNPJ")

tk.Label(root, text="CNPJ:").grid(row=0, column=0)
entry_cnpj = tk.Entry(root, width=40)
entry_cnpj.grid(row=0, column=1)

tk.Label(root, text="Nome Fantasia:").grid(row=1, column=0)
entry_nome = tk.Entry(root, width=40)
entry_nome.grid(row=1, column=1)

tk.Label(root, text="Razão Social:").grid(row=2, column=0)
entry_razao = tk.Entry(root, width=40)
entry_razao.grid(row=2, column=1)

tk.Label(root, text="Inscrição Estadual:").grid(row=3, column=0)
entry_ie = tk.Entry(root, width=40)
entry_ie.grid(row=3, column=1)

tk.Label(root, text="Endereço:").grid(row=4, column=0)
entry_endereco = tk.Entry(root, width=40)
entry_endereco.grid(row=4, column=1)

tk.Label(root, text="Telefone:").grid(row=5, column=0)
entry_telefone = tk.Entry(root, width=40)
entry_telefone.grid(row=5, column=1)

tk.Label(root, text="Email:").grid(row=6, column=0)
entry_email = tk.Entry(root, width=40)
entry_email.grid(row=6, column=1)

tk.Label(root, text="CNAE Principal:").grid(row=7, column=0)
entry_cnae = tk.Entry(root, width=40)
entry_cnae.grid(row=7, column=1)

tk.Label(root, text="Natureza Jurídica:").grid(row=8, column=0)
entry_natureza = tk.Entry(root, width=40)
entry_natureza.grid(row=8, column=1)

tk.Label(root, text="Situação Cadastral:").grid(row=9, column=0)
entry_situacao = tk.Entry(root, width=40)
entry_situacao.grid(row=9, column=1)

tk.Label(root, text="Opção MEI:").grid(row=10, column=0)
entry_mei = tk.Entry(root, width=40)
entry_mei.grid(row=10, column=1)

tk.Label(root, text="Opção Simples:").grid(row=11, column=0)
entry_simples = tk.Entry(root, width=40)
entry_simples.grid(row=11, column=1)

tk.Label(root, text="Início Atividade:").grid(row=12, column=0)
entry_inicio = tk.Entry(root, width=40)
entry_inicio.grid(row=12, column=1)

tk.Button(root, text="Consultar", command=consultar_cnpj).grid(row=13, column=0)
tk.Button(root, text="Salvar", command=salvar_empresa).grid(row=13, column=1)
tk.Button(root, text="Excluir", command=excluir_empresa).grid(row=14, column=1)
tk.Button(root, text="Ver Empresas Salvas", command=listar_empresas).grid(row=15, column=1)
tk.Button(root, text="Exportar CSV", command=exportar_csv).grid(row=16, column=1)
tk.Button(root, text="Gerar PDF", command=gerar_pdf).grid(row=17, column=1)

campos = [
    entry_nome, entry_razao, entry_ie, entry_endereco, entry_telefone, entry_email,
    entry_cnae, entry_natureza, entry_situacao, entry_mei, entry_simples, entry_inicio
]

root.mainloop()