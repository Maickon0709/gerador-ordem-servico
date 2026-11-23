"""Janela simples em Tkinter: escolher o PDF, informar estoque/PMP e gerar a OS."""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .leitura import OrcamentoInvalido, ler_orcamento
from .planilha import gerar_os


def _abrir(caminho: Path) -> None:
    """Abre a OS no programa padrão de planilhas (Excel, LibreOffice...)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(caminho)])
    except OSError:
        messagebox.showinfo("Gerador de OS", f"OS salva em:\n{caminho}")


def abrir_janela(modelo: Path) -> None:
    raiz = tk.Tk()
    raiz.title("Gerador de Ordem de Serviço")
    raiz.geometry("820x540")
    estado = {"pdf": None, "orc": None}

    topo = ttk.Frame(raiz, padding=12)
    topo.pack(fill="x")
    ttk.Label(topo, text="Orçamento (PDF):").pack(side="left")
    var_pdf = tk.StringVar()
    ttk.Entry(topo, textvariable=var_pdf, state="readonly").pack(side="left", padx=6, fill="x", expand=True)

    resumo = ttk.Label(raiz, text="Escolha o PDF do orçamento.", padding=(12, 0), wraplength=790)
    resumo.pack(fill="x")

    tabela = ttk.Treeview(raiz, columns=("item", "codigo", "descricao", "qtd", "utilizar"),
                          show="headings", height=8)
    for col, titulo, larg in [("item", "Item", 40), ("codigo", "Código", 90), ("descricao", "Descrição", 420),
                              ("qtd", "Qtde", 60), ("utilizar", "Estoque: utilizar", 120)]:
        tabela.heading(col, text=titulo)
        tabela.column(col, width=larg, anchor="w" if col == "descricao" else "center")
    tabela.pack(fill="both", expand=True, padx=12, pady=8)

    def editar_utilizar(evento):
        linha = tabela.identify_row(evento.y)
        if not linha or tabela.identify_column(evento.x) != "#5":
            return
        x, y, w, h = tabela.bbox(linha, "#5")
        entrada = ttk.Entry(tabela)
        entrada.place(x=x, y=y, width=w, height=h)
        entrada.insert(0, tabela.set(linha, "utilizar"))
        entrada.focus()

        def salvar(_=None):
            tabela.set(linha, "utilizar", entrada.get().strip())
            entrada.destroy()
        entrada.bind("<Return>", salvar)
        entrada.bind("<FocusOut>", salvar)

    tabela.bind("<Double-1>", editar_utilizar)

    opcoes = ttk.Frame(raiz, padding=12)
    opcoes.pack(fill="x")
    ttk.Label(opcoes, text="PMP:").grid(row=0, column=0, sticky="w")
    var_pmp = tk.StringVar()
    ttk.Entry(opcoes, textvariable=var_pmp, width=12).grid(row=0, column=1, sticky="w", padx=6)
    var_seguir = tk.BooleanVar()
    ttk.Checkbutton(opcoes, text="Escrever *SEGUIR DESENHO* no quadro grande",
                    variable=var_seguir).grid(row=0, column=2, sticky="w", padx=20)

    def escolher():
        caminho = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not caminho:
            return
        try:
            orc = ler_orcamento(caminho)
        except OrcamentoInvalido as erro:
            messagebox.showerror("Orçamento", str(erro))
            return
        estado.update(pdf=caminho, orc=orc)
        var_pdf.set(caminho)
        resumo.config(text=f"Orçamento {orc.numero} · {orc.cliente} · {orc.local} · frete {orc.frete} · "
                           f"prazo {orc.prazo_menor}/{orc.prazo_maior} dias úteis · {len(orc.itens)} itens\n"
                           "Duplo clique na coluna \"Estoque: utilizar\" para informar o estoque de cada item.")
        tabela.delete(*tabela.get_children())
        for it in orc.itens:
            tabela.insert("", "end", values=(it.numero, it.codigo, it.descricao, f"{it.quantidade:g}", ""))

    def gerar():
        if not estado["pdf"]:
            messagebox.showwarning("Gerador de OS", "Escolha o PDF do orçamento primeiro.")
            return
        utilizar = [tabela.set(l, "utilizar") for l in tabela.get_children()]
        saida = gerar_os(estado["pdf"], modelo, Path(estado["pdf"]).parent / "OS geradas",
                         utilizar=utilizar, pmp=var_pmp.get(), seguir_desenho=var_seguir.get())
        _abrir(saida)

    ttk.Button(topo, text="Escolher...", command=escolher).pack(side="left")
    ttk.Button(raiz, text="GERAR OS", command=gerar).pack(pady=(0, 14), ipadx=30, ipady=6)
    raiz.mainloop()
