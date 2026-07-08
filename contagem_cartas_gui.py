# -*- coding: utf-8 -*-
"""Interface gráfica da contagem de cartas.

Serve para escolher a planilha, aplicar o filtro de datas e gerar o arquivo
final sem precisar rodar pela linha de comando.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import date

# Garante que a GUI consiga importar o módulo principal mesmo abrindo direto pelo arquivo.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from contagem_cartas import (
    processar_planilha,
    exportar_resultados,
    carregar_config,
    extrair_datas_planilha,
)


# Paleta usada na interface.
CORES = {
    "fundo": "#f1f5f9",
    "card": "#ffffff",
    "primaria": "#0f766e",
    "primaria_hover": "#0d9488",
    "texto": "#1e293b",
    "texto_secundario": "#64748b",
    "sucesso": "#059669",
    "aviso": "#d97706",
    "borda": "#e2e8f0",
}


def abrir_pasta(caminho):
    """Abre a pasta do resultado no explorador do Windows."""
    caminho = Path(caminho)
    if caminho.is_file():
        caminho = caminho.parent
    if not caminho.is_dir():
        return
    os.startfile(caminho) if sys.platform == "win32" else None


class AppContagemCartas(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Contagem de Cartas – Autos de Infração")
        self.geometry("640x680")
        self.minsize(520, 520)
        self.configure(bg=CORES["fundo"])

        self.arquivo_planilha = tk.StringVar()
        self.pasta_saida = tk.StringVar()
        self.arquivo_gerado = None
        self.totais = None
        self._em_processamento = False
        # O filtro é montado depois que a planilha é lida.
        self.datas_unicas = []  # lista de date
        self.meses_anos = []   # lista de (mes, ano)
        self.filtro_modo = tk.StringVar(value="todos")  # "todos" | "mes" | "datas"
        self._carregando_datas = False

        self._criar_estilos()
        self._criar_widgets()

    def _criar_estilos(self):
        self.estilo = ttk.Style(self)
        self.estilo.theme_use("clam")

        # Estilo do botão principal.
        self.estilo.configure(
            "Primario.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(20, 10),
            background=CORES["primaria"],
            foreground="white",
        )
        self.estilo.map(
            "Primario.TButton",
            background=[("active", CORES["primaria_hover"]), ("disabled", CORES["texto_secundario"])],
        )

        # Estilo dos blocos da tela.
        self.estilo.configure(
            "Card.TFrame",
            background=CORES["card"],
        )
        self.estilo.configure(
            "Card.TLabel",
            background=CORES["card"],
            foreground=CORES["texto"],
            font=("Segoe UI", 10),
        )
        self.estilo.configure(
            "Titulo.TLabel",
            background=CORES["fundo"],
            foreground=CORES["primaria"],
            font=("Segoe UI", 16, "bold"),
        )
        self.estilo.configure(
            "Subtitulo.TLabel",
            background=CORES["fundo"],
            foreground=CORES["texto_secundario"],
            font=("Segoe UI", 9),
        )

    def _criar_widgets(self):
        # Container principal da janela.
        container = ttk.Frame(self, padding=24)
        container.pack(fill=tk.BOTH, expand=True)

        # Cabeçalho da tela.
        ttk.Label(
            container,
            text="Contagem de Cartas",
            style="Titulo.TLabel",
        ).pack(anchor=tk.W)
        ttk.Label(
            container,
            text="Processe a planilha de autos e gere o resultado por dia e por CPF/CNPJ.",
            style="Subtitulo.TLabel",
        ).pack(anchor=tk.W, pady=(0, 16))

        # Bloco da planilha de entrada.
        frame_arquivo = ttk.Frame(container, style="Card.TFrame", padding=16)
        frame_arquivo.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(frame_arquivo, text="Planilha Excel (.xlsx)", style="Card.TLabel").pack(anchor=tk.W)
        row_arq = ttk.Frame(frame_arquivo)
        row_arq.pack(fill=tk.X, pady=(6, 0))
        self.entry_arquivo = ttk.Entry(row_arq, textvariable=self.arquivo_planilha, width=50)
        self.entry_arquivo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(
            row_arq,
            text="Selecionar…",
            command=self._selecionar_planilha,
        ).pack(side=tk.RIGHT)

        # Bloco da pasta de saída.
        frame_pasta = ttk.Frame(container, style="Card.TFrame", padding=16)
        frame_pasta.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(
            frame_pasta,
            text="Pasta de saída (opcional – vazio = mesma pasta da planilha)",
            style="Card.TLabel",
        ).pack(anchor=tk.W)
        row_pasta = ttk.Frame(frame_pasta)
        row_pasta.pack(fill=tk.X, pady=(6, 0))
        self.entry_pasta = ttk.Entry(row_pasta, textvariable=self.pasta_saida, width=50)
        self.entry_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(
            row_pasta,
            text="Selecionar…",
            command=self._selecionar_pasta,
        ).pack(side=tk.RIGHT)

        # Bloco do filtro de datas.
        frame_filtro = ttk.Frame(container, style="Card.TFrame", padding=16)
        frame_filtro.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(
            frame_filtro,
            text="Filtro de datas (Data de inscrição)",
            style="Card.TLabel",
        ).pack(anchor=tk.W)
        row_btn_carregar = ttk.Frame(frame_filtro)
        row_btn_carregar.pack(fill=tk.X, pady=(6, 0))
        self.btn_carregar_datas = ttk.Button(
            row_btn_carregar,
            text="Carregar datas da planilha",
            command=self._carregar_datas_planilha,
        )
        self.btn_carregar_datas.pack(side=tk.LEFT, padx=(0, 8))
        self.label_filtro_status = ttk.Label(
            row_btn_carregar,
            text="Selecione uma planilha e clique em 'Carregar datas da planilha'.",
            style="Card.TLabel",
        )
        self.label_filtro_status.pack(side=tk.LEFT)

        self.frame_opcoes_filtro = ttk.Frame(frame_filtro)
        self.frame_opcoes_filtro.pack(fill=tk.X, pady=(10, 0))
        self._r_todos = ttk.Radiobutton(
            self.frame_opcoes_filtro,
            text="Todos os dias",
            variable=self.filtro_modo,
            value="todos",
            command=self._atualizar_visibilidade_filtro,
        )
        self._r_todos.pack(anchor=tk.W)
        self._r_mes = ttk.Radiobutton(
            self.frame_opcoes_filtro,
            text="Por mês:",
            variable=self.filtro_modo,
            value="mes",
            command=self._atualizar_visibilidade_filtro,
        )
        self._r_mes.pack(anchor=tk.W)
        row_mes = ttk.Frame(self.frame_opcoes_filtro)
        row_mes.pack(anchor=tk.W, padx=(24, 0), pady=(2, 4))
        self.combo_mes = ttk.Combobox(row_mes, state="readonly", width=18)
        self.combo_mes.pack(side=tk.LEFT)
        self._r_datas = ttk.Radiobutton(
            self.frame_opcoes_filtro,
            text="Por datas específicas (marque os dias abaixo):",
            variable=self.filtro_modo,
            value="datas",
            command=self._atualizar_visibilidade_filtro,
        )
        self._r_datas.pack(anchor=tk.W, pady=(4, 0))
        row_lista = ttk.Frame(self.frame_opcoes_filtro)
        row_lista.pack(fill=tk.BOTH, expand=False, padx=(24, 0), pady=(4, 0))
        self.listbox_datas = tk.Listbox(
            row_lista,
            selectmode=tk.EXTENDED,
            height=8,
            font=("Segoe UI", 9),
            exportselection=False,
        )
        scroll = ttk.Scrollbar(row_lista, orient=tk.VERTICAL, command=self.listbox_datas.yview)
        self.listbox_datas.configure(yscrollcommand=scroll.set)
        self.listbox_datas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        row_btns = ttk.Frame(self.frame_opcoes_filtro)
        row_btns.pack(anchor=tk.W, padx=(24, 0), pady=(4, 0))
        ttk.Button(row_btns, text="Selecionar todas", command=self._datas_selecionar_todas).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row_btns, text="Desmarcar todas", command=self._datas_desmarcar_todas).pack(side=tk.LEFT)
        self._atualizar_visibilidade_filtro()

        # Botão que dispara o processamento em thread.
        self.btn_processar = ttk.Button(
            container,
            text="Processar e gerar resultado",
            style="Primario.TButton",
            command=self._iniciar_processamento,
        )
        self.btn_processar.pack(pady=16, ipadx=12, ipady=6)

        # Área onde o resumo e o status aparecem.
        self.frame_resultado = ttk.Frame(container, style="Card.TFrame", padding=16)
        self.frame_resultado.pack(fill=tk.BOTH, expand=True)
        self.label_status = ttk.Label(
            self.frame_resultado,
            text="Selecione uma planilha e clique em Processar.",
            style="Card.TLabel",
            wraplength=500,
        )
        self.label_status.pack(anchor=tk.W)
        self.frame_detalhes = ttk.Frame(self.frame_resultado)
        self.frame_detalhes.pack(anchor=tk.W, pady=(12, 0))
        self.btn_abrir_pasta = ttk.Button(
            self.frame_resultado,
            text="Abrir pasta do resultado",
            command=self._abrir_pasta_resultado,
            state=tk.DISABLED,
        )
        self.btn_abrir_pasta.pack(anchor=tk.W, pady=(8, 0))

    def _selecionar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialdir=SCRIPT_DIR,
        )
        if caminho:
            self.arquivo_planilha.set(caminho)

    def _selecionar_pasta(self):
        caminho = filedialog.askdirectory(title="Pasta de saída", initialdir=SCRIPT_DIR)
        if caminho:
            self.pasta_saida.set(caminho)

    def _carregar_datas_planilha(self):
        arq = self.arquivo_planilha.get().strip()
        if not arq or not Path(arq).exists():
            messagebox.showwarning("Aviso", "Selecione uma planilha válida antes de carregar as datas.")
            return
        if self._carregando_datas:
            return
        self._carregando_datas = True
        self.btn_carregar_datas.configure(state=tk.DISABLED)
        self.label_filtro_status.configure(text="Carregando datas…")
        thread = threading.Thread(target=self._thread_carregar_datas, args=(arq,), daemon=True)
        thread.start()
        self.after(100, self._verificar_carregar_datas, thread)

    def _thread_carregar_datas(self, arquivo):
        try:
            config = carregar_config()
            datas_unicas, meses_anos = extrair_datas_planilha(arquivo, config)
            self.after(0, self._preencher_filtro_datas, datas_unicas, meses_anos, None)
        except Exception as e:
            self.after(0, self._preencher_filtro_datas, [], [], str(e))

    def _verificar_carregar_datas(self, thread):
        if thread.is_alive():
            self.after(200, self._verificar_carregar_datas, thread)
            return
        self._carregando_datas = False
        self.btn_carregar_datas.configure(state=tk.NORMAL)

    def _preencher_filtro_datas(self, datas_unicas, meses_anos, erro):
        self.datas_unicas = datas_unicas or []
        self.meses_anos = meses_anos or []
        self.combo_mes.set("")
        self.listbox_datas.delete(0, tk.END)
        if erro:
            self.label_filtro_status.configure(text=f"Erro ao carregar: {erro}")
            messagebox.showerror("Erro", erro)
            return
        n = len(self.datas_unicas)
        self.label_filtro_status.configure(text=f"{n} data(s) encontrada(s). Escolha como filtrar abaixo.")
        # O combobox usa o formato curto de mês para ficar mais prático na tela.
        MESES = ("Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez")
        valores_mes = [
            f"{MESES[m - 1]}/{a}" for m, a in self.meses_anos
        ]
        self.combo_mes["values"] = valores_mes
        if valores_mes:
            self.combo_mes.current(0)
        # A lista mostra cada data já formatada.
        for d in self.datas_unicas:
            self.listbox_datas.insert(tk.END, d.strftime("%d/%m/%Y"))
        self._atualizar_visibilidade_filtro()

    def _atualizar_visibilidade_filtro(self):
        modo = self.filtro_modo.get()
        self.combo_mes.configure(state="readonly" if modo == "mes" else "disabled")
        self.listbox_datas.configure(state=tk.NORMAL if modo == "datas" else tk.DISABLED)

    def _datas_selecionar_todas(self):
        self.listbox_datas.selection_set(0, tk.END)

    def _datas_desmarcar_todas(self):
        self.listbox_datas.selection_clear(0, tk.END)

    def _obter_datas_selecionadas(self):
        """Lê o filtro atual e devolve as datas que entram no processamento."""
        modo = self.filtro_modo.get()
        if modo == "todos":
            return None
        if modo == "mes":
            idx = self.combo_mes.current()
            if idx < 0 or not self.meses_anos:
                return None
            mes, ano = self.meses_anos[idx]
            return [d for d in self.datas_unicas if d.month == mes and d.year == ano]
        if modo == "datas":
            sel = self.listbox_datas.curselection()
            return [self.datas_unicas[i] for i in sel]
        return None

    def _iniciar_processamento(self):
        if self._em_processamento:
            return
        arq = self.arquivo_planilha.get().strip()
        if not arq:
            messagebox.showwarning("Aviso", "Selecione uma planilha Excel (.xlsx).")
            return
        if not Path(arq).exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{arq}")
            return
        if self.filtro_modo.get() == "datas":
            datas_sel = self._obter_datas_selecionadas()
            if not datas_sel:
                messagebox.showwarning(
                    "Aviso",
                    "Em 'Por datas específicas' selecione ao menos uma data na lista.",
                )
                return
        self._em_processamento = True
        self.btn_processar.configure(state=tk.DISABLED)
        self.label_status.configure(text="Processando… Aguarde.")
        for w in self.frame_detalhes.winfo_children():
            w.destroy()
        self.btn_abrir_pasta.configure(state=tk.DISABLED)
        self.arquivo_gerado = None
        self.totais = None

        pasta = self.pasta_saida.get().strip() or None
        datas = self._obter_datas_selecionadas()
        thread = threading.Thread(
            target=self._processar,
            args=(arq, pasta, datas),
            daemon=True,
        )
        thread.start()
        self.after(100, self._verificar_thread, thread)

    def _processar(self, arquivo, pasta_saida, datas_selecionadas):
        try:
            config = carregar_config()
            out_path, totais = exportar_resultados(
                arquivo, pasta_saida, config, datas_selecionadas=datas_selecionadas
            )
            self.arquivo_gerado = str(out_path)
            self.totais = totais
        except Exception as e:
            self.arquivo_gerado = None
            self.totais = {"erro": str(e)}

    def _verificar_thread(self, thread):
        if thread.is_alive():
            self.after(200, self._verificar_thread, thread)
            return
        self._em_processamento = False
        self.btn_processar.configure(state=tk.NORMAL)
        self._exibir_resultado()

    def _exibir_resultado(self):
        if self.totais is None:
            self.label_status.configure(text="Nenhum resultado.")
            return
        if "erro" in self.totais:
            self.label_status.configure(text=f"Erro: {self.totais['erro']}")
            messagebox.showerror("Erro", self.totais["erro"])
            return

        total_cartas = self.totais.get("total_cartas", 0)
        total_autos = self.totais.get("total_autos", 0)
        self.label_status.configure(
            text=f"Processamento concluído.\nTotal de cartas: {total_cartas}  |  Total de autos (únicos): {total_autos}"
        )

        for w in self.frame_detalhes.winfo_children():
            w.destroy()
        ttk.Label(
            self.frame_detalhes,
            text=f"• Cartas: {total_cartas}\n• Autos (únicos): {total_autos}\n• Arquivo: {Path(self.arquivo_gerado).name}",
            style="Card.TLabel",
            justify=tk.LEFT,
        ).pack(anchor=tk.W)
        if self.totais.get("linhas_data_invalida", 0) > 0:
            self.estilo.configure(
                "Aviso.TLabel",
                background=CORES["card"],
                foreground=CORES["aviso"],
                font=("Segoe UI", 9),
            )
            ttk.Label(
                self.frame_detalhes,
                text=f"⚠ {self.totais['linhas_data_invalida']} linha(s) com data inválida foram ignoradas.",
                style="Aviso.TLabel",
            ).pack(anchor=tk.W, pady=(4, 0))

        self.btn_abrir_pasta.configure(state=tk.NORMAL)
        messagebox.showinfo(
            "Concluído",
            f"Resultado salvo em:\n{self.arquivo_gerado}\n\nTotal de cartas: {total_cartas}\nTotal de autos (únicos): {total_autos}",
        )

    def _abrir_pasta_resultado(self):
        if self.arquivo_gerado:
            abrir_pasta(self.arquivo_gerado)


def main():
    app = AppContagemCartas()
    app.mainloop()


if __name__ == "__main__":
    main()
