# -*- coding: utf-8 -*-
"""Busca Aluguel — janela simples para buscar aluguéis e gerar Excel."""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import motor

TAMANHOS = {"2 1/2": 2, "3 1/2": 3, "4 1/2": 4, "5 1/2": 5, "6 1/2": 6}


def abrir_no_sistema(caminho):
    caminho = str(caminho)
    if sys.platform == "darwin":
        # Planilhas abrem no Numbers (vem em todo Mac); se não houver, no app padrão
        if caminho.endswith(".xlsx") and subprocess.call(
                ["open", "-a", "Numbers", caminho]) == 0:
            return
        subprocess.Popen(["open", caminho])
    elif sys.platform.startswith("win"):
        os.startfile(caminho)
    else:
        subprocess.Popen(["xdg-open", caminho])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Busca Aluguel")
        self.geometry("620x640")
        self.minsize(560, 560)
        self.cfg = motor.carregar_config()
        self.fila = queue.Queue()
        self.parar_flag = threading.Event()
        self.thread = None
        self._montar()
        self.after(150, self._ler_fila)

    # ── interface ──
    def _montar(self):
        pad = {"padx": 10, "pady": 4}
        f = ttk.LabelFrame(self, text="Critérios")
        f.pack(fill="x", padx=12, pady=(12, 6))
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Cidades (separadas por vírgula)").grid(row=0, column=0, sticky="w", **pad)
        self.v_cidades = tk.StringVar(value=self.cfg["cidades"])
        ttk.Entry(f, textvariable=self.v_cidades).grid(row=0, column=1, sticky="ew", **pad)

        ttk.Label(f, text="Preço máximo ($/mês)").grid(row=1, column=0, sticky="w", **pad)
        self.v_preco = tk.StringVar(value=str(self.cfg["preco_max"]))
        ttk.Entry(f, textvariable=self.v_preco, width=10).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(f, text="Tamanho mínimo").grid(row=2, column=0, sticky="w", **pad)
        inv = {v: k for k, v in TAMANHOS.items()}
        self.v_tam = tk.StringVar(value=inv.get(int(self.cfg["min_pecas"]), "4 1/2"))
        ttk.Combobox(f, textvariable=self.v_tam, values=list(TAMANHOS), state="readonly",
                     width=8).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(f, text="Área mínima (pi²)").grid(row=3, column=0, sticky="w", **pad)
        self.v_area = tk.StringVar(value=str(self.cfg["area_min"]))
        ttk.Entry(f, textvariable=self.v_area, width=10).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(f, text="Idade máxima do prédio (anos)").grid(row=4, column=0, sticky="w", **pad)
        self.v_idade = tk.StringVar(value=str(self.cfg["idade_max"]))
        ttk.Entry(f, textvariable=self.v_idade, width=10).grid(row=4, column=1, sticky="w", **pad)

        self.v_garagem = tk.BooleanVar(value=self.cfg["exigir_garagem"])
        ttk.Checkbutton(f, text="Exigir garagem interna", variable=self.v_garagem).grid(
            row=5, column=0, columnspan=2, sticky="w", **pad)

        s = ttk.LabelFrame(self, text="Sites")
        s.pack(fill="x", padx=12, pady=6)
        self.v_sites = {}
        for i, (k, nome) in enumerate(motor.NOMES.items()):
            v = tk.BooleanVar(value=self.cfg["sites"].get(k, True))
            self.v_sites[k] = v
            ttk.Checkbutton(s, text=nome, variable=v).grid(row=0, column=i, padx=10, pady=6)

        b = ttk.Frame(self)
        b.pack(fill="x", padx=12, pady=6)
        self.btn_buscar = ttk.Button(b, text="Buscar", command=self.buscar)
        self.btn_buscar.pack(side="left")
        self.btn_parar = ttk.Button(b, text="Parar", command=self.parar, state="disabled")
        self.btn_parar.pack(side="left", padx=6)
        ttk.Button(b, text="Abrir pasta dos resultados",
                   command=self.abrir_pasta).pack(side="left", padx=6)
        ttk.Button(b, text="Limpar histórico", command=self.limpar).pack(side="right")

        self.log = tk.Text(self, height=14, wrap="word", state="disabled",
                           font=("Menlo" if sys.platform == "darwin" else "Consolas", 11))
        self.log.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        self._escrever("Ajuste os critérios e clique em Buscar.\n"
                       "Uma janela do Chrome vai abrir e percorrer os sites sozinha — não feche.\n"
                       f"Os resultados ficam em: {motor.PASTA}\n")

    # ── ações ──
    def _ler_cfg(self):
        try:
            cfg = dict(self.cfg)
            cfg["cidades"] = self.v_cidades.get().strip()
            cfg["preco_max"] = int(float(self.v_preco.get().replace(",", ".").replace(" ", "")))
            cfg["min_pecas"] = TAMANHOS[self.v_tam.get()]
            cfg["area_min"] = int(float(self.v_area.get().replace(" ", "") or 0))
            cfg["idade_max"] = int(float(self.v_idade.get().replace(" ", "") or 100))
            cfg["exigir_garagem"] = bool(self.v_garagem.get())
            cfg["sites"] = {k: bool(v.get()) for k, v in self.v_sites.items()}
        except ValueError:
            messagebox.showerror("Busca Aluguel", "Preço, área e idade precisam ser números.")
            return None
        if not motor.lista_cidades(cfg):
            messagebox.showerror("Busca Aluguel", "Informe pelo menos uma cidade.")
            return None
        if not any(cfg["sites"].values()):
            messagebox.showerror("Busca Aluguel", "Marque pelo menos um site.")
            return None
        return cfg

    def buscar(self):
        cfg = self._ler_cfg()
        if not cfg:
            return
        self.cfg = cfg
        motor.salvar_config(cfg)
        self.parar_flag.clear()
        self.btn_buscar.config(state="disabled")
        self.btn_parar.config(state="normal")
        self._escrever("\n— Iniciando busca —\n")
        self.thread = threading.Thread(target=self._rodar, daemon=True)
        self.thread.start()

    def _rodar(self):
        try:
            arq, _ = motor.executar(self.cfg, lambda m: self.fila.put(("log", m)),
                                    self.parar_flag.is_set)
            self.fila.put(("fim", arq))
        except Exception as e:
            self.fila.put(("erro", str(e)))

    def parar(self):
        self.parar_flag.set()
        self._escrever("Parando... (termina o anúncio atual)\n")

    def abrir_pasta(self):
        motor.PASTA.mkdir(parents=True, exist_ok=True)
        abrir_no_sistema(motor.PASTA)

    def limpar(self):
        if messagebox.askyesno("Busca Aluguel",
                               "Apagar o histórico? Na próxima busca todos os anúncios "
                               "voltam a aparecer."):
            if motor.HISTORICO.exists():
                motor.HISTORICO.unlink()
            self._escrever("Histórico apagado.\n")

    # ── fila de mensagens da busca ──
    def _ler_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(dado + "\n")
                else:
                    self.btn_buscar.config(state="normal")
                    self.btn_parar.config(state="disabled")
                    if tipo == "erro":
                        self._escrever(f"\nERRO: {dado}\n")
                        messagebox.showerror("Busca Aluguel", dado)
                    elif dado:
                        if messagebox.askyesno("Busca Aluguel",
                                               ("Busca concluída! Abrir a planilha no Numbers agora?" if sys.platform == "darwin" else "Busca concluída! Abrir a planilha agora?")):
                            abrir_no_sistema(dado)
        except queue.Empty:
            pass
        self.after(150, self._ler_fila)

    def _escrever(self, txt):
        self.log.config(state="normal")
        self.log.insert("end", txt)
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
