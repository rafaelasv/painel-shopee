"""
Resumo da Loja — Shopee
========================
Instalar dependências:
    pip install customtkinter pandas openpyxl matplotlib

Gerar .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed painel_shopee.py
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import Counter
from datetime import datetime
import pandas as pd
import os

# ── Matplotlib (gráfico) ──────────────────────────────────────────────────────
try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# ══════════════════════════════════════════════════════════════════════════════
# TEMA
# ══════════════════════════════════════════════════════════════════════════════

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

C = {
    "laranja":     "#EE4D2D",
    "laranja_hov": "#C43D20",
    "laranja_clr": "#FFF3F0",
    "laranja_brd": "#F8C1B5",
    "fundo":       "#F8F8F8",
    "branco":      "#FFFFFF",
    "texto":       "#2D2D2D",
    "texto_sec":   "#888888",
    "borda":       "#EEEEEE",
    "borda_med":   "#DDDDDD",
    "verde":       "#1A8F5A",
    "vermelho":    "#E05252",
    "verm_fnd":    "#FDECEA",
    "amarelo":     "#B7860D",
    "amar_fnd":    "#FFFBEA",
    "sidebar":     "#FFFFFF",
    "sidebar_ativ":"#FFF3F0",
    "cinza_lbl":   "#999999",
}

FONTE = "Segoe UI" if os.name == "nt" else "Helvetica Neue"

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA DE NEGÓCIO
# ══════════════════════════════════════════════════════════════════════════════

def brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return "R$ 0,00"

def parse_valor(v):
    if pd.isna(v): return 0.0
    s = str(v).replace("R$","").replace(".","").replace(",",".").strip()
    try: return float(s)
    except: return 0.0

def get_faixa(preco, frete_gratis):
    if preco < 8:   return {"num":1,"comissao":0.50,"fixa":0.0}
    if preco < 80:  return {"num":2,"comissao":0.20 if frete_gratis else 0.14,"fixa":4.0}
    if preco < 100: return {"num":3,"comissao":0.14,"fixa":16.0}
    if preco < 200: return {"num":4,"comissao":0.14,"fixa":20.0}
    return             {"num":5,"comissao":0.14,"fixa":26.0}

def calcular_taxas(preco, frete_gratis, conta_alto):
    f = get_faixa(preco, frete_gratis)
    comissao_val = preco * f["comissao"]
    fixa_val = f["fixa"]
    if conta_alto and f["num"] >= 2: fixa_val += 3.0
    total = comissao_val + fixa_val
    return {"faixa":f["num"],"comissao_pct":f["comissao"]*100,
            "comissao_val":comissao_val,"fixa_val":fixa_val,
            "total_taxas":total,"liquido":preco-total}

def ler_pedidos(path):
    df = pd.read_excel(path, engine="openpyxl")
    receita  = pd.to_numeric(df["Subtotal do produto"], errors="coerce").sum()
    comissao = pd.to_numeric(df["Taxa de comissão bruta"], errors="coerce").sum()
    servico  = pd.to_numeric(df["Taxa de serviço bruta"], errors="coerce").sum()
    n = len(df)
    df["_data"] = pd.to_datetime(df["Data de criação do pedido"], errors="coerce").dt.date
    vendas_dia = df.groupby("_data")["Subtotal do produto"].apply(
        lambda x: pd.to_numeric(x, errors="coerce").sum()
    ).sort_index().tail(7)
    return {
        "total_pedidos": n,
        "receita_bruta": receita,
        "comissao": comissao,
        "taxa_servico": servico,
        "total_liquido": receita - comissao - servico,
        "ticket_medio": receita / n if n else 0,
        "top_produtos": df["Nome do Produto"].value_counts().head(5),
        "variacoes": df["Nome da variação"].dropna().value_counts().head(5),
        "vendas_dia": vendas_dia,
    }

def ler_devolucoes(path):
    df = pd.read_excel(path, engine="openpyxl")
    MOTIVOS = {
        "não recebeu":       "Não recebeu o pedido",
        "produto(s) errado": "Produto errado (tamanho/cor/item)",
        "mudei de ideia":    "Mudou de ideia",
        "defeito":           "Produto com defeito",
        "falso":             "Produto falso",
        "danificado":        "Produto danificado",
        "não corresponde":   "Diferente da descrição",
    }
    def resumir(m):
        ml = str(m).lower()
        for k, v in MOTIVOS.items():
            if k in ml: return v
        palavras = str(m).split()
        return " ".join(palavras[:5]) + ("..." if len(palavras)>5 else "")
    motivos = Counter([resumir(m) for m in df["Motivo da Devolução"].dropna()])
    reemb = df["Quantia total de reembolsos"].apply(parse_valor).sum()
    historico = df[["ID da Devolução","Tempo de envio de devolução",
                    "Motivo da Devolução","Quantia total de reembolsos"]].head(5)
    return {
        "total": len(df),
        "status": df["Status da Devolução / Reembolso"].value_counts().to_dict(),
        "motivos": motivos,
        "total_reembolsado": reemb,
        "historico": historico,
    }

def ler_financeiro(path):
    df = pd.read_excel(path, engine="openpyxl", header=None)
    mapa = {
        "1. Receita Total":             "receita_total",
        "Valor do Reembolso":           "valor_reembolso",
        "2. Despesas Totais":           "despesas_totais",
        "Taxa de comissão líquida":     "comissao_liq",
        "Taxa de serviço líquida":      "servico_liq",
        "3. Quantidade Total Liberada": "total_liberado",
    }
    dados = {}
    for _, row in df.iterrows():
        chave = str(row.iloc[0]).strip()
        for k, var in mapa.items():
            if k in chave:
                try: dados[var] = float(str(row.iloc[1]).replace(",","."))
                except: dados[var] = 0.0
    return dados

# ══════════════════════════════════════════════════════════════════════════════
# COMPONENTES
# ══════════════════════════════════════════════════════════════════════════════

def card(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=C["branco"], corner_radius=10,
                        border_width=1, border_color=C["borda"], **kw)

def lbl_hint(parent, text):
    return ctk.CTkLabel(parent, text=text.upper(), font=(FONTE,11),
                        text_color=C["cinza_lbl"])

def lbl_valor(parent, text, size=26, cor=None):
    return ctk.CTkLabel(parent, text=text, font=(FONTE,size,"bold"),
                        text_color=cor or C["texto"])

def sep(parent):
    ctk.CTkFrame(parent, height=1, fg_color=C["borda"],
                 corner_radius=0).pack(fill="x", padx=16, pady=6)

def metrica(parent, titulo, valor, cor=None):
    f = card(parent)
    lbl_hint(f, titulo).pack(anchor="w", padx=16, pady=(14,2))
    lbl_valor(f, valor, cor=cor).pack(anchor="w", padx=16, pady=(0,14))
    return f

def btn_import(parent, text, cmd):
    return ctk.CTkButton(parent, text=text, command=cmd,
                         fg_color=C["branco"], text_color=C["texto_sec"],
                         hover_color=C["fundo"], border_width=1,
                         border_color=C["borda_med"], corner_radius=8,
                         font=(FONTE,11), height=34, width=160)

# ══════════════════════════════════════════════════════════════════════════════
# APP
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Resumo da Loja — Shopee")
        self.geometry("1100x720")
        self.minsize(900,580)
        self.configure(fg_color=C["fundo"])
        self.dados_pedidos    = None
        self.dados_devolucoes = None
        self.dados_financeiro = None
        self.path_ped  = ctk.StringVar()
        self.path_dev  = ctk.StringVar()
        self.path_fin  = ctk.StringVar()
        self._aba = "visao"
        self._sidebar()
        self.content = ctk.CTkFrame(self, fg_color=C["fundo"], corner_radius=0)
        self.content.pack(side="left", fill="both", expand=True)
        self._ir("visao")

    # ── sidebar ───────────────────────────────────────────────────────────────
    def _sidebar(self):
        sb = ctk.CTkFrame(self, width=220, fg_color=C["sidebar"],
                          corner_radius=0, border_width=1,
                          border_color=C["borda"])
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        logo = ctk.CTkFrame(sb, fg_color=C["sidebar"], corner_radius=0)
        logo.pack(fill="x", padx=20, pady=(24,20))
        ic = ctk.CTkFrame(logo, width=36, height=36,
                          fg_color=C["laranja"], corner_radius=8)
        ic.pack(side="left"); ic.pack_propagate(False)
        ctk.CTkLabel(ic, text="S", font=(FONTE,16,"bold"),
                     text_color=C["branco"]).place(relx=.5,rely=.5,anchor="center")
        tx = ctk.CTkFrame(logo, fg_color=C["sidebar"], corner_radius=0)
        tx.pack(side="left", padx=(10,0))
        ctk.CTkLabel(tx, text="Resumo da Loja",
                     font=(FONTE,13,"bold"), text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(tx, text="PARA VENDEDORES",
                     font=(FONTE,8), text_color=C["texto_sec"]).pack(anchor="w")

        self._nav = {}
        for key, lbl in [("visao","📊  Visão Geral"),("calc","🧮  Calculadora"),
                          ("dev","↩   Devoluções"),("resumo","💬  Resumo WhatsApp")]:
            b = ctk.CTkButton(sb, text=lbl, anchor="w",
                              font=(FONTE,13), height=42,
                              fg_color=C["sidebar"], text_color=C["texto_sec"],
                              hover_color=C["fundo"], corner_radius=8,
                              command=lambda k=key: self._ir(k))
            b.pack(fill="x", padx=12, pady=2)
            self._nav[key] = b

        rod = ctk.CTkFrame(sb, fg_color=C["sidebar"], corner_radius=0)
        rod.pack(side="bottom", fill="x", padx=16, pady=16)
        ctk.CTkFrame(rod, height=1, fg_color=C["borda"],
                     corner_radius=0).pack(fill="x", pady=(0,10))
        ctk.CTkLabel(rod, text="Loja Shopee",
                     font=(FONTE,12,"bold"), text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(rod, text="Painel v1.0",
                     font=(FONTE,10), text_color=C["texto_sec"]).pack(anchor="w")

    def _nav_ativa(self, key):
        for k, b in self._nav.items():
            if k == key:
                b.configure(fg_color=C["sidebar_ativ"],
                            text_color=C["laranja"],
                            font=(FONTE,13,"bold"))
            else:
                b.configure(fg_color=C["sidebar"],
                            text_color=C["texto_sec"],
                            font=(FONTE,13))

    def _ir(self, key):
        self._aba = key
        self._nav_ativa(key)
        for w in self.content.winfo_children(): w.destroy()
        {"visao":self._visao,"calc":self._calc,
         "dev":self._dev,"resumo":self._resumo}[key]()

    # ══════════════════════════════════════════════════════════════════════════
    # VISÃO GERAL
    # ══════════════════════════════════════════════════════════════════════════
    def _visao(self):
        p = self.content
        topo = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        topo.pack(fill="x", padx=28, pady=(24,0))
        ctk.CTkLabel(topo, text="Visão Geral",
                     font=(FONTE,22,"bold"), text_color=C["texto"]).pack(side="left")
        bf = ctk.CTkFrame(topo, fg_color=C["fundo"], corner_radius=0)
        bf.pack(side="right")
        for t, v, c in [("Importar Pedidos",self.path_ped,self._load_ped),
                         ("Importar Devoluções",self.path_dev,self._load_dev),
                         ("Importar Financeiro",self.path_fin,self._load_fin)]:
            btn_import(bf, t, c).pack(side="left", padx=4)

        sc = ctk.CTkScrollableFrame(p, fg_color=C["fundo"], corner_radius=0)
        sc.pack(fill="both", expand=True, padx=20, pady=16)

        if not self.dados_pedidos:
            ctk.CTkLabel(sc,
                         text="Importe os arquivos acima para ver o resumo do mês.",
                         font=(FONTE,13), text_color=C["texto_sec"]).pack(pady=60)
            return

        d = self.dados_pedidos

        # métricas principais
        r1 = ctk.CTkFrame(sc, fg_color=C["fundo"], corner_radius=0)
        r1.pack(fill="x", pady=(0,12))
        r1.grid_columnconfigure((0,1,2,3), weight=1)
        for i,(t,v,c) in enumerate([
            ("Pedidos no Mês",   str(d["total_pedidos"]),             C["texto"]),
            ("Receita Bruta",    brl(d["receita_bruta"]),             C["laranja"]),
            ("Taxas Shopee",     brl(d["comissao"]+d["taxa_servico"]),C["vermelho"]),
            ("Líquido Estimado", brl(d["total_liquido"]),             C["verde"]),
        ]):
            metrica(r1, t, v, cor=c).grid(row=0, column=i, padx=6, sticky="ew")

        # ticket + repasse
        r2 = ctk.CTkFrame(sc, fg_color=C["fundo"], corner_radius=0)
        r2.pack(fill="x", pady=(0,12))
        r2.grid_columnconfigure((0,1), weight=1)
        metrica(r2, "Ticket Médio", brl(d["ticket_medio"])).grid(
            row=0, column=0, padx=6, sticky="ew")
        if self.dados_financeiro:
            tl = self.dados_financeiro.get("total_liberado",0)
            metrica(r2, "Repasse Líquido Shopee", brl(tl),
                    cor=C["verde"]).grid(row=0, column=1, padx=6, sticky="ew")

        # produtos + variações
        r3 = ctk.CTkFrame(sc, fg_color=C["fundo"], corner_radius=0)
        r3.pack(fill="x", pady=(0,12))
        r3.grid_columnconfigure((0,1), weight=1)
        for col,(titulo,dados) in enumerate([
            ("Produtos Mais Vendidos", d["top_produtos"]),
            ("Variações Mais Vendidas", d["variacoes"]),
        ]):
            c = card(r3)
            c.grid(row=0, column=col, padx=6, sticky="nsew")
            lbl_hint(c, titulo).pack(anchor="w", padx=16, pady=(14,8))
            sep(c)
            for nome, qtd in dados.items():
                ln = ctk.CTkFrame(c, fg_color=C["branco"], corner_radius=0)
                ln.pack(fill="x", padx=16, pady=3)
                nc = str(nome)[:44]+"…" if len(str(nome))>44 else str(nome)
                ctk.CTkLabel(ln, text=nc, font=(FONTE,12),
                             text_color=C["texto"], anchor="w").pack(side="left")
                ctk.CTkLabel(ln, text=f"{qtd}x", font=(FONTE,12,"bold"),
                             text_color=C["laranja"], anchor="e").pack(side="right")
            ctk.CTkFrame(c, height=10, fg_color=C["branco"],
                         corner_radius=0).pack()

        # gráfico
        if HAS_MATPLOTLIB and not d["vendas_dia"].empty:
            gc = card(sc)
            gc.pack(fill="x", pady=(0,16))
            lbl_hint(gc, "Fluxo de Vendas Diário — Últimos 7 dias").pack(
                anchor="w", padx=16, pady=(14,4))
            sep(gc)
            vd = d["vendas_dia"]
            labels  = [str(dt.strftime("%d/%m") if hasattr(dt,"strftime") else dt)
                       for dt in vd.index]
            valores = list(vd.values)
            cores   = [C["laranja"] if v==max(valores) else "#F2C5BB" for v in valores]
            fig = Figure(figsize=(8,2.4), dpi=96, facecolor=C["branco"])
            ax  = fig.add_subplot(111)
            ax.set_facecolor(C["branco"])
            ax.bar(labels, valores, color=cores, width=0.55)
            ax.set_ylabel("R$", fontsize=9, color=C["texto_sec"])
            ax.tick_params(colors=C["texto_sec"], labelsize=9)
            ax.spines[["top","right","left"]].set_visible(False)
            ax.spines["bottom"].set_color(C["borda"])
            ax.yaxis.set_tick_params(length=0)
            fig.tight_layout(pad=1.2)
            cv = FigureCanvasTkAgg(fig, master=gc)
            cv.draw()
            cv.get_tk_widget().pack(fill="x", padx=16, pady=(0,14))

    # ══════════════════════════════════════════════════════════════════════════
    # CALCULADORA
    # ══════════════════════════════════════════════════════════════════════════
    def _calc(self):
        p = self.content
        topo = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        topo.pack(fill="x", padx=28, pady=(24,12))
        ctk.CTkLabel(topo, text="SIMULADOR E FINANCEIRO",
                     font=(FONTE,11), text_color=C["laranja"]).pack(anchor="w")
        ctk.CTkLabel(topo, text="Calculadora de Taxas",
                     font=(FONTE,22,"bold"), text_color=C["texto"]).pack(anchor="w")

        corpo = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        corpo.pack(fill="both", expand=True, padx=20, pady=(0,16))
        corpo.grid_columnconfigure(0, weight=3)
        corpo.grid_columnconfigure(1, weight=2)
        corpo.grid_rowconfigure(0, weight=1)

        # ── esquerda (sem scroll) ─────────────────────────────────────────────
        esq = ctk.CTkFrame(corpo, fg_color=C["fundo"], corner_radius=0)
        esq.grid(row=0, column=0, sticky="nsew", padx=(0,8))

        # configurações em linha horizontal
        cfg = card(esq)
        cfg.pack(fill="x", pady=(0,10))
        lbl_hint(cfg, "Configurações de Canal").pack(anchor="w", padx=16, pady=(12,6))
        sep(cfg)

        self.var_frete = ctk.BooleanVar(value=True)
        self.var_alto  = ctk.BooleanVar(value=False)

        linha_tog = ctk.CTkFrame(cfg, fg_color=C["branco"], corner_radius=0)
        linha_tog.pack(fill="x", padx=16, pady=(4,12))
        linha_tog.grid_columnconfigure((0,1), weight=1)

        for col, (var, titulo, sub) in enumerate([
            (self.var_frete, "Programa Frete Grátis", "Comissão de 20% sobre a venda"),
            (self.var_alto,  "Vendedor CPF",           "+450 pedidos/90 dias (+R$3/item)"),
        ]):
            blk = ctk.CTkFrame(linha_tog, fg_color=C["fundo"], corner_radius=8)
            blk.grid(row=0, column=col, padx=(0,8) if col==0 else (8,0), sticky="ew")
            top_row = ctk.CTkFrame(blk, fg_color=C["fundo"], corner_radius=0)
            top_row.pack(fill="x", padx=10, pady=(8,0))
            ctk.CTkLabel(top_row, text=titulo, font=(FONTE,13,"bold"),
                         text_color=C["texto"]).pack(side="left")
            ctk.CTkSwitch(top_row, variable=var, text="",
                          progress_color=C["laranja"],
                          button_color=C["branco"],
                          button_hover_color=C["fundo"],
                          width=44, height=22,
                          command=self._calc_render).pack(side="right")
            ctk.CTkLabel(blk, text=sub, font=(FONTE,11),
                         text_color=C["texto_sec"]).pack(anchor="w", padx=10, pady=(2,8))

        # tabela de produtos
        tab = card(esq)
        tab.pack(fill="both", expand=True, pady=(0,10))
        lbl_hint(tab, "Itens para Cálculo").pack(anchor="w", padx=16, pady=(12,4))
        sep(tab)

        hdr = ctk.CTkFrame(tab, fg_color=C["branco"], corner_radius=0)
        hdr.pack(fill="x", padx=16, pady=(4,0))
        hdr.grid_columnconfigure(0, weight=3); hdr.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(hdr, text="NOME DO PRODUTO", font=(FONTE,10),
                     text_color=C["cinza_lbl"]).grid(row=0,column=0,sticky="w")
        ctk.CTkLabel(hdr, text="PREÇO (R$)", font=(FONTE,10),
                     text_color=C["cinza_lbl"]).grid(row=0,column=1,sticky="w",padx=(8,0))

        self.prod_frame = ctk.CTkFrame(tab, fg_color=C["branco"], corner_radius=0)
        self.prod_frame.pack(fill="x", padx=16)
        self.produtos_vars = []
        for _ in range(3): self._add_prod()

        btn_row = ctk.CTkFrame(tab, fg_color=C["branco"], corner_radius=0)
        btn_row.pack(fill="x", padx=16, pady=(4,12))
        btn_row.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btn_row, text="+ Adicionar linha",
                      fg_color=C["branco"], text_color=C["laranja"],
                      hover_color=C["laranja_clr"], font=(FONTE,12,"bold"),
                      corner_radius=8, height=36, border_width=1,
                      border_color=C["laranja_brd"],
                      command=self._add_prod).grid(row=0,column=0,padx=(0,4),sticky="ew")

        ctk.CTkButton(btn_row, text="Calcular",
                      fg_color=C["laranja"], hover_color=C["laranja_hov"],
                      text_color=C["branco"], font=(FONTE,12,"bold"),
                      corner_radius=8, height=36,
                      command=self._calc_render).grid(row=0,column=1,padx=(4,0),sticky="ew")

        # ── direita (resultados, com scroll) ─────────────────────────────────
        self.res_sc = ctk.CTkScrollableFrame(corpo, fg_color=C["fundo"],
                                              corner_radius=0)
        self.res_sc.grid(row=0, column=1, sticky="nsew", padx=(8,0))
        self._calc_render()

    def _add_prod(self):
        vn = ctk.StringVar(); vp = ctk.StringVar()
        self.produtos_vars.append((vn, vp))
        rw = ctk.CTkFrame(self.prod_frame, fg_color=C["branco"], corner_radius=0)
        rw.pack(fill="x", pady=3)
        rw.grid_columnconfigure(0, weight=3); rw.grid_columnconfigure(1, weight=1)
        e1 = ctk.CTkEntry(rw, textvariable=vn, placeholder_text="Nome do produto",
                          font=(FONTE,13), height=32,
                          fg_color=C["fundo"], border_color=C["borda"],
                          text_color=C["texto"])
        e1.grid(row=0, column=0, sticky="ew")
        e1.bind("<KeyRelease>", lambda e: self._calc_render())
        e2 = ctk.CTkEntry(rw, textvariable=vp, placeholder_text="0,00",
                          font=(FONTE,13), height=32, width=100,
                          fg_color=C["fundo"], border_color=C["borda"],
                          text_color=C["texto"])
        e2.grid(row=0, column=1, padx=(8,0), sticky="ew")
        e2.bind("<KeyRelease>", lambda e: self._calc_render())
        idx = len(self.produtos_vars)-1
        ctk.CTkButton(rw, text="×", width=26, height=26,
                      fg_color=C["fundo"], text_color=C["texto_sec"],
                      hover_color=C["verm_fnd"], corner_radius=6,
                      command=lambda r=rw, i=idx: self._rem_prod(r,i)
                      ).grid(row=0, column=2, padx=(4,0))

    def _rem_prod(self, rw, idx):
        if len(self.produtos_vars) <= 1: return
        rw.destroy()
        if idx < len(self.produtos_vars):
            self.produtos_vars.pop(idx)
        self._calc_render()

    def _calc_render(self):
        for w in self.res_sc.winfo_children(): w.destroy()
        lbl_hint(self.res_sc, "Simulação de Repasse").pack(
            anchor="w", padx=4, pady=(0,8))
        frete = self.var_frete.get(); alto = self.var_alto.get()
        algum = False
        for vn, vp in self.produtos_vars:
            nome = vn.get().strip() or "Produto"
            try: preco = float(vp.get().strip().replace(",","."))
            except: continue
            if preco <= 0: continue
            algum = True
            c = calcular_taxas(preco, frete, alto)
            bl = card(self.res_sc)
            bl.pack(fill="x", pady=6, padx=4)
            cab = ctk.CTkFrame(bl, fg_color=C["branco"], corner_radius=0)
            cab.pack(fill="x", padx=14, pady=(12,4))
            ctk.CTkLabel(cab, text=nome, font=(FONTE,13,"bold"),
                         text_color=C["texto"]).pack(side="left")
            ctk.CTkLabel(cab, text=f"FAIXA {c['faixa']}",
                         font=(FONTE,9,"bold"), text_color=C["branco"],
                         fg_color=C["laranja"], corner_radius=4,
                         padx=8, pady=3).pack(side="right")
            ctk.CTkLabel(bl, text=f"Venda: {brl(preco)}",
                         font=(FONTE,11), text_color=C["texto_sec"]).pack(
                             anchor="w", padx=14, pady=(0,8))
            sep(bl)
            for descr, val, cor in [
                (f"Comissão ({c['comissao_pct']:.0f}%)", f"- {brl(c['comissao_val'])}", C["vermelho"]),
                ("Tarifa fixa",                           f"- {brl(c['fixa_val'])}",     C["vermelho"]),
                ("Total de taxas",                        f"- {brl(c['total_taxas'])}",  C["vermelho"]),
            ]:
                r = ctk.CTkFrame(bl, fg_color=C["branco"], corner_radius=0)
                r.pack(fill="x", padx=14, pady=2)
                ctk.CTkLabel(r, text=descr, font=(FONTE,12),
                             text_color=C["texto_sec"]).pack(side="left")
                ctk.CTkLabel(r, text=val, font=(FONTE,12),
                             text_color=cor).pack(side="right")
            sep(bl)
            rl = ctk.CTkFrame(bl, fg_color=C["branco"], corner_radius=0)
            rl.pack(fill="x", padx=14, pady=(4,14))
            ctk.CTkLabel(rl, text="LÍQUIDO RECEBIDO", font=(FONTE,9),
                         text_color=C["cinza_lbl"]).pack(side="left")
            ctk.CTkLabel(rl, text=brl(c["liquido"]), font=(FONTE,20,"bold"),
                         text_color=C["verde"]).pack(side="right")
        if not algum:
            ctk.CTkLabel(self.res_sc,
                         text="Preencha nome e preço para ver o resultado.",
                         font=(FONTE,12), text_color=C["texto_sec"]).pack(pady=24)
        av = ctk.CTkFrame(self.res_sc, fg_color=C["fundo"], corner_radius=8,
                          border_width=1, border_color=C["borda"])
        av.pack(fill="x", padx=4, pady=8)
        ctk.CTkLabel(av,
                     text="ℹ  Valores estimados com base nas taxas Shopee de março/2026.",
                     font=(FONTE,10), text_color=C["texto_sec"],
                     wraplength=280, justify="left").pack(padx=12, pady=10)

    # ══════════════════════════════════════════════════════════════════════════
    # DEVOLUÇÕES
    # ══════════════════════════════════════════════════════════════════════════
    def _dev(self):
        p = self.content
        topo = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        topo.pack(fill="x", padx=28, pady=(24,16))
        ctk.CTkLabel(topo, text="PAINEL DE CONTROLE",
                     font=(FONTE,10), text_color=C["laranja"]).pack(anchor="w")
        ctk.CTkLabel(topo, text="Gestão de Devoluções",
                     font=(FONTE,22,"bold"), text_color=C["texto"]).pack(anchor="w")
        ctk.CTkLabel(topo,
                     text="Monitore e gerencie solicitações de reembolso.",
                     font=(FONTE,11), text_color=C["texto_sec"]).pack(anchor="w")

        if not self.dados_devolucoes:
            ctk.CTkLabel(p,
                         text='Importe o "Relatório de Devoluções" na aba Visão Geral.',
                         font=(FONTE,13), text_color=C["texto_sec"]).pack(pady=60)
            return

        sc = ctk.CTkScrollableFrame(p, fg_color=C["fundo"], corner_radius=0)
        sc.pack(fill="both", expand=True, padx=20, pady=(0,16))
        d = self.dados_devolucoes

        # métricas
        r1 = ctk.CTkFrame(sc, fg_color=C["fundo"], corner_radius=0)
        r1.pack(fill="x", pady=(0,12))
        r1.grid_columnconfigure((0,1), weight=1)

        m1 = card(r1); m1.grid(row=0,column=0,padx=6,sticky="ew")
        lbl_hint(m1, "Total de Solicitações").pack(anchor="w", padx=16, pady=(14,4))
        rn = ctk.CTkFrame(m1, fg_color=C["branco"], corner_radius=0)
        rn.pack(anchor="w", padx=16, pady=(0,14))
        ctk.CTkLabel(rn, text=str(d["total"]), font=(FONTE,32,"bold"),
                     text_color=C["texto"]).pack(side="left")
        ctk.CTkLabel(rn, text=" Pedidos", font=(FONTE,13),
                     text_color=C["texto_sec"]).pack(side="left", pady=(10,0))

        m2 = card(r1); m2.grid(row=0,column=1,padx=6,sticky="ew")
        lbl_hint(m2, "Total Reembolsado").pack(anchor="w", padx=16, pady=(14,4))
        lbl_valor(m2, brl(d["total_reembolsado"]), size=28,
                  cor=C["vermelho"]).pack(anchor="w", padx=16, pady=(0,14))

        # status + motivos
        r2 = ctk.CTkFrame(sc, fg_color=C["fundo"], corner_radius=0)
        r2.pack(fill="x", pady=(0,12))
        r2.grid_columnconfigure((0,1), weight=1)

        ICONES = {
            "Aprovada":                                   ("✅", C["verde"]),
            "Solicitação aprovada":                       ("✅", C["verde"]),
            "Em devolução":                               ("🔄", C["laranja"]),
            "Solicitação cancelada pelo comprador":       ("🚫", C["vermelho"]),
            "Retornado - Pendente validação do vendedor": ("⏳", C["amarelo"]),
        }

        cs = card(r2); cs.grid(row=0,column=0,padx=6,sticky="nsew")
        lbl_hint(cs, "Status das Solicitações").pack(anchor="w", padx=16, pady=(14,4))
        sep(cs)
        for status, qtd in d["status"].items():
            ic, cor = ICONES.get(status, ("•", C["texto_sec"]))
            ln = ctk.CTkFrame(cs, fg_color=C["branco"], corner_radius=0)
            ln.pack(fill="x", padx=16, pady=6)
            ctk.CTkLabel(ln, text=f"{ic}  {status}", font=(FONTE,12),
                         text_color=C["texto"], anchor="w").pack(side="left")
            ctk.CTkLabel(ln, text=str(qtd), font=(FONTE,13,"bold"),
                         text_color=cor, anchor="e").pack(side="right")
        ctk.CTkFrame(cs, height=10, fg_color=C["branco"], corner_radius=0).pack()

        cm = card(r2); cm.grid(row=0,column=1,padx=6,sticky="nsew")
        lbl_hint(cm, "Principais Motivos").pack(anchor="w", padx=16, pady=(14,4))
        sep(cm)
        for motivo, qtd in d["motivos"].most_common(6):
            ln = ctk.CTkFrame(cm, fg_color=C["branco"], corner_radius=0)
            ln.pack(fill="x", padx=16, pady=4)
            ctk.CTkLabel(ln, text=motivo, font=(FONTE,12),
                         text_color=C["texto"], anchor="w").pack(side="left")
            ctk.CTkLabel(ln, text=f"{qtd:02d}", font=(FONTE,13,"bold"),
                         text_color=C["vermelho"], anchor="e").pack(side="right")
        sep(cm)
        df_box = ctk.CTkFrame(cm, fg_color=C["amar_fnd"], corner_radius=8)
        df_box.pack(fill="x", padx=16, pady=(4,14))
        ctk.CTkLabel(df_box, text="💡  DICAS DE PERFORMANCE",
                     font=(FONTE,9,"bold"), text_color=C["amarelo"]).pack(
                         anchor="w", padx=12, pady=(10,4))
        DICAS = {
            "não recebeu": "📮 Verificar atrasos com a transportadora.",
            "errado":      "📸 Conferir fotos e descrição dos anúncios.",
            "defeito":     "🔍 Revisar estoque e embalagem antes do envio.",
            "diferente":   "📸 Atualizar as fotos do produto no anúncio.",
        }
        vis = set()
        for motivo, _ in d["motivos"].most_common():
            for ch, dc in DICAS.items():
                if ch in motivo.lower() and dc not in vis:
                    vis.add(dc)
                    ctk.CTkLabel(df_box, text=dc, font=(FONTE,11),
                                 text_color=C["amarelo"],
                                 wraplength=280, justify="left").pack(
                                     anchor="w", padx=12, pady=2)
        ctk.CTkFrame(df_box, height=8, fg_color=C["amar_fnd"],
                     corner_radius=0).pack()

        # histórico
        hist = d["historico"]
        if not hist.empty:
            ch_card = card(sc); ch_card.pack(fill="x", pady=(0,12))
            hr = ctk.CTkFrame(ch_card, fg_color=C["branco"], corner_radius=0)
            hr.pack(fill="x", padx=16, pady=(14,4))
            lbl_hint(hr, "Histórico Recente").pack(side="left")
            ctk.CTkLabel(hr, text="VER RELATÓRIO COMPLETO",
                         font=(FONTE,10,"bold"),
                         text_color=C["laranja"]).pack(side="right")
            sep(ch_card)
            cols_k = ["ID da Devolução","Tempo de envio de devolução",
                      "Motivo da Devolução","Quantia total de reembolsos"]
            hdrs   = ["ID DO PEDIDO","DATA","MOTIVO","VALOR"]
            hrow = ctk.CTkFrame(ch_card, fg_color=C["fundo"], corner_radius=0)
            hrow.pack(fill="x", padx=16, pady=4)
            for h in hdrs:
                ctk.CTkLabel(hrow, text=h, font=(FONTE,9),
                             text_color=C["cinza_lbl"],
                             width=150, anchor="w").pack(side="left", padx=4)
            for _, row in hist.iterrows():
                dr = ctk.CTkFrame(ch_card, fg_color=C["branco"], corner_radius=0)
                dr.pack(fill="x", padx=16, pady=3)
                for ck in cols_k:
                    v = str(row[ck]) if not pd.isna(row[ck]) else "—"
                    if ck=="Motivo da Devolução" and len(v)>28: v=v[:26]+"…"
                    ctk.CTkLabel(dr, text=v, font=(FONTE,11),
                                 text_color=C["texto"],
                                 width=150, anchor="w").pack(side="left", padx=4)
            ctk.CTkFrame(ch_card, height=10, fg_color=C["branco"],
                         corner_radius=0).pack()

    # ══════════════════════════════════════════════════════════════════════════
    # RESUMO WHATSAPP
    # ══════════════════════════════════════════════════════════════════════════
    def _resumo(self):
        p = self.content
        topo = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        topo.pack(fill="x", padx=28, pady=(24,16))
        ctk.CTkLabel(topo, text="↗  RELATÓRIO INSTANTÂNEO",
                     font=(FONTE,10), text_color=C["laranja"]).pack(anchor="w")
        ctk.CTkLabel(topo, text="Resumo para WhatsApp",
                     font=(FONTE,22,"bold"), text_color=C["texto"]).pack(anchor="w")

        bf = ctk.CTkFrame(p, fg_color=C["fundo"], corner_radius=0)
        bf.pack(fill="x", padx=28, pady=(0,12))
        ctk.CTkButton(bf, text="🔄  Gerar resumo",
                      fg_color=C["laranja"], hover_color=C["laranja_hov"],
                      text_color=C["branco"], font=(FONTE,12,"bold"),
                      corner_radius=8, height=40, width=160,
                      command=self._gerar).pack(side="left")
        ctk.CTkButton(bf, text="📋  Copiar",
                      fg_color=C["branco"], text_color=C["texto"],
                      hover_color=C["fundo"], border_width=1,
                      border_color=C["borda_med"], font=(FONTE,12),
                      corner_radius=8, height=40, width=120,
                      command=self._copiar).pack(side="left", padx=(8,0))
        ctk.CTkLabel(bf, text="Copie o texto e cole direto no WhatsApp.",
                     font=(FONTE,11), text_color=C["texto_sec"]).pack(
                         side="left", padx=16)

        cr = card(p)
        cr.pack(fill="both", expand=True, padx=28, pady=(0,20))
        lbl_hint(cr, "Preview do Texto").pack(anchor="e", padx=16, pady=(12,4))
        self.txt = ctk.CTkTextbox(cr, font=(FONTE,12),
                                   fg_color=C["branco"], text_color=C["texto"],
                                   border_width=0, corner_radius=0,
                                   wrap="word", state="disabled")
        self.txt.pack(fill="both", expand=True, padx=8, pady=(0,8))

        self.st_lbl = ctk.CTkLabel(p, text="", font=(FONTE,10),
                                    text_color=C["verde"])
        self.st_lbl.pack(pady=(0,8))

    def _gerar(self):
        hoje = datetime.today().strftime("%d de %B de %Y")
        ls = ["*PAINEL SHOPEE - RESUMO DIÁRIO* 📊",
              f"_Data: {hoje}_","","━━━━━━━━━━━━━━━━━━━━━━"]
        if self.dados_pedidos:
            d = self.dados_pedidos
            ls += ["","📦 *VENDAS*",
                   f"- Pedidos: *{d['total_pedidos']}*",
                   f"- Receita Bruta: *{brl(d['receita_bruta'])}*",
                   f"- Taxas Shopee: *{brl(d['comissao']+d['taxa_servico'])}*",
                   f"- Líquido Estimado: *{brl(d['total_liquido'])}*",
                   f"- Ticket Médio: *{brl(d['ticket_medio'])}*"]
            if not d["top_produtos"].empty:
                ls += ["","📈 *TOP PRODUTOS*"]
                for i,(n,q) in enumerate(d["top_produtos"].items(),1):
                    nc = str(n)[:40]+"…" if len(str(n))>40 else str(n)
                    ls.append(f"{i}. {nc} ({q} vds)")
        if self.dados_financeiro:
            tl = self.dados_financeiro.get("total_liberado",0)
            if tl: ls.append(f"\n💰 *Repasse líquido Shopee: {brl(tl)}*")
        if self.dados_devolucoes:
            d = self.dados_devolucoes
            ls += ["","━━━━━━━━━━━━━━━━━━━━━━","","↩ *DEVOLUÇÕES*",
                   f"- Solicitações: *{d['total']}*",
                   f"- Valor em Trânsito: *{brl(d['total_reembolsado'])}*"]
            if d["motivos"]:
                top = [f"{m} ({q})" for m,q in d["motivos"].most_common(3)]
                ls.append(f"- Motivos: {', '.join(top)}")
        ls += ["","━━━━━━━━━━━━━━━━━━━━━━"]
        if len(ls) <= 6:
            ls = ["Importe pelo menos um arquivo para gerar o resumo."]
        texto = "\n".join(ls)
        self.txt.configure(state="normal")
        self.txt.delete("0.0","end")
        self.txt.insert("0.0", texto)
        self.txt.configure(state="disabled")

    def _copiar(self):
        texto = self.txt.get("0.0","end").strip()
        if texto and "Importe" not in texto:
            self.clipboard_clear(); self.clipboard_append(texto)
            if hasattr(self,"st_lbl"):
                self.st_lbl.configure(text="✅ Texto copiado! Cole no WhatsApp.")
        else:
            messagebox.showinfo("Aviso","Gere o resumo primeiro.")

    # ══════════════════════════════════════════════════════════════════════════
    # CARREGAMENTO
    # ══════════════════════════════════════════════════════════════════════════
    def _escolher(self, var):
        p = filedialog.askopenfilename(
            filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
        if p: var.set(p)
        return p

    def _load_ped(self):
        path = self._escolher(self.path_ped)
        if not path: return
        try:
            self.dados_pedidos = ler_pedidos(path); self._ir(self._aba)
        except Exception as e:
            messagebox.showerror("Erro ao ler pedidos", str(e))

    def _load_dev(self):
        path = self._escolher(self.path_dev)
        if not path: return
        try:
            self.dados_devolucoes = ler_devolucoes(path); self._ir(self._aba)
        except Exception as e:
            messagebox.showerror("Erro ao ler devoluções", str(e))

    def _load_fin(self):
        path = self._escolher(self.path_fin)
        if not path: return
        try:
            self.dados_financeiro = ler_financeiro(path); self._ir(self._aba)
        except Exception as e:
            messagebox.showerror("Erro ao ler financeiro", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
