# -*- coding: utf-8 -*-
"""Motor de busca de aluguel: coleta anúncios, aplica critérios, gera Excel
só com anúncios novos. Usado pela janela (app.py)."""

import json
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

# Pasta de dados do usuário: ~/Documents/Busca Aluguel
PASTA = Path.home() / "Documents" / "Busca Aluguel"
HISTORICO = PASTA / "historico_anuncios.json"
CONFIG = PASTA / "config.json"
PERFIL = PASTA / ".perfil_navegador"

PADRAO = {
    "cidades": "Repentigny, Lachenaie, Charlemagne, Terrebonne, Mascouche",
    "preco_max": 2700,
    "min_pecas": 4,          # 4 1/2
    "area_min": 1000,        # pi²
    "idade_max": 10,         # anos
    "exigir_garagem": True,
    "sites": {"kijiji": True, "centris": True, "appartogo": True, "marketplace": True},
    "max_paginas": 3,
}

# Setores que pertencem a outras cidades (aceitos junto com a cidade-mãe)
SETORES = {"repentigny": ["le gardeur"], "terrebonne": ["la plaine", "lachenaie"]}
# Cidade-mãe para sites que não separam o setor
CIDADE_MAE = {"lachenaie": "terrebonne", "le gardeur": "repentigny", "la plaine": "terrebonne"}
# Regiões do Appartogo (URL /canada/quebec/<regiao>/<cidade>)
REGIAO_APPARTOGO = {
    "repentigny": "lanaudiere", "terrebonne": "lanaudiere", "mascouche": "lanaudiere",
    "charlemagne": "lanaudiere", "l-assomption": "lanaudiere", "joliette": "lanaudiere",
    "saint-lin-laurentides": "lanaudiere", "laval": "laval", "montreal": "montreal",
    "longueuil": "monteregie", "brossard": "monteregie", "boucherville": "monteregie",
    "la-prairie": "monteregie", "blainville": "laurentides", "saint-jerome": "laurentides",
    "mirabel": "laurentides", "boisbriand": "laurentides", "sainte-therese": "laurentides",
}

PADROES = {
    "kijiji": re.compile(r"https?://www\.kijiji\.ca/v-[^?#\s]+/(\d{6,})"),
    "centris": re.compile(r"https?://www\.centris\.ca/(?:fr|en)/[^/?#\s]+~[^/?#\s]+/(\d{7,})"),
    "appartogo": re.compile(r"https?://appartogo\.com/(?:fr|en)/[a-z\-]+/canada/quebec/"
                            r"[a-z\-]+/[a-z\-]+/(\d{6,})"),
    "marketplace": re.compile(r"https?://www\.facebook\.com/marketplace/item/(\d+)"),
}
BASES = {"kijiji": "https://www.kijiji.ca", "centris": "https://www.centris.ca",
         "appartogo": "https://appartogo.com", "marketplace": "https://www.facebook.com"}
NOMES = {"kijiji": "Kijiji", "centris": "Centris", "appartogo": "Appartogo",
         "marketplace": "Marketplace"}


# ───────────────────────────── utilidades ─────────────────────────────
def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def norm(t):
    return sem_acento(t).lower().replace(" ", " ").replace(" ", " ")


def slug(c):
    return re.sub(r"[^a-z0-9]+", "-", norm(c)).strip("-")


def lista_cidades(cfg):
    return [c.strip() for c in cfg["cidades"].split(",") if c.strip()]


def carregar_config():
    cfg = json.loads(json.dumps(PADRAO))
    if CONFIG.exists():
        try:
            salvo = json.loads(CONFIG.read_text(encoding="utf-8"))
            cfg.update({k: v for k, v in salvo.items() if k != "sites"})
            cfg["sites"].update(salvo.get("sites", {}))
        except Exception:
            pass
    return cfg


def salvar_config(cfg):
    PASTA.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")


def carregar_historico():
    if HISTORICO.exists():
        try:
            return json.loads(HISTORICO.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def salvar_historico(h):
    PASTA.mkdir(parents=True, exist_ok=True)
    HISTORICO.write_text(json.dumps(h, ensure_ascii=False, indent=1), encoding="utf-8")


# ───────────────────────────── URLs de busca ─────────────────────────────
def urls(site, cfg):
    cidades = lista_cidades(cfg)
    pm, pag = int(cfg["preco_max"]), int(cfg["max_paginas"])
    vistos = set()

    def uma(u):
        if u not in vistos:
            vistos.add(u)
            return True
        return False

    for c in cidades:
        s = slug(c)
        mae = CIDADE_MAE.get(norm(c), None)
        if site == "kijiji":
            for p in range(1, pag + 1):
                pg = "" if p == 1 else f"page-{p}/"
                u = (f"https://www.kijiji.ca/b-apartments-condos/grand-montreal/"
                     f"{s}/{pg}k0c37l80002?price=__{pm}")
                if uma(u):
                    yield u
        elif site == "centris":
            cs = f"{slug(mae)}-{s}" if mae else s
            u = f"https://www.centris.ca/fr/propriete~a-louer~{cs}"
            if uma(u):
                yield u
        elif site == "appartogo":
            cs = slug(mae) if mae else s
            reg = REGIAO_APPARTOGO.get(cs)
            if not reg:
                continue
            for p in range(1, pag + 1):
                u = f"https://appartogo.com/en/rent/canada/quebec/{reg}/{cs}?page={p}"
                if uma(u):
                    yield u
        elif site == "marketplace":
            q = quote_plus(f"{c} {cfg['min_pecas']} 1/2")
            u = (f"https://www.facebook.com/marketplace/montreal/search?"
                 f"query={q}&maxPrice={pm}")
            if uma(u):
                yield u
    if site == "marketplace":
        quartos = max(1, int(cfg["min_pecas"]) - 2)
        yield (f"https://www.facebook.com/marketplace/montreal/propertyrentals?"
               f"maxPrice={pm}&minBedrooms={quartos}")


# ───────────────────────────── critérios ─────────────────────────────
OK, FALHA, DESC = "ok", "falha", "?"


def _num(s):
    s = re.sub(r"[\s,]", "", s).split(".")[0]
    return int(s) if s.isdigit() else None


def checar_cidade(t, cfg):
    alvo = []
    for c in lista_cidades(cfg):
        alvo.append(norm(c))
        alvo += SETORES.get(norm(c), [])
    for c in alvo:
        if re.search(r"\b" + re.escape(c) + r"\b", t):
            return OK, c.title()
    return FALHA, "fora da região"


def checar_preco(t, cfg):
    pads = [r"(\d{1,2}[ ,]?\d{3})(?:[.,]00)?\s*\$", r"\$\s*(\d{1,2}[ ,]?\d{3})",
            r"(?<!\d)(\d{3})(?:[.,]00)?\s*\$", r"\$\s*(\d{3})(?!\d)"]
    achados = []
    for p in pads:
        for m in re.finditer(p, t):
            v = _num(m.group(1))
            if v and 500 <= v <= 9000:
                achados.append((m.start(), v))
    if not achados:
        return DESC, None
    v = min(achados)[1]
    return (OK if v <= int(cfg["preco_max"]) else FALHA), v


def checar_tamanho(t, cfg):
    minp = int(cfg["min_pecas"])
    m = re.search(r"(?<![\d/])([1-9])\s*(?:1/2|½|\.5|,5)(?!\d)", t)
    if m:
        p = int(m.group(1))
        return (OK if p >= minp else FALHA), f"{p} 1/2"
    for pad in [r"(?:chambres?(?: a coucher)?|bedrooms?|cac|c\.a\.c)\s*:?\s*\n?\s*(\d)",
                r"(\d)\s*(?:chambres?|cac|c\.a\.c|cc|ch|bedrooms?|beds?|br|bdr|bdrm)\b"]:
        m = re.search(pad, t)
        if m:
            q = int(m.group(1))
            return (OK if q >= minp - 2 else FALHA), f"{q} quarto(s)"
    if re.search(r"\b(studio|bachelor)\b", t):
        return FALHA, "studio"
    return DESC, None


def checar_area(t, cfg):
    vals = []
    unid = (r"(?:pi2|pi²|pi\.?\s?ca\.?|pc\b|p\.c\.|sq\.?\s?ft|sqft|sf\b|square feet|"
            r"pieds carres|pieds²|pi/carre)")
    for m in re.finditer(r"(\d[\d ,\.]{2,6})\s*" + unid, t):
        vals.append(_num(m.group(1)))
    for m in re.finditer(r"(?:taille|size|superficie|aire habitable|living area)"
                         r"[^\d\n]{0,25}\n?\s*(\d[\d ,]{2,6})", t):
        vals.append(_num(m.group(1)))
    for m in re.finditer(r"(\d{2,3}(?:[.,]\d+)?)\s*(?:m2|m²|metres carres)", t):
        v = _num(m.group(1).replace(",", "."))
        if v:
            vals.append(round(v * 10.764))
    vals = [v for v in vals if v and 300 <= v <= 6000]
    if not vals:
        return DESC, None
    v = max(vals)
    return (OK if v >= int(cfg["area_min"]) else FALHA), v


def checar_garagem(t, cfg):
    if any(n in t for n in ["sans garage", "pas de garage", "no garage", "aucun garage",
                            "stationnement exterieur seulement", "outdoor parking only"]):
        return (FALHA if cfg["exigir_garagem"] else OK), "sem garagem"
    for k in ["garage interieur", "stationnement interieur", "garage souterrain",
              "stationnement souterrain", "indoor parking", "indoor garage",
              "underground parking", "garage attache", "attached garage", "garage integre",
              "heated garage", "garage chauffe", "parking interieur"]:
        if k in t:
            return OK, k
    if re.search(r"\bgarages?\b", t):
        return OK, "garage"
    return (DESC if cfg["exigir_garagem"] else OK), None


def checar_idade(t, cfg):
    ano_min = datetime.now().year - int(cfg["idade_max"])
    m = re.search(r"(?:construit|construction|built|bati|annee|year)[^\d]{0,30}"
                  r"((?:19|20)\d{2})", t)
    if m:
        a = int(m.group(1))
        return (OK if a >= ano_min else FALHA), a
    if any(n in t for n in ["neuf", "nouvelle construction", "new construction",
                            "brand new", "jamais habite", "never lived", "new build"]):
        return OK, "novo"
    return DESC, None


CRITERIOS = [("Cidade", checar_cidade), ("Preço", checar_preco),
             ("Tamanho", checar_tamanho), ("Área (pi²)", checar_area),
             ("Garagem", checar_garagem), ("Ano", checar_idade)]


def avaliar(texto, cfg):
    t = norm(texto)
    res, motivos, desc = {}, [], []
    for nome, f in CRITERIOS:
        st, val = f(t, cfg)
        res[nome] = val if val is not None else "?"
        if st == FALHA:
            motivos.append(f"{nome}: {val}")
        elif st == DESC:
            desc.append(nome)
    status = "Excluídos" if motivos else ("Verificar" if desc else "Atende")
    res["Falta confirmar"] = ", ".join(desc)
    res["Motivo exclusão"] = "; ".join(motivos)
    return status, res


# ───────────────────────────── Excel ─────────────────────────────
def gerar_excel(linhas, arquivo):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = ["Site", "Título", "Link", "Preço", "Cidade", "Tamanho", "Área (pi²)",
            "Garagem", "Ano", "Falta confirmar", "Motivo exclusão", "Encontrado em"]
    larg = [12, 50, 14, 9, 13, 11, 10, 18, 8, 26, 30, 17]
    cores = {"Atende": "C6EFCE", "Verificar": "FFEB9C", "Excluídos": "F2F2F2"}
    wb = Workbook()
    wb.remove(wb.active)
    for aba in cores:
        ws = wb.create_sheet(aba)
        ws.append(cols)
        for i, c in enumerate(cols, 1):
            cel = ws.cell(row=1, column=i)
            cel.font = Font(bold=True)
            cel.fill = PatternFill("solid", fgColor=cores[aba])
            ws.column_dimensions[get_column_letter(i)].width = larg[i - 1]
        ws.freeze_panes = "A2"
        for l in [x for x in linhas if x["status"] == aba]:
            ws.append([l.get(c, "") if c != "Link" else "Abrir anúncio" for c in cols])
            cel = ws.cell(row=ws.max_row, column=3)
            cel.hyperlink = l["Link"]
            cel.font = Font(color="0563C1", underline="single")
        ws.auto_filter.ref = ws.dimensions
    wb.save(arquivo)


# ───────────────────────────── navegação ─────────────────────────────
class Parar(Exception):
    pass


def _canal():
    if sys.platform == "darwin":
        return ["chrome", "msedge"]
    if sys.platform.startswith("win"):
        return ["msedge", "chrome"]
    return ["chrome", None]


def executar(cfg, log, parar=lambda: False):
    """Roda a busca completa. log(str) mostra mensagens; parar() -> True cancela.
    Retorna (caminho_excel | None, contagens)."""
    from playwright.sync_api import sync_playwright

    PASTA.mkdir(parents=True, exist_ok=True)
    hist = carregar_historico()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    log(f"Histórico: {len(hist)} anúncios já vistos.")
    sites = [s for s in ["kijiji", "centris", "appartogo", "marketplace"] if cfg["sites"].get(s)]
    linhas = []

    def checa():
        if parar():
            raise Parar()

    def abrir(page, url):
        checa()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            time.sleep(1.5)
            return True
        except Exception as e:
            log(f"   ! não abriu: {e.__class__.__name__}")
            return False

    def cookies(page):
        for txt in ["Tout accepter", "Accepter", "Accept all", "Accept",
                    "Autoriser tous les cookies", "Allow all cookies"]:
            try:
                b = page.get_by_role("button", name=txt, exact=False)
                if b.count():
                    b.first.click(timeout=1500)
                    return
            except Exception:
                pass

    def coletar(page, site):
        achados = {}
        for a in page.query_selector_all("a[href]"):
            try:
                href = a.get_attribute("href") or ""
                if href.startswith("/"):
                    href = BASES[site] + href
                m = PADROES[site].search(href)
                if m:
                    k = f"{site}:{m.group(1)}"
                    if k not in achados:
                        tit = (a.inner_text() or "").strip().replace("\n", " | ")[:200]
                        achados[k] = {"link": m.group(0), "titulo": tit}
            except Exception:
                pass
        return achados

    def texto(page):
        for txt in ["Voir plus", "See more", "Afficher plus", "Lire la suite"]:
            try:
                b = page.get_by_text(txt, exact=True)
                if b.count():
                    b.first.click(timeout=1000)
            except Exception:
                pass
        for sel in ["main", "[role='main']", "body"]:
            try:
                el = page.query_selector(sel)
                if el:
                    t = el.inner_text()
                    if len(t) > 200:
                        return t[:15000]
            except Exception:
                pass
        return ""

    def login_facebook(page):
        abrir(page, "https://www.facebook.com/marketplace/")
        if "login" in page.url or page.query_selector("input[name='email']"):
            log(">>> Faça login no Facebook na janela do navegador (até 5 min).")
            for _ in range(150):
                checa()
                time.sleep(2)
                if "login" not in page.url and not page.query_selector("input[name='email']"):
                    break
            time.sleep(3)

    try:
        with sync_playwright() as p:
            ctx, erro = None, None
            for canal in _canal():
                try:
                    ctx = p.chromium.launch_persistent_context(
                        str(PERFIL), channel=canal, headless=False, locale="fr-CA",
                        viewport={"width": 1300, "height": 900})
                    break
                except Exception as e:
                    erro = e
            if ctx is None:
                raise RuntimeError("Não encontrei o Google Chrome. Instale o Chrome "
                                   f"(google.com/chrome) e tente de novo.\n{erro}")
            page = ctx.pages[0] if ctx.pages else ctx.new_page()

            for site in sites:
                try:
                    log(f"\n=== {NOMES[site]} ===")
                    if site == "marketplace":
                        login_facebook(page)
                    achados = {}
                    for u in urls(site, cfg):
                        if not abrir(page, u):
                            continue
                        cookies(page)
                        for _ in range(8 if site == "marketplace" else 3):
                            page.mouse.wheel(0, 4000)
                            time.sleep(1.2)
                        novos = coletar(page, site)
                        log(f" página: {len(novos)} anúncios")
                        for k, v in novos.items():
                            achados.setdefault(k, v)
                        if site == "centris":
                            for _ in range(int(cfg["max_paginas"]) - 1):
                                try:
                                    nxt = page.query_selector("li.next:not(.inactive) a")
                                    if not nxt:
                                        break
                                    nxt.click()
                                    time.sleep(2.5)
                                    for k, v in coletar(page, site).items():
                                        achados.setdefault(k, v)
                                except Exception:
                                    break
                    novos = {k: v for k, v in achados.items() if k not in hist}
                    log(f" {len(achados)} encontrados, {len(novos)} novos.")
                    for i, (k, info) in enumerate(novos.items(), 1):
                        log(f"   lendo {i}/{len(novos)}")
                        txt = info["titulo"]
                        if abrir(page, info["link"]):
                            txt = info["titulo"] + "\n" + texto(page)
                        status, res = avaliar(txt, cfg)
                        linha = {"status": status, "Site": NOMES[site],
                                 "Título": info["titulo"] or txt.strip().split("\n")[0][:150],
                                 "Link": info["link"], "Encontrado em": agora, **res}
                        linhas.append(linha)
                        hist[k] = {"link": info["link"], "visto_em": agora,
                                   "status": status, "dados": linha}
                        if i % 10 == 0:
                            salvar_historico(hist)
                    salvar_historico(hist)
                except Parar:
                    raise
                except Exception as e:
                    if "closed" in str(e).lower():
                        log("\nA janela do navegador foi fechada.")
                        raise Parar()
                    log(f"\n! Erro no {NOMES[site]} ({e.__class__.__name__}). Indo para o próximo site.")
            try:
                ctx.close()
            except Exception:
                pass
    except Parar:
        log("\nBusca interrompida. Gerando a planilha com o que já foi lido.")
    except Exception as e:
        if not linhas:
            raise
        log(f"\n! Erro: {e}\nGerando a planilha com o que já foi lido.")
    finally:
        salvar_historico(hist)

    cont = {s: sum(1 for l in linhas if l["status"] == s)
            for s in ["Atende", "Verificar", "Excluídos"]}
    if not linhas:
        log("\nNenhum anúncio novo desde a última busca.")
        return None, cont
    arq = PASTA / f"aluguel_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    gerar_excel(linhas, arq)
    log(f"\nPronto! {arq.name}")
    log(f"Atende: {cont['Atende']} | Verificar: {cont['Verificar']} | "
        f"Excluídos: {cont['Excluídos']}")
    return arq, cont
