# -*- coding: utf-8 -*-
"""Testes dos filtros (roda antes de gerar o app)."""
import motor

cfg = dict(motor.PADRAO)
casos = [
    ("Atende", "Condo à louer\n150, Rue de la Pruche, Terrebonne (Lachenaie)\n2 620 $ /mois\n"
               "5 pièces\n2 chambres\nSuperficie nette\n1 025 pc\nAnnée de construction\n2023\n"
               "Stationnement total\nGarage (1)"),
    ("Atende", "Condo neuf Repentigny\n$2,300.00\nChambres: 2\nTaille (pi²)\n1100\n"
               "Stationnement intérieur inclus"),
    ("Excluídos", "3 1/2 Mascouche 1500$ garage"),
    ("Excluídos", "5½ Charlemagne 2 950 $ garage 1300 pc construit en 2020"),
    ("Excluídos", "4 1/2 Laval 2000$ garage intérieur 1100 pi2 2018"),
    ("Verificar", "Beau 4½ à Repentigny, 2200$, stationnement extérieur"),
    ("Excluídos", "4 1/2 Terrebonne 2100$ garage 1200 sq ft built in 1995"),
]
erros = 0
for esperado, txt in casos:
    st, res = motor.avaliar(txt, cfg)
    ok = st == esperado
    erros += not ok
    print("OK " if ok else "ERRO", esperado, "->", st, res)

cfg2 = dict(cfg, cidades="Laval, Brossard", preco_max=3000, min_pecas=3, exigir_garagem=False)
st, _ = motor.avaliar("3 1/2 Laval 2500$ 1100 pi2 construit en 2020", cfg2)
print("Critérios editáveis:", st)
erros += st != "Atende"

u = list(motor.urls("centris", cfg))
assert "https://www.centris.ca/fr/propriete~a-louer~terrebonne-lachenaie" in u, u
assert any("lanaudiere/mascouche" in x for x in motor.urls("appartogo", cfg))
assert motor.PADROES["appartogo"].search(
    "https://appartogo.com/fr/appartement-louer/canada/quebec/lanaudiere/terrebonne/2429251")
assert motor.PADROES["centris"].search(
    "https://www.centris.ca/fr/condo-appartement~a-louer~terrebonne-lachenaie/13990712")

for txt, v in [("CA$1.487 | Appartement", 1487), ("2 450 $ / mois", 2450), ("$2,300.00", 2300),
               ("1500$", 1500), ("CA$1 | Lachine 5 1/2", None), ("$1,895 - Available", 1895)]:
    assert motor.ler_preco(motor.norm(txt)) == v, (txt, motor.ler_preco(motor.norm(txt)))
pf = motor.prefiltro_marketplace
assert pf("CA$2.100 | Appartement de 2 chambres et 1 salle de bain | Terrebonne, QC", cfg)
assert not pf("CA$1.500 | 2011 Mazda 3 sport | Repentigny, QC", cfg)
assert not pf("CA$1.487 | Appartement de 3 chambres | Longueuil, QC", cfg)
assert not pf("CA$2.950 | Condo 4 1/2 | Mascouche, QC", cfg)
st, res = motor.avaliar("busca: Repentigny\nCA$1.400 | Appartement 2 chambres | Montréal, QC\ngarage",
                        cfg, "CA$1.400 | Appartement 2 chambres | Montréal, QC")
assert st == "Excluídos", res
print("Preço e pré-filtro do Marketplace OK")

import tempfile, os
f = os.path.join(tempfile.gettempdir(), "teste.xlsx")
motor.gerar_excel([{"status": "Atende", "Site": "Centris", "Título": "x",
                    "Link": "https://www.centris.ca/fr/a~b/12345678", **res}], f)
print("Excel OK")
if erros:
    raise SystemExit(f"{erros} teste(s) falharam")
print("Todos os testes passaram.")
