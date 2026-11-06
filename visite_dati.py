# -*- coding: utf-8 -*-
"""
Prepara i dati del tracker visite PDV.
======================================

Scrive in "visite dati\" un file JSON per ogni CBE, con i suoi negozi:
target del mese (in pezzi lordi, come le piste del variabile: fisso x1,10,
mobile x1,17, energia cosi' com'e'), negozi del patto e produzione giorno per
giorno presa dai report realtime. Scrive anche "_stato.json" con il giorno a
cui arrivano i dati.

Non pubblica niente: quei file li carica Claude nel database della pagina
(https://claude.ai/artifact/... vedi "link tracker visite.txt"), di solito
con la routine delle 8:15.

Ogni mese vanno aggiornati i tre percorsi qui sotto: file dei target, file
dei portafogli, patti del mese.

Uso:  python visite_dati.py
"""

import json
import math
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import openpyxl

import cruscotto as C

C.CARTELLA_REPORT = Path(__file__).resolve().parent / "report"

BASE = Path(__file__).resolve().parent
USCITA = BASE / "visite dati"

# ---- da aggiornare ogni mese
# versione cloud: tutto dentro il repository
FILE_TARGET = BASE / "dati" / "target.xlsx"
FOGLIO_TARGET = "NOVEMBRE"   # i fogli si chiamano "<MESE> store" e "<MESE> corner"
FILE_PORTAFOGLI = BASE / "dati" / "portafogli.xlsx"
FILE_PATTI = BASE / "dati" / "patti.json"
# assegnazioni del Trend al 31.10 che il file dei portafogli di ottobre non ha
CAMBI_PORTAFOGLI = {"0298": "ROSALBA SOMMA", "0243": "VANESSA MASO"}
MESE = (2026, 11)

MARKUP = {"f": 1.10, "m": 1.17, "e": 1.0}
SIGLA = {"Fisso": "f", "Mobile": "m", "Energia": "e"}
GIORNI_STORICO = 45


def slug(nome):
    return nome.lower().replace("'", "").replace(" ", "-")


def nome_proprio(s):
    return " ".join(w.capitalize() if not w.startswith("D'") else "D'" + w[2:].capitalize()
                    for w in s.split())


def leggi_target():
    """Target del file per negozio: store e corner (sui corner TGT ACCESSI BU va sul fisso)."""
    wb = openpyxl.load_workbook(FILE_TARGET, data_only=True, read_only=True)
    tgt = {}
    righe = list(wb[FOGLIO_TARGET + " store"].iter_rows(values_only=True))
    # la riga d'intestazione non e' sempre la stessa (ottobre: 2, novembre: 4)
    i = next(i for i, r in enumerate(righe) if r and r[0] == "CDC")
    h = [str(x).strip() if x else "" for x in righe[i]]
    for r in righe[i + 1:]:
        if r[h.index("CDC")]:
            tgt[str(r[h.index("CDC")]).zfill(4)] = {
                "f": r[h.index("PMI FISSO")] or 0, "m": r[h.index("PMI MOBILE")] or 0,
                "e": r[h.index("TIM ENERGIA")] or 0}
    righe = list(wb[FOGLIO_TARGET + " corner"].iter_rows(values_only=True))
    i = next(i for i, r in enumerate(righe) if r and "Cdc" in r)
    h = [str(x).strip() if x else "" for x in righe[i]]
    for r in righe[i + 1:]:
        if r[h.index("Cdc")]:
            tgt[str(r[h.index("Cdc")]).zfill(4)] = {
                "f": r[h.index("TGT ACCESSI BU")] or 0, "m": 0, "e": r[h.index("TGT ENERGIA")] or 0}
    wb.close()
    return tgt


def leggi_portafogli():
    wb = openpyxl.load_workbook(FILE_PORTAFOGLI, data_only=True, read_only=True)
    righe = list(wb["Portafogli ottobre"].iter_rows(values_only=True))
    h = list(righe[0])
    port = {}
    for r in righe[1:]:
        if r[0]:
            port[str(r[0]).zfill(4)] = {"cbe": r[h.index("CBE ottobre")], "loc": r[h.index("Località")],
                                        "tipo": r[h.index("Tipo")], "area": r[h.index("Area")]}
    wb.close()
    for cdc, cbe in CAMBI_PORTAFOGLI.items():
        if cdc in port:
            port[cdc]["cbe"] = cbe
    return port


def leggi_produzione():
    """Pezzi per negozio, giorno e pista: per ogni giorno l'ultimo invio che lo contiene."""
    snapshot = {}
    for percorso in sorted(C.CARTELLA_REPORT.glob("*.xlsx")):
        if percorso.name.startswith("~"):
            continue
        try:
            istante = C.ora_estrazione(percorso)
            snapshot.setdefault(istante, C.leggi_righe(percorso))
        except Exception as errore:
            print("  ! salto %s -> %s" % (percorso.name, errore))
    ultima = {}
    for istante in sorted(snapshot):
        for r in snapshot[istante]:
            ultima[(r["mese"], r["giorno"])] = istante
    oggi = date.today()
    prod = defaultdict(dict)
    giorni = []
    for (mese, giorno), istante in ultima.items():
        nm = C.chiave_mese(mese)
        anno = istante.year if nm <= istante.month else istante.year - 1
        try:
            g = date(anno, nm, int(giorno))
        except ValueError:
            continue
        if (oggi - g).days > GIORNI_STORICO:
            continue
        giorni.append((g, istante))
        chiave = g.isoformat()
        for r in snapshot[istante]:
            if r["mese"] != mese or r["giorno"] != giorno:
                continue
            fam = C.FAM_DI_TIPO.get(r["tipo"])
            if not fam:
                continue
            v = prod[r["pdv"]].setdefault(chiave, [0, 0, 0])
            v["fme".index(SIGLA[fam])] += r["pezzi"]
    return prod, giorni


def main():
    tgt = leggi_target()
    port = leggi_portafogli()
    prod, giorni = leggi_produzione()
    if not giorni:
        print("Nessun report realtime leggibile.")
        return 1
    patti = json.loads(FILE_PATTI.read_text(encoding="utf-8"))["cbe"] if FILE_PATTI.exists() else {}

    # giorno chiuso = ha la fotografia serale (dalle 22); l'ultimo giorno puo' essere a meta'
    ultimo_giorno, ultimo_istante = max(giorni)
    chiuso = ultimo_istante.hour >= 22

    per_cbe = defaultdict(list)
    tgt_cbe = defaultdict(lambda: {"f": 0, "m": 0, "e": 0})   # target del file, sommati per CBE
    for cdc, p in port.items():
        cbe = p["cbe"]
        if not cbe or str(cbe).startswith("Senza"):
            continue
        t = tgt.get(cdc)
        negozio = {
            "cdc": cdc, "loc": " ".join(w.capitalize() for w in str(p["loc"]).lower().split()),
            "tipo": p["tipo"], "area": p["area"],
            "tgt": None if t is None else {k: math.ceil(t[k] * MARKUP[k]) if t[k] else 0 for k in "fme"},
            "prod": {g: v for g, v in sorted(prod.get(cdc, {}).items())},
        }
        patto = patti.get(cbe, {})
        for f in patto.get("negozi", []):
            if f["cdc"] == cdc:
                negozio["patto"] = {"pista": f["pista"], "obiettivo": f["obiettivo"]}
        per_cbe[cbe].append(negozio)
        if t is not None:
            for k in "fme":
                tgt_cbe[cbe][k] += t[k] or 0

    USCITA.mkdir(exist_ok=True)
    for vecchio in USCITA.glob("*.json"):
        vecchio.unlink()
    for cbe, negozi in sorted(per_cbe.items()):
        patto = patti.get(cbe, {})
        doc = {"cbe": slug(cbe), "nome": nome_proprio(cbe), "mese": "%d-%02d" % MESE,
               "negozi": sorted(negozi, key=lambda n: n["loc"]),
               # target delle piste del CBE: maggiorazione sul totale, mai negozio per negozio
               "target": {k: round(tgt_cbe[cbe][k] * MARKUP[k]) for k in "fme"},
               "patto": {"tipo": patto.get("tipo")} if patto.get("tipo") != "mantieni" else
                        {"tipo": "mantieni", "fisso_mobile": patto["fisso_mobile"], "energia": patto["energia"]}}
        (USCITA / (slug(cbe) + ".json")).write_text(json.dumps(doc, ensure_ascii=False), encoding="utf-8")
    stato = {"ultimo_giorno": ultimo_giorno.isoformat(), "giorno_chiuso": chiuso,
             "ultimo_invio": ultimo_istante.strftime("%Y-%m-%d %H:%M"), "mese": "%d-%02d" % MESE,
             "giorni_mese": (date(MESE[0] + MESE[1] // 12, MESE[1] % 12 + 1, 1) - timedelta(days=1)).day}
    (USCITA / "_stato.json").write_text(json.dumps(stato, ensure_ascii=False), encoding="utf-8")
    print("Scritti %d CBE in %s" % (len(per_cbe), USCITA))
    print("Dati fino al %s (invio %s%s)" % (ultimo_giorno.strftime("%d/%m"), ultimo_istante.strftime("%H:%M"),
                                           ", giornata chiusa" if chiuso else ", giornata in corso"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
