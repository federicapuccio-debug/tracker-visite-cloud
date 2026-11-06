# -*- coding: utf-8 -*-
"""
Cruscotto realtime - commessa TIM Retail
========================================

Legge tutti i report realtime scaricati in "report ricevuti\\realtime", li ordina
per ora di estrazione, e risponde a tre domande:

  quanti pezzi mancano       TGT del mese meno il prodotto finora
  che passo bisogna tenere   pezzi al giorno sui giorni che restano
  su che negozi intervenire  i PDV la cui proiezione a fine mese sta sotto il TGT

Produce:
  cruscotto.html         due viste: NEGOZI (dove intervenire) e CBE
  pagine per CBE\\*.html  una pagina per ciascun CBE, da inoltrare
  messaggi.txt           un blocco di testo per CBE, da incollare su WhatsApp

Due cose che il report non dice e che qui vengono ricostruite:

  1. L'ORA DI ESTRAZIONE non e' in nessuna cella. Sta nei dati interni della
     cartella di lavoro (aggiornamento della pivot, in mancanza le proprieta'
     del file). Senza quella non esiste il confronto fra un invio e l'altro.

  2. IL PROGRESSIVO DEL MESE non c'e': ogni file contiene solo la giornata in
     corso. Va ricostruito prendendo, per ciascun giorno, l'ultimo invio
     disponibile di quel giorno. Per questo l'archivio non va cancellato.

E una terza cosa che il report non puo' dire da solo: i negozi FERMI A ZERO non
compaiono, perche' non hanno righe. Il perimetro sta in "obiettivi PDV.csv".

Uso: doppio clic su "AGGIORNA CRUSCOTTO.bat"
"""

import csv
import html
import os
import re
import sys
import zipfile
from calendar import monthrange
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("\n  Manca il componente 'openpyxl', che serve a leggere i file Excel.")
    print("  Apri il Prompt dei comandi e scrivi:   pip install openpyxl\n")
    sys.exit(1)

# ---------------------------------------------------------------- impostazioni

BASE = Path(__file__).resolve().parent
CARTELLA_REPORT = BASE.parent / "report ricevuti" / "realtime"
CARTELLA_PAGINE = BASE / "pagine per CBE"
FILE_PERIMETRO = BASE / "obiettivi PDV.csv"
FILE_PERSE = BASE / "giornate perse.txt"

# "TOT GAS e LUCE" non e' una famiglia: e' gia' la somma di TIM GAS + TIM LUCE.
# Sommarla insieme alle altre conta i pezzi due volte.
TIPO_DOPPIONE = "TOT GAS e LUCE"
TIPI = ["TIM GAS", "TIM LUCE", "PMI-Fisso", "PMI mobili"]
SIGLE = {"TIM GAS": "GAS", "TIM LUCE": "LUCE", "PMI-Fisso": "PMI-F", "PMI mobili": "PMI-M"}
# I target arrivano per famiglia commerciale, il report per tipo di pista.
# Energia e' l'unica che raggruppa due tipi: GAS e LUCE.
FAMIGLIE = ["Fisso", "Mobile", "Energia"]
FAM_DI_TIPO = {"PMI-Fisso": "Fisso", "PMI mobili": "Mobile",
               "TIM GAS": "Energia", "TIM LUCE": "Energia"}
SENZA_CBE = "_"
# Un negozio che sta nel perimetro ma non ha ancora un CBE assegnato NON e'
# "fuori perimetro": e' roba tua che aspetta un nome. Tenerli distinti evita
# che 833 pezzi di target spariscano sotto l'etichetta sbagliata.
DA_ASSEGNARE = "(da assegnare)"

MESI = {"01": 1, "02": 2, "03": 3, "04": 4, "05": 5, "06": 6,
        "07": 7, "08": 8, "09": 9, "10": 10, "11": 11, "12": 12}


# ------------------------------------------------------------------ lettura

def ora_estrazione(percorso):
    """L'ora in cui il report e' stato generato, letta dai dati interni."""
    try:
        with zipfile.ZipFile(percorso) as z:
            nomi = z.namelist()
            for nome in nomi:
                if "pivotCacheDefinition" in nome:
                    xml = z.read(nome).decode("utf-8", "replace")
                    m = re.search(r'refreshedDate="([0-9.]+)"', xml)
                    if m:  # gia' in ora locale di chi ha generato il file
                        istante = datetime(1899, 12, 30) + timedelta(days=float(m.group(1)))
                        return istante.replace(microsecond=0)
            if "docProps/core.xml" in nomi:
                xml = z.read("docProps/core.xml").decode("utf-8", "replace")
                m = re.search(r"<dcterms:created[^>]*>([^<]+)</dcterms:created>", xml)
                if m:  # in UTC, da riportare all'ora italiana
                    utc = datetime.strptime(m.group(1), "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    try:
                        from zoneinfo import ZoneInfo
                        locale = utc.astimezone(ZoneInfo("Europe/Rome"))
                    except Exception:
                        locale = utc + timedelta(hours=2)
                    return locale.replace(tzinfo=None, microsecond=0)
    except Exception:
        pass
    return datetime.fromtimestamp(os.path.getmtime(percorso)).replace(microsecond=0)


def leggi_righe(percorso):
    """Righe del foglio Dati, senza il doppione TOT GAS e LUCE."""
    # Il file viene letto dalla memoria, non dal disco: se il contenuto e'
    # rovinato openpyxl solleva un errore restando agganciato al file, e su
    # Windows quel file non si riesce piu' a cancellare.
    import io
    with open(percorso, "rb") as f:
        contenuto = io.BytesIO(f.read())
    wb = openpyxl.load_workbook(contenuto, data_only=True, read_only=True)
    try:
        if "Dati" not in wb.sheetnames:
            raise ValueError("il foglio 'Dati' non c'e'")
        griglia = [list(r) for r in wb["Dati"].iter_rows(values_only=True)]
    finally:
        wb.close()

    intestazione = None
    for i, riga in enumerate(griglia[:6]):
        valori = [str(v).strip() if v is not None else "" for v in riga]
        if "Cbe" in valori and "Pz Tot" in valori:
            intestazione, inizio = valori, i + 1
            break
    if intestazione is None:
        raise ValueError("non trovo la riga di intestazione nel foglio 'Dati'")

    col = {nome: i for i, nome in enumerate(intestazione) if nome}
    mancanti = [c for c in ("Mese", "Giorno", "Tipo", "Cbe", "PDV", "Localita", "Pz Tot") if c not in col]
    if mancanti:
        raise ValueError("colonne mancanti nel foglio 'Dati': " + ", ".join(mancanti))

    righe = []
    for riga in griglia[inizio:]:
        if not riga or riga[col["Cbe"]] in (None, ""):
            continue
        tipo = str(riga[col["Tipo"]]).strip()
        if tipo == TIPO_DOPPIONE:
            continue
        try:
            pezzi = int(riga[col["Pz Tot"]] or 0)
        except (TypeError, ValueError):
            continue
        righe.append({
            "mese": str(riga[col["Mese"]]).strip(),
            "giorno": str(riga[col["Giorno"]]).strip(),
            "tipo": tipo,
            "cbe": str(riga[col["Cbe"]]).strip(),
            "pdv": codice_pdv(riga[col["PDV"]]),
            "loc": str(riga[col["Localita"]]).strip(),
            "canale": str(riga[col["Canale"]]).strip() if "Canale" in col and riga[col["Canale"]] else "",
            "area": str(riga[col["Area"]]).strip() if "Area" in col and riga[col["Area"]] else "",
            "pezzi": pezzi,
        })
    return righe


def rinomina(percorso, istante):
    """Da 'report_realtime_CBE (3).xlsx' a '2026-09-01 2228 - Realtime CBE.xlsx'."""
    if not percorso.name.lower().startswith("report_realtime"):
        return percorso
    nuovo = percorso.with_name(istante.strftime("%Y-%m-%d %H%M") + " - Realtime CBE.xlsx")
    if nuovo.exists():
        return percorso
    try:
        percorso.rename(nuovo)
        return nuovo
    except OSError:
        return percorso


def leggi_stato():
    """Com'e' andato l'ultimo controllo della posta, scritto da scarica.py."""
    import json
    percorso = BASE / "stato.json"
    if not percorso.exists():
        return {}
    try:
        return json.loads(percorso.read_text(encoding="utf-8"))
    except Exception:
        return {}


def automatico_attivo():
    """Se l'operazione pianificata non c'e', non ha senso lamentarsi che il
    controllo non gira da sola: e' lei che lo lancia a mano."""
    import subprocess
    try:
        esito = subprocess.run(["schtasks", "/Query", "/TN", "Cruscotto TIM Retail"],
                               capture_output=True, timeout=15,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return esito.returncode == 0
    except Exception:
        return False


def leggi_perimetro():
    """obiettivi PDV.csv: PDV;Localita;CBE;TGT Fisso;TGT Mobile;TGT Energia.

    E' sia l'elenco dei negozi da tenere d'occhio (compresi quelli fermi a zero,
    che nel report non compaiono affatto) sia il posto dove mettere i target.
    Le caselle vuote restano vuote: nessun numero viene inventato.
    """
    perimetro = {}
    if not FILE_PERIMETRO.exists():
        return perimetro
    try:
        with open(FILE_PERIMETRO, encoding="utf-8-sig", newline="") as f:
            for riga in csv.reader(f, delimiter=";"):
                if len(riga) < 3 or not riga[0].strip() or riga[0].strip().upper() == "PDV":
                    continue
                chiuso = False

                def numero(i):
                    nonlocal chiuso
                    if len(riga) <= i or not riga[i].strip():
                        return None
                    testo = riga[i].strip()
                    # nel TGT un negozio chiuso ha la parola al posto del numero
                    if testo.lower().startswith("chius"):
                        chiuso = True
                        return None
                    try:
                        return int(float(testo.replace(".", "").replace(",", ".")))
                    except ValueError:
                        return None
                tgt = {f: numero(3 + k) for k, f in enumerate(FAMIGLIE)}
                perimetro[codice_pdv(riga[0])] = {
                    "loc": riga[1].strip(),
                    "cbe": riga[2].strip().upper(),
                    "area": riga[6].strip() if len(riga) > 6 else "",
                    "tgt": tgt,
                    "tgt_tot": sum(v for v in tgt.values() if v) or None,
                    "chiuso": chiuso,
                }
    except OSError as errore:
        print("  Attenzione: non riesco a leggere 'obiettivi PDV.csv' (%s). Lo ignoro." % errore)
    return perimetro


def leggi_giornate_perse():
    """Le giornate che non arriveranno mai piu', dichiarate perse a mano.

    Un report mancante blocca tutto: il progressivo del mese sarebbe sbagliato,
    quindi la mail ai CBE non parte. Se pero' il fornitore non puo' piu'
    rimandarlo - come il 14/09/2026, saltato per intero da E3 - il blocco
    resterebbe per sempre. Elencando qui la giornata si accetta il buco: il
    cruscotto smette di bloccare ma NON smette di dirlo, perche' quei pezzi
    restano fuori dai totali.

    File "giornate perse.txt", una data per riga: 2026-09-14  motivo libero.
    Torna {numero del mese: set dei giorni}.
    """
    perse = defaultdict(set)
    if not FILE_PERSE.exists():
        return perse
    for riga in FILE_PERSE.read_text(encoding="utf-8").splitlines():
        riga = riga.strip()
        if not riga or riga.startswith("#"):
            continue
        pezzo = riga.split()[0]
        try:
            data = datetime.strptime(pezzo, "%Y-%m-%d")
        except ValueError:
            continue
        perse[data.month].add("%02d" % data.day)
    return perse


def nota_giornate(buchi, perse):
    """La frase sotto 'Giornate archiviate'. Una giornata persa non e' un buco
    da riempire, ma non e' nemmeno 'nessun giorno mancante': va detta."""
    if buchi:
        return "mancano i giorni %s" % ", ".join(buchi)
    if perse:
        return "il %s non arrivera': giornata persa, non e' nei totali" % ", ".join(sorted(perse))
    return "nessun giorno mancante"


# --------------------------------------------------------------- aggregazione

def chiave_mese(mese):
    return MESI.get(mese[:2], 0)


def codice_pdv(valore):
    """I codici PDV sono '0214'. Excel li tratta come numeri e li riduce a 214,
    o peggio a '214.0'. Qui tornano sempre alla forma a quattro cifre."""
    testo = str(valore).strip()
    if testo.endswith(".0"):
        testo = testo[:-2]
    return testo.zfill(4) if testo.isdigit() else testo


def nome_proprio(s):
    if s == SENZA_CBE:
        return "Fuori perimetro"
    if s == DA_ASSEGNARE:
        return "Da assegnare"
    return " ".join(p[:1] + p[1:].lower() for p in s.split(" ")).replace("D'a", "D'A")


def voce():
    """I numeri di una singola famiglia, o del totale."""
    return {"tgt": None, "mese": 0, "oggi": 0, "delta": None}


def nuovo_pdv():
    p = {"loc": "", "cbe": "", "canale": "", "area": "", "tipi": Counter(), "giorni": set(),
         "nel_perimetro": False, "chiuso": False,
         "tot": voce()}
    for f in FAMIGLIE:
        p[f] = voce()
    return p


def costruisci(righe_mese, giorno_corrente, perimetro):
    """Un dizionario per PDV, che parte dal perimetro e ci versa dentro il prodotto."""
    pdv = {}
    for codice, info in perimetro.items():
        p = nuovo_pdv()
        p.update(loc=info["loc"], cbe=info["cbe"] or DA_ASSEGNARE, area=info.get("area", ""),
                 chiuso=info["chiuso"], nel_perimetro=True)
        for f in FAMIGLIE:
            p[f]["tgt"] = info["tgt"].get(f)
        p["tot"]["tgt"] = info["tgt_tot"]
        pdv[codice] = p
    for r in righe_mese:
        p = pdv.get(r["pdv"])
        if p is None:
            p = pdv[r["pdv"]] = nuovo_pdv()
            p["loc"] = r["loc"]
        if not p["cbe"]:
            p["cbe"] = r["cbe"].upper()
        p["canale"] = r["canale"] or p["canale"]
        p["area"] = r["area"] or p["area"]
        p["tipi"][r["tipo"]] += r["pezzi"]
        p["giorni"].add(r["giorno"])
        fam = FAM_DI_TIPO.get(r["tipo"])
        for chiave in ("tot", fam):
            if chiave:
                p[chiave]["mese"] += r["pezzi"]
                if r["giorno"] == giorno_corrente:
                    p[chiave]["oggi"] += r["pezzi"]
    return pdv


def curva_giornata(snapshot, escludi):
    """Quanta parte della giornata e' gia' fatta a una certa ora, ricavata dalle
    giornate chiuse in archivio. Serve a leggere un invio di meta' pomeriggio:
    '105 pezzi' non dice niente finche' non sai che a quell'ora se ne fa la meta'.

    Torna {pista: [(minuti dalla mezzanotte, quota), ...]}.
    """
    per_giorno = defaultdict(lambda: defaultdict(Counter))
    for istante, righe in snapshot.items():
        for r in righe:
            giorno = (r["mese"], r["giorno"])
            fam = FAM_DI_TIPO.get(r["tipo"])
            per_giorno[giorno][istante]["tot"] += r["pezzi"]
            if fam:
                per_giorno[giorno][istante][fam] += r["pezzi"]

    punti, usate = defaultdict(list), 0
    for giorno, istanti in per_giorno.items():
        if giorno == escludi:
            continue
        ultimo = max(istanti)
        # una giornata vale solo se ne abbiamo la fotografia finale
        if ultimo.hour < 21:
            continue
        usate += 1
        finale = istanti[ultimo]
        for istante, conto in istanti.items():
            minuti = istante.hour * 60 + istante.minute
            for pista in FAMIGLIE + ["tot"]:
                if finale[pista]:
                    punti[pista].append((minuti, conto[pista] / finale[pista]))
    return punti, usate


def stima_quota(punti, pista, minuti, finestra=45):
    """La quota attesa a quell'ora, come media delle giornate passate.
    Se per la pista i dati sono pochi ripiega sulla curva del totale."""
    for chiave in (pista, "tot"):
        vicini = [q for m, q in punti.get(chiave, []) if abs(m - minuti) <= finestra]
        if len(vicini) >= 2:
            return sum(vicini) / len(vicini)
    return None


def calcola_passo(v, trascorsi, residui, giorni_mese, quota=None):
    """Passo tenuto, passo da tenere, proiezione. Vale sia per una famiglia
    che per il totale: la voce ha sempre la stessa forma.

    Se la giornata in corso e' solo a meta', 'quota' dice a che punto dovrebbe
    essere: cosi' il passo non risulta artificialmente basso ogni mattina.
    """
    v["quota"] = quota
    v["oggi_stima"] = None
    if quota and 0.02 < quota < 0.97 and v["oggi"] > 0:
        v["oggi_stima"] = round(v["oggi"] / quota)

    # per il passo, la giornata in corso conta per quello che diventera'
    fine_oggi = v["oggi_stima"] if v["oggi_stima"] is not None else v["oggi"]
    mese_pieno = v["mese"] - v["oggi"] + fine_oggi
    v["proiettato"] = v["oggi_stima"] is not None

    v["passo"] = mese_pieno / trascorsi if trascorsi else 0
    v["proiezione"] = round(v["passo"] * giorni_mese)
    if v["tgt"]:
        v["mancano"] = max(v["tgt"] - v["mese"], 0)
        # residui = i giorni ancora da fare, oggi escluso perche' e' gia' contato
        # fra i trascorsi. Cosi' "serve/gg" e la proiezione dicono sempre la
        # stessa cosa: se serve piu' del passo tenuto, la proiezione sta sotto.
        v["serve"] = v["mancano"] / residui if residui else None
        v["scarto"] = v["proiezione"] - v["tgt"]
        v["perc"] = round(100 * v["mese"] / v["tgt"])
    else:
        v["mancano"] = v["serve"] = v["scarto"] = v["perc"] = None


def calcola_tutto(p, trascorsi, residui, giorni_mese, quote):
    for chiave in ["tot"] + FAMIGLIE:
        calcola_passo(p[chiave], trascorsi, residui, giorni_mese, quote.get(chiave))


def raggruppa_per_cbe(pdv, trascorsi, residui, giorni_mese, quote):
    cbe = defaultdict(lambda: dict({"tipi": Counter(), "pdv": {}, "tot": voce()},
                                   **{f: voce() for f in FAMIGLIE}))
    for codice, p in pdv.items():
        c = cbe[p["cbe"] or SENZA_CBE]
        c["pdv"][codice] = p
        c["tipi"] += p["tipi"]
        for chiave in ["tot"] + FAMIGLIE:
            d, s = c[chiave], p[chiave]
            d["mese"] += s["mese"]
            d["oggi"] += s["oggi"]
            if s["tgt"]:
                d["tgt"] = (d["tgt"] or 0) + s["tgt"]
            if s["delta"] is not None:
                d["delta"] = (d["delta"] or 0) + s["delta"]
    for c in cbe.values():
        for chiave in ["tot"] + FAMIGLIE:
            calcola_passo(c[chiave], trascorsi, residui, giorni_mese, quote.get(chiave))
    return cbe


def somma_voci(elenco, trascorsi, residui, giorni_mese, quote):
    """Aggrega una lista di PDV in un'unica riga per famiglia + totale."""
    fuori = {chiave: voce() for chiave in ["tot"] + FAMIGLIE}
    for p in elenco:
        for chiave in ["tot"] + FAMIGLIE:
            d, s = fuori[chiave], p[chiave]
            d["mese"] += s["mese"]
            d["oggi"] += s["oggi"]
            if s["tgt"]:
                d["tgt"] = (d["tgt"] or 0) + s["tgt"]
            if s["delta"] is not None:
                d["delta"] = (d["delta"] or 0) + s["delta"]
    for chiave in ["tot"] + FAMIGLIE:
        calcola_passo(fuori[chiave], trascorsi, residui, giorni_mese, quote.get(chiave))
    return fuori


# ------------------------------------------------------------------- pagina

CSS = """
:root{--ground:#f2f5f8;--surface:#fff;--surface-2:#f8fafc;--ink:#141e28;--ink-2:#33475b;
--muted:#63758a;--line:#dbe3ec;--line-2:#eaeff5;--accent:#0a58c8;--accent-soft:#e7effc;
--pmi:#0d6f74;--ok:#12704a;--ok-soft:#e4f2ea;--crit:#ab2637;--crit-soft:#fbeaec;
--warn:#9a5b06;--warn-soft:#fbf0dd;
--gap:#8a6a1f;--gap-soft:#f7f0dd;--gap-line:#ddc98f;
--shadow:0 1px 2px rgba(20,30,40,.06),0 8px 24px -16px rgba(20,30,40,.28)}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){
--ground:#0d141c;--surface:#161f2a;--surface-2:#1b2531;--ink:#e7eef6;--ink-2:#b9c8d7;
--muted:#8ea0b3;--line:#28333f;--line-2:#212b36;--accent:#6aa6ff;--accent-soft:#12263e;
--pmi:#48b8b2;--ok:#4cbc8a;--ok-soft:#122b21;--crit:#f0798a;--crit-soft:#2e161b;
--warn:#d9a253;--warn-soft:#2b2013;
--gap:#d5b76b;--gap-soft:#2a2415;--gap-line:#4c421f;
--shadow:0 1px 2px rgba(0,0,0,.4),0 8px 24px -16px rgba(0,0,0,.8)}}
*{box-sizing:border-box}
body{background:var(--ground);color:var(--ink);margin:0;padding:28px 18px 60px;
font:15px/1.55 "Segoe UI",system-ui,-apple-system,sans-serif;-webkit-font-smoothing:antialiased}
h1,h2,h3{text-wrap:balance}
.wrap{max-width:1200px;margin:0 auto;display:flex;flex-direction:column;gap:22px}
.eyebrow{font-size:11.5px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;
color:var(--muted);margin:0 0 6px}
h1{font-size:clamp(24px,4vw,33px);font-weight:700;letter-spacing:-.02em;margin:0 0 8px;line-height:1.14}
.masthead p{margin:0;color:var(--ink-2);max-width:70ch}
.stamp{display:inline-flex;flex-wrap:wrap;gap:4px 8px;margin-top:12px;font-size:12.5px;
font-family:Consolas,"Courier New",monospace;border-radius:6px;padding:6px 10px;
background:var(--accent-soft);color:var(--accent)}
.stamp.warn{background:var(--gap-soft);color:var(--gap);border:1px dashed var(--gap-line)}

/* riga di sintesi */
.sommario{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:1px;
background:var(--line);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.som{background:var(--surface);padding:14px 16px 15px}
.som .k{font-size:10.5px;font-weight:600;letter-spacing:.11em;text-transform:uppercase;color:var(--muted)}
.som .v{font-size:27px;font-weight:700;letter-spacing:-.02em;margin-top:5px;font-variant-numeric:tabular-nums}
.som .v small{font-size:13px;font-weight:500;color:var(--muted);letter-spacing:0}
.som .n{font-size:11.5px;color:var(--muted);margin-top:4px;line-height:1.35}
.som.vuoto{background:var(--surface-2)}
.som.vuoto .v{color:var(--gap-line)}
.som.vuoto .k,.som.vuoto .n{color:var(--gap)}
.som.buono .v{color:var(--ok)}
.som.male .v{color:var(--crit)}

/* la fascia d'allarme: calcolata quando apri la pagina, non quando viene scritta */
.allarme-fascia{display:none;border-radius:10px;padding:13px 16px;font-size:14px;line-height:1.5;
border:1px solid transparent}
.allarme-fascia.mostra{display:block}
.allarme-fascia b{display:block;font-size:15px;margin-bottom:3px}
.allarme-fascia.rosso{background:var(--crit-soft);border-color:var(--crit);color:var(--crit)}
.allarme-fascia.giallo{background:var(--warn-soft);border-color:var(--warn);color:var(--warn)}
.allarme-fascia small{display:block;margin-top:5px;opacity:.85;font-size:12.5px}

/* il progressivo del mese, una carta per pista */
.piste{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:14px}
.pista{background:var(--surface);border:1px solid var(--line);border-top:3px solid var(--muted);
border-radius:12px;padding:15px 17px 8px;box-shadow:var(--shadow);display:flex;flex-direction:column}
.pista.pmi{border-top-color:var(--pmi)}
.pista.energia{border-top-color:var(--accent)}
.pista.totale{background:var(--surface-2);border-top-color:var(--line)}
.p-nome{font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}
.pista.pmi .p-nome{color:var(--pmi)}
.pista.energia .p-nome{color:var(--accent)}
.p-mese{font-size:34px;font-weight:700;letter-spacing:-.025em;line-height:1.05;margin-top:7px;
font-variant-numeric:tabular-nums}
.p-mese small{font-size:12.5px;font-weight:500;color:var(--muted);letter-spacing:0}
.p-avanz{margin-top:9px;font-size:12px;font-weight:600;font-variant-numeric:tabular-nums}
.p-avanz span{display:block;font-weight:400;color:var(--muted);font-size:11px;margin-top:3px}
.p-righe{margin-top:11px;border-top:1px solid var(--line-2)}
.p-riga{display:flex;align-items:baseline;gap:6px;padding:6px 0;border-bottom:1px solid var(--line-2);
font-size:12.5px}
.p-riga:last-child{border-bottom:none}
.p-riga span{color:var(--muted);flex:1;min-width:0}
.p-riga b{font-weight:700;font-size:14.5px;font-variant-numeric:tabular-nums}
.p-riga em{font-style:normal;color:var(--muted);font-size:11px}
.p-riga i{font-style:normal;font-size:12px}
.p-riga.allarme b{color:var(--crit)}
.p-riga.ok b{color:var(--ok)}
.p-riga.stima{background:var(--accent-soft);margin:0 -17px;padding:6px 17px}
.p-riga.stima span,.p-riga.stima em{color:var(--accent);opacity:.85}
.p-riga.stima b{color:var(--accent)}
.p-riga.vuoto b,.p-riga.vuoto em{color:var(--gap)}
.piede{padding:0;margin:14px 0 0}

/* schede */
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--line)}
.tabs button{font:600 13px/1 "Segoe UI",sans-serif;background:none;border:none;color:var(--muted);
padding:10px 15px;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
.tabs button:hover{color:var(--ink)}
.tabs button[aria-selected=true]{color:var(--accent);border-bottom-color:var(--accent)}
.tabs button:focus-visible{outline:2px solid var(--accent);outline-offset:-2px}

.pannello{background:var(--surface);border:1px solid var(--line);border-radius:12px;
box-shadow:var(--shadow);overflow:hidden}
.pannello > h3,.sec > h3{font-size:10.5px;font-weight:600;letter-spacing:.11em;
text-transform:uppercase;color:var(--muted);margin:0;padding:14px 18px 0}
.spiega{font-size:12.5px;color:var(--muted);padding:6px 18px 0;margin:0;max-width:82ch}

/* filtri */
.filtri{display:flex;flex-wrap:wrap;gap:6px;padding:12px 18px 14px}
.filtri select,.filtri input{font:13px "Segoe UI",sans-serif;background:var(--surface);
color:var(--ink);border:1px solid var(--line);border-radius:7px;padding:6px 9px}
.filtri input{min-width:190px}
.filtri select:focus-visible,.filtri input:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.conteggio{margin-left:auto;align-self:center;font-size:12.5px;color:var(--muted)}

.tbl-wrap{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13px}

/* Le colonne esistono tutte nel documento; il selettore di famiglia decide
   quali mostrare. Nella vista d'insieme le tre famiglie stanno affiancate,
   ciascuna con le sole colonne essenziali. */
table th.c-tot,table td.c-tot,table th.c-fisso,table td.c-fisso,
table th.c-mobile,table td.c-mobile,table th.c-energia,table td.c-energia{display:none}
/* vista su una pista sola: tutte le sue colonne, con l'intestazione larga */
table.v-tot th.c-tot:not(.g-base),table.v-tot td.c-tot{display:table-cell}
table.v-fisso th.c-fisso:not(.g-base),table.v-fisso td.c-fisso{display:table-cell}
table.v-mobile th.c-mobile:not(.g-base),table.v-mobile td.c-mobile{display:table-cell}
table.v-energia th.c-energia:not(.g-base),table.v-energia td.c-energia{display:table-cell}
/* viste affiancate: solo le colonne essenziali, e l'intestazione stretta che le copre */
table.v-tutte th.c-fisso.k-base,table.v-tutte td.c-fisso.k-base,
table.v-tutte th.c-mobile.k-base,table.v-tutte td.c-mobile.k-base,
table.v-tutte th.c-energia.k-base,table.v-tutte td.c-energia.k-base,
table.v-tutte th.c-fisso.g-base,table.v-tutte th.c-mobile.g-base,
table.v-tutte th.c-energia.g-base{display:table-cell}
th.gruppo{text-align:center;font-size:10.5px;color:var(--accent);letter-spacing:.1em;
border-bottom:1px solid var(--line-2);padding-bottom:5px}
table.v-tutte td.c-fisso.k-mancano,table.v-tutte th.c-fisso.k-mancano,
table.v-tutte td.c-mobile.k-mancano,table.v-tutte th.c-mobile.k-mancano{
border-right:1px solid var(--line)}
th{font-size:10px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);
text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);white-space:nowrap;
background:var(--surface-2);position:sticky;top:0}
th.num{text-align:right}
th.sortable{cursor:pointer;user-select:none}
th.sortable:hover{color:var(--accent)}
th[aria-sort] .fr::after{content:" \\2193";color:var(--accent)}
th[aria-sort=ascending] .fr::after{content:" \\2191"}
td{padding:7px 10px;border-bottom:1px solid var(--line-2);font-variant-numeric:tabular-nums;
white-space:nowrap}
td.num{text-align:right}
tbody tr:hover{background:var(--surface-2)}
tr:last-child td{border-bottom:none}
.code{font-family:Consolas,"Courier New",monospace;font-size:11.5px;color:var(--muted)}
.loc{font-weight:600}
.sub{color:var(--muted);font-size:11.5px}
.chip{font-size:10px;font-weight:600;letter-spacing:.05em;padding:2px 7px;border-radius:4px;
background:var(--line-2);color:var(--ink-2);display:inline-block}
.chip.rosso{background:var(--crit-soft);color:var(--crit)}
.chip.giallo{background:var(--warn-soft);color:var(--warn)}
.chip.verde{background:var(--ok-soft);color:var(--ok)}
.up{color:var(--ok);font-weight:700}
.giu{color:var(--crit);font-weight:700}
.zero{color:var(--muted)}
.vuota{color:var(--gap-line)}
.barra{display:inline-block;width:54px;height:7px;border-radius:4px;background:var(--line-2);
overflow:hidden;vertical-align:middle;margin-right:7px}
.barra i{display:block;height:100%;background:var(--accent);border-radius:4px}
.barra i.male{background:var(--crit)}
.barra i.medio{background:var(--warn)}
.barra i.bene{background:var(--ok)}
.niente{padding:26px 18px;color:var(--muted);font-size:13.5px}

/* pagina singola CBE */
.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;
box-shadow:var(--shadow);overflow:hidden}
.ph-head{padding:16px 18px 14px;border-bottom:1px solid var(--line-2)}
.ph-name{font-size:19px;font-weight:700;letter-spacing:-.01em;line-height:1.2}
.ph-sub{font-size:12.5px;color:var(--muted);margin-top:3px}
.sec{padding:0 0 16px;border-top:1px solid var(--line-2)}
.sec h3{padding:14px 18px 10px}
.bar-row{display:grid;grid-template-columns:78px 1fr 30px;align-items:center;gap:10px;
margin:0 18px 7px}
.bar-track{height:8px;background:var(--line-2);border-radius:4px;overflow:hidden}
.bar-fill{height:100%;border-radius:4px;background:var(--accent)}
.bar-fill.pmi{background:var(--pmi)}
.bar-lab{font-size:12.5px;color:var(--ink-2)}
.bar-num{font-size:12.5px;font-weight:600;text-align:right;font-variant-numeric:tabular-nums}
.fam{font-size:11px;color:var(--muted);margin:10px 18px 0;display:flex;gap:14px}
.fam i{font-style:normal;display:inline-flex;align-items:center;gap:5px}
.sw{width:8px;height:8px;border-radius:2px;background:var(--accent)}
.sw.pmi{background:var(--pmi)}
.picker{display:flex;flex-wrap:wrap;gap:6px;padding:12px 18px 0}
.picker button{font:500 12.5px/1 "Segoe UI",sans-serif;background:var(--surface);color:var(--ink-2);
border:1px solid var(--line);border-radius:999px;padding:7px 13px;cursor:pointer}
.picker button:hover{border-color:var(--accent);color:var(--accent)}
.picker button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.picker button[aria-pressed=true]{background:var(--accent);border-color:var(--accent);color:#fff}
.picker button.orphan{color:var(--crit);border-style:dashed}
.picker button.orphan[aria-pressed=true]{background:var(--crit);border-color:var(--crit);color:#fff}
.picker b{font-weight:600;opacity:.7;margin-left:5px}

.note{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:14px 16px 15px;
box-shadow:var(--shadow)}
.note.good{border-left:3px solid var(--accent)}
.note.flag{border-left:3px solid var(--gap)}
.note.alert{border-left:3px solid var(--crit)}
.note h2{font-size:14px;font-weight:600;margin:0 0 6px}
.note p{margin:0;font-size:13.5px;color:var(--ink-2);line-height:1.5}
.note p+p{margin-top:7px}
.note strong{color:var(--ink);font-weight:600}
.avvisi{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px}
footer{border-top:1px solid var(--line);padding-top:15px;font-size:12.5px;color:var(--muted);max-width:80ch}
@media print{body{background:#fff;padding:0}.tabs,.filtri,.picker,.avvisi,footer{display:none}
.pannello,.card{border:none;box-shadow:none}}
"""


def e(s):
    return html.escape(str(s), quote=False)


def num(x, vuoto='<span class="vuota">&mdash;</span>'):
    return vuoto if x is None else "%d" % x


def dec(x):
    return '<span class="vuota">&mdash;</span>' if x is None else ("%.1f" % x).replace(".", ",")


def mil(x):
    """1500 -> 1.500, come si scrivono i numeri in italiano."""
    if x is None:
        return '<span class="vuota">&mdash;</span>'
    return "{:,}".format(int(x)).replace(",", ".")


def barra_perc(perc, atteso):
    """Barra di avanzamento, colorata rispetto a quanto ci si aspetterebbe a oggi."""
    if perc is None:
        return '<span class="vuota">&mdash;</span>'
    classe = "bene" if perc >= atteso else ("medio" if perc >= atteso * 0.75 else "male")
    return ('<span class="barra"><i class="%s" style="width:%d%%"></i></span>%d%%'
            % (classe, min(perc, 100), perc))


def stato(p, v, atteso):
    """Il giudizio su una famiglia (o sul totale) di un negozio."""
    if p["chiuso"]:
        return "chiuso"
    # "fermo" viene prima di tutto: un negozio a zero e' da guardare anche
    # senza target, ed e' l'unica cosa visibile finche' i TGT non arrivano.
    if v["mese"] == 0:
        return "fermo"
    # attenzione: un target puo' valere ZERO (i corner non hanno mobile), e in
    # quel caso scarto/perc restano None. Zero e "nessun target" vanno trattati
    # allo stesso modo: non c'e' niente rispetto a cui misurarsi.
    if not v["tgt"] or v["scarto"] is None:
        return "senzatgt"
    if v["scarto"] >= 0:
        return "bene"
    return "sotto" if v["perc"] >= atteso * 0.75 else "critico"


ETICHETTA = {"chiuso": ('<span class="chip">chiuso</span>'),
             "fermo": '<span class="chip rosso">fermo</span>',
             "senzatgt": '<span class="chip">no TGT</span>',
             "bene": '<span class="chip verde">in linea</span>',
             "sotto": '<span class="chip giallo">sotto</span>',
             "critico": '<span class="chip rosso">critico</span>'}


def cella_delta(d):
    if d is None:
        return '<span class="zero">&ndash;</span>'
    if d > 0:
        return '<span class="up">+%d</span>' % d
    if d < 0:
        return '<span class="giu">%d</span>' % d
    return '<span class="zero">0</span>'


# Le colonne per ciascuna famiglia. "base" = quelle che restano visibili anche
# nella vista d'insieme, dove le tre famiglie stanno affiancate.
COLONNE = [("tgt", "TGT", True), ("mese", "Fatto", True), ("perc", "Avanz.", False),
           ("mancano", "Mancano", True), ("serve", "Serve/gg", False),
           ("passo", "Tenuto/gg", False), ("proiezione", "Proiez.", False),
           ("oggi", "Oggi", False), ("delta", "Ult. ore", False)]


def celle_voce(v, atteso, cls):
    """Le nove celle di una famiglia, tutte emesse e mostrate via CSS."""
    valori = {
        "tgt": num(v["tgt"]), "mese": "%d" % v["mese"],
        "perc": barra_perc(v["perc"], atteso), "mancano": num(v["mancano"]),
        "serve": dec(v["serve"]), "passo": dec(v["passo"]),
        "proiezione": "%d" % v["proiezione"], "oggi": "%d" % v["oggi"],
        "delta": cella_delta(v["delta"]),
    }
    return "".join('<td class="num %s k-%s%s">%s</td>'
                   % (cls, chiave, " k-base" if base else "", valori[chiave])
                   for chiave, _, base in COLONNE)


def chiave_ordine(p, quale):
    """Dal piu' indietro al piu' avanti, sulla famiglia scelta.
    I chiusi in fondo: non sono negozi da recuperare."""
    v = p[quale]
    return (1 if p["chiuso"] else 0,
            0 if v["tgt"] else 1,
            v["scarto"] if v["scarto"] is not None else 0,
            -v["mese"])


def tabella_negozi(pdv, atteso, con_cbe=True, tid=None, quale="tot"):
    """La vista che risponde a: su che negozi intervenire, e su quale pista."""
    righe = []
    for codice, p in sorted(pdv.items(), key=lambda kv: chiave_ordine(kv[1], quale)):
        stati = " ".join('data-stato-%s="%s"' % (k.lower(), stato(p, p[k], atteso))
                         for k in ["tot"] + FAMIGLIE)
        etichette = "".join('<td class="c-%s">%s</td>'
                            % (k.lower(), ETICHETTA[stato(p, p[k], atteso)])
                            for k in ["tot"] + FAMIGLIE)
        righe.append(
            '<tr data-cbe="%s" %s data-cerca="%s">'
            '<td><span class="loc">%s</span> <span class="code">%s</span>%s</td>%s%s%s</tr>'
            % (e(p["cbe"] or SENZA_CBE), stati, e((p["loc"] + " " + codice).lower()),
               e(p["loc"] or "?"), e(codice),
               '<br><span class="sub">%s</span>' % e(nome_proprio(p["cbe"])) if con_cbe and p["cbe"] else "",
               '<td><span class="chip">%s</span></td>' % e(p["canale"]) if p["canale"] else "<td></td>",
               "".join(celle_voce(p[k], atteso, "c-" + k.lower()) for k in ["tot"] + FAMIGLIE),
               etichette))

    # intestazioni: due righe, la prima raggruppa per famiglia
    gruppi = ['<th rowspan="2" class="sortable" data-col="0">Negozio<span class="fr"></span></th>',
              '<th rowspan="2">Canale</th>']
    testate, i = [], 2
    quante_base = sum(1 for _, _, base in COLONNE if base)
    for k in ["tot"] + FAMIGLIE:
        cls = "c-" + k.lower()
        # due intestazioni di gruppo: una larga per la vista su una pista sola,
        # una stretta per la vista affiancata, che mostra meno colonne
        gruppi.append('<th class="%s gruppo" colspan="%d">%s</th>'
                      % (cls, len(COLONNE), "Totale" if k == "tot" else k))
        if k != "tot":
            gruppi.append('<th class="%s gruppo g-base" colspan="%d">%s</th>'
                          % (cls, quante_base, k))
        for chiave, titolo, base in COLONNE:
            testate.append('<th class="num sortable %s k-%s%s" data-col="%d" data-n="1">%s'
                           '<span class="fr"></span></th>'
                           % (cls, chiave, " k-base" if base else "", i, titolo))
            i += 1
    for k in ["tot"] + FAMIGLIE:
        gruppi.append('<th rowspan="2" class="c-%s">Stato</th>' % k.lower())

    return """
<div class="tbl-wrap"><table%s class="v-tot"><thead>
<tr>%s</tr><tr>%s</tr></thead><tbody>%s</tbody></table></div>%s
""" % (' id="%s"' % tid if tid else "", "".join(gruppi), "".join(testate), "".join(righe),
       '\n<div class="niente" id="nessuno" hidden>Nessun negozio con questi filtri.</div>' if tid else "")


def blocco_piste(d, atteso, residui, prec, giornate_curva=0):
    """Il progressivo del mese, diviso nelle tre piste. E' la lettura principale:
    il totale c'e' ma sta in fondo e in sordina, perche' nasconde quale leva tirare."""
    carte = []
    for k in FAMIGLIE + ["tot"]:
        v = d[k]
        totale = (k == "tot")
        colore = "" if totale else (" energia" if k == "Energia" else " pmi")

        righe = ['<div class="p-riga"><span>Oggi</span><b>%s</b>%s</div>'
                 % (mil(v["oggi"]),
                    (' <i>%s</i>' % cella_delta(v["delta"])) if v["delta"] is not None else "")]
        if v["oggi_stima"] is not None:
            righe.append('<div class="p-riga stima"><span>In rotta per</span><b>%s</b>'
                         '<em>stasera &middot; fatto il %d%% della giornata</em></div>'
                         % (mil(v["oggi_stima"]), round(100 * v["quota"])))
        righe.append('<div class="p-riga"><span>Passo tenuto</span><b>%s</b><em>al giorno%s</em></div>'
                     % (dec(v["passo"]), ", oggi proiettato" if v["proiettato"] else ""))
        if v["tgt"]:
            indietro = v["serve"] is not None and v["serve"] > v["passo"]
            segno = "+" if v["scarto"] > 0 else ("&minus;" if v["scarto"] < 0 else "")
            righe.append('<div class="p-riga"><span>Obiettivo</span><b>%s</b><em>%d%% fatto</em></div>'
                         % (mil(v["tgt"]), v["perc"]))
            righe.append('<div class="p-riga"><span>Mancano</span><b>%s</b></div>' % mil(v["mancano"]))
            righe.append('<div class="p-riga%s"><span>Passo da tenere</span><b>%s</b><em>%s</em></div>'
                         % (" allarme" if indietro else " ok", dec(v["serve"]),
                            "al giorno, sui %d che restano" % residui if residui
                            else "e' l'ultimo giorno del mese"))
            righe.append('<div class="p-riga%s"><span>Proiezione</span><b>%s</b>'
                         '<em>%s%s sul TGT</em></div>'
                         % (" allarme" if v["scarto"] < 0 else " ok", mil(v["proiezione"]),
                            segno, mil(abs(v["scarto"]))))
        else:
            righe.append('<div class="p-riga"><span>Proiezione</span><b>%s</b>'
                         '<em>a fine mese</em></div>' % mil(v["proiezione"]))
            righe.append('<div class="p-riga vuoto"><span>Obiettivo</span><b>&mdash;</b>'
                         '<em>TGT da inserire</em></div>')

        avanzamento = ('<div class="p-avanz">%s <span>a oggi ci si aspetta il %d%%</span></div>'
                       % (barra_perc(v["perc"], atteso), atteso)) if v["tgt"] else ""

        carte.append(
            '<div class="pista%s%s"><div class="p-nome">%s</div>'
            '<div class="p-mese">%s <small>pz nel mese</small></div>%s'
            '<div class="p-righe">%s</div></div>'
            % (colore, " totale" if totale else "",
               "Totale delle tre" if totale else e(k),
               mil(v["mese"]), avanzamento, "".join(righe)))

    return ('<div class="piste">%s</div>'
            '<p class="spiega piede">Energia somma TIM GAS e TIM LUCE, perche\' il target arriva accorpato. '
            'Il <strong>passo da tenere</strong> e\' calcolato sui %d giorni di calendario ancora da fare, '
            'oggi escluso: quando supera il <strong>passo tenuto</strong>, quella pista sta perdendo terreno '
            'e la proiezione scende sotto il target.%s%s</p>'
            % ("".join(carte), residui,
               (" <strong>In rotta per</strong> stima come chiudera' la giornata in corso, confrontando "
                "l'ora dell'invio con l'andamento di %d giornate gia' chiuse in archivio: a quell'ora, "
                "in media, ne era stata fatta quella quota. Ogni pista ha la sua curva." % giornate_curva)
               if giornate_curva >= 2 else
               " Serviranno almeno due giornate intere in archivio per stimare come chiude quella in corso.",
               " Il confronto &laquo;ultime ore&raquo; e' con l'invio delle %s." % prec.strftime("%H:%M")
               if prec else ""))


def blocco_cbe(nome, c, prec, corr, atteso, giornate, residui, posizione, totali, giornate_curva=0):
    pdv = dict(sorted(c["pdv"].items(), key=lambda kv: chiave_ordine(kv[1], "tot")))
    attivi = sum(1 for p in pdv.values() if p["tot"]["mese"] > 0)
    fermi = sum(1 for p in pdv.values() if p["tot"]["mese"] == 0 and not p["chiuso"])
    t = c["tot"]

    stamp = ('<div class="stamp">%s &middot; ore %s &nbsp;&rarr;&nbsp; confronto con le %s</div>'
             % (corr.strftime("%d/%m/%Y"), corr.strftime("%H:%M"), prec.strftime("%H:%M"))
             if prec else
             '<div class="stamp warn">%s &middot; ore %s &nbsp;&nbsp;primo invio della giornata</div>'
             % (corr.strftime("%d/%m/%Y"), corr.strftime("%H:%M")))

    posiz = ("" if nome == SENZA_CBE else
             ' &middot; %d&deg; su %d per pezzi totali' % (posizione, len(totali)))

    return """
<div class="ph-head">
  <div class="ph-name">%s</div>
  <div class="ph-sub">%d negozi &middot; %d hanno prodotto, %d fermi a zero%s</div>
  %s
</div>
<div class="sec"><h3>Il suo mese, pista per pista &middot; %d giornate archiviate</h3>
  <div style="padding:0 18px">%s</div></div>
<div class="sec"><h3>I suoi negozi, dal piu' indietro al piu' avanti</h3>%s</div>
""" % (e(nome_proprio(nome)), len(pdv), attivi, fermi, posiz, stamp,
       giornate, blocco_piste(c, atteso, residui, prec, giornate_curva),
       tabella_negozi(pdv, atteso, con_cbe=False))


# --------------------------------------------------------------------- main

def analizza():
    """Legge l'archivio e calcola tutto. Torna un dizionario con i dati pronti,
    cosi' se ne servono anche altri programmi (le bozze per i CBE) senza
    ricostruire il ragionamento una seconda volta.
    Torna 1 invece del dizionario quando non c'e' niente da leggere."""
    print()
    print("  CRUSCOTTO REALTIME - TIM RETAIL")
    print("  " + "-" * 62)

    CARTELLA_REPORT.mkdir(parents=True, exist_ok=True)
    # la tilde iniziale marca sia i file temporanei di Excel (~$) sia gli scarti
    # dello scaricamento (~scaricato-): non sono report
    file_xlsx = [p for p in CARTELLA_REPORT.glob("*.xlsx") if not p.name.startswith("~")]
    if not file_xlsx:
        print("\n  Nella cartella non c'e' nessun report.")
        print("  Scarica gli allegati qui dentro e riprova:\n  %s\n" % CARTELLA_REPORT)
        return 1

    snapshot = {}
    for percorso in sorted(file_xlsx):
        try:
            istante = ora_estrazione(percorso)
            percorso = rinomina(percorso, istante)
            righe = leggi_righe(percorso)
        except Exception as errore:
            print("  ! salto %s -> %s" % (percorso.name, errore))
            continue
        if istante in snapshot:
            print("  = %s doppione delle %s, ignorato" % (percorso.name, istante.strftime("%H:%M")))
            continue
        snapshot[istante] = righe
        print("  . %-40s %s  (%d righe)" % (percorso.name[:40], istante.strftime("%d/%m %H:%M"), len(righe)))

    if not snapshot:
        print("\n  Nessun report leggibile.\n")
        return 1

    istanti = sorted(snapshot)
    corrente = istanti[-1]
    righe = snapshot[corrente]
    mese = max((r["mese"] for r in righe), key=chiave_mese)
    giorno = max(r["giorno"] for r in righe if r["mese"] == mese)

    # progressivo del mese: per ogni giorno, l'ultimo invio che lo contiene
    ultima_del_giorno = {}
    for istante in istanti:
        for r in snapshot[istante]:
            k = (r["mese"], r["giorno"])
            if k not in ultima_del_giorno or istante > ultima_del_giorno[k]:
                ultima_del_giorno[k] = istante
    righe_mese = []
    for (m, g), istante in sorted(ultima_del_giorno.items()):
        if m == mese:
            righe_mese += [r for r in snapshot[istante] if r["mese"] == m and r["giorno"] == g]

    giorni_archiviati = sorted(g for (m, g) in ultima_del_giorno if m == mese)
    perse = leggi_giornate_perse().get(chiave_mese(mese), set())
    buchi = [g for g in ("%02d" % n for n in range(1, int(giorno) + 1))
             if g not in giorni_archiviati and g not in perse]
    perse_viste = sorted(g for g in perse if g not in giorni_archiviati)

    # confronto con l'invio precedente DELLA STESSA GIORNATA
    stesso = [i for i in istanti if i < corrente
              and any(r["mese"] == mese and r["giorno"] == giorno for r in snapshot[i])]
    precedente = stesso[-1] if stesso else None

    perimetro = leggi_perimetro()
    pdv = costruisci(righe_mese, giorno, perimetro)
    if precedente:
        prima = costruisci([r for r in snapshot[precedente] if r["mese"] == mese and r["giorno"] == giorno],
                           giorno, {})
        for codice, p in pdv.items():
            for chiave in ["tot"] + FAMIGLIE:
                avanti = prima[codice][chiave]["oggi"] if codice in prima else 0
                p[chiave]["delta"] = p[chiave]["oggi"] - avanti

    # il calendario: giorni fatti, giorni che restano
    numero_mese = chiave_mese(mese)
    anno = corrente.year
    giorni_mese = monthrange(anno, numero_mese)[1]
    trascorsi = int(giorno)
    residui = max(giorni_mese - trascorsi, 0)
    # a che punto dell'obiettivo si dovrebbe essere oggi, a ritmo costante
    atteso = round(100 * trascorsi / giorni_mese)

    # la curva della giornata, ricavata dalle giornate chiuse in archivio
    punti, giornate_curva = curva_giornata(snapshot, (mese, giorno))
    minuti_ora = corrente.hour * 60 + corrente.minute
    quote = {k: stima_quota(punti, k, minuti_ora) for k in ["tot"] + FAMIGLIE}

    for p in pdv.values():
        calcola_tutto(p, trascorsi, residui, giorni_mese, quote)

    cbe = raggruppa_per_cbe(pdv, trascorsi, residui, giorni_mese, quote)
    reali = sorted([n for n in cbe if n not in (SENZA_CBE, DA_ASSEGNARE)],
                   key=lambda n: -cbe[n]["tot"]["mese"])
    totali = [cbe[n]["tot"]["mese"] for n in reali]
    ordine = reali + [n for n in (DA_ASSEGNARE, SENZA_CBE) if n in cbe]

    # ---- numeri di testa
    dentro = {k: v for k, v in pdv.items()
              if v["cbe"] and v["cbe"] not in (SENZA_CBE, DA_ASSEGNARE)}
    attesa = {k: v for k, v in pdv.items() if v["cbe"] == DA_ASSEGNARE}
    chiusi = [k for k, v in dentro.items() if v["chiuso"]]
    attivi = {k: v for k, v in dentro.items() if not v["chiuso"]}
    globale = somma_voci(dentro.values(), trascorsi, residui, giorni_mese, quote)
    t = globale["tot"]
    fuori = sum(p["tot"]["mese"] for p in pdv.values() if not p["cbe"] or p["cbe"] == SENZA_CBE)
    fermi = [k for k, p in attivi.items() if p["tot"]["mese"] == 0]

    return {"snapshot": snapshot, "istanti": istanti, "corrente": corrente,
            "precedente": precedente, "mese": mese, "giorno": giorno,
            "giorni_archiviati": giorni_archiviati, "buchi": buchi, "perse": perse_viste,
            "trascorsi": trascorsi, "residui": residui, "giorni_mese": giorni_mese,
            "atteso": atteso, "giornate_curva": giornate_curva, "quote": quote,
            "perimetro": perimetro, "pdv": pdv, "cbe": cbe, "reali": reali,
            "totali": totali, "ordine": ordine, "dentro": dentro, "attivi": attivi,
            "chiusi": chiusi, "globale": globale, "fuori": fuori, "fermi": fermi,
            "attesa": attesa}


def main(apri=True):
    """apri=False quando gira dall'Utilita' di pianificazione: rigenera i file
    ma non spalanca il browser ogni volta."""
    d = analizza()
    if not isinstance(d, dict):
        return d
    (snapshot, istanti, corrente, precedente, mese, giorno, giorni_archiviati, buchi,
     trascorsi, residui, giorni_mese, atteso, giornate_curva, quote, perimetro, pdv,
     cbe, reali, totali, ordine, dentro, attivi, chiusi, globale, fuori, fermi, attesa) = (
        d["snapshot"], d["istanti"], d["corrente"], d["precedente"], d["mese"], d["giorno"],
        d["giorni_archiviati"], d["buchi"], d["trascorsi"], d["residui"], d["giorni_mese"],
        d["atteso"], d["giornate_curva"], d["quote"], d["perimetro"], d["pdv"], d["cbe"],
        d["reali"], d["totali"], d["ordine"], d["dentro"], d["attivi"], d["chiusi"],
        d["globale"], d["fuori"], d["fermi"], d["attesa"])
    t = globale["tot"]

    sommario = """
<div class="sommario">
  <div class="som"><div class="k">Aggiornato all'invio delle</div><div class="v">%s</div>
    <div class="n">del %s &middot; %s</div></div>
  <div class="som"><div class="k">Giornate archiviate</div><div class="v">%d <small>su %d</small></div>
    <div class="n">%s</div></div>
  <div class="som"><div class="k">Giorni che restano</div><div class="v">%d</div>
    <div class="n">a ritmo costante si dovrebbe essere al %d%%</div></div>
  <div class="som %s"><div class="k">Negozi fermi</div><div class="v">%d <small>su %d</small></div>
    <div class="n">nessun pezzo in tutto il mese%s</div></div>
</div>""" % (corrente.strftime("%H:%M"), corrente.strftime("%d/%m"),
             ("confronto con le %s" % precedente.strftime("%H:%M")) if precedente
             else "primo invio della giornata",
             len(giorni_archiviati), trascorsi,
             nota_giornate(buchi, d["perse"]),
             residui, atteso,
             "male" if fermi else "buono", len(fermi), len(attivi),
             ", %d chiusi esclusi" % len(chiusi) if chiusi else "")

    # ---- filtri della vista negozi
    opzioni = "".join('<option value="%s">%s</option>' % (e(n), e(nome_proprio(n))) for n in ordine)
    filtri = """
<div class="filtri">
  <select id="ffam" aria-label="Scegli la pista">
    <option value="tot">Totale piste</option>
    <option value="tutte">Le tre piste affiancate</option>
    <option value="fisso">Solo Fisso</option>
    <option value="mobile">Solo Mobile</option>
    <option value="energia">Solo Energia</option>
  </select>
  <input type="search" id="cerca" placeholder="Cerca negozio o codice" aria-label="Cerca negozio">
  <select id="fcbe" aria-label="Filtra per CBE"><option value="">Tutti i CBE</option>%s</select>
  <select id="fstato" aria-label="Filtra per stato">
    <option value="">Tutti gli stati</option>
    <option value="fermo">Solo fermi a zero</option>
    <option value="critico">Solo critici</option>
    <option value="sotto">Solo sotto</option>
    <option value="bene">Solo in linea</option>
  </select>
  <span class="conteggio" id="conteggio"></span>
</div>""" % opzioni

    # ---- vista CBE
    blocchi, bottoni = [], []
    for i, nome in enumerate(ordine):
        pos = reali.index(nome) + 1 if nome in reali else 0
        blocchi.append('<div class="card" id="c%d"%s>%s</div>'
                       % (i, "" if i == 0 else " hidden",
                          blocco_cbe(nome, cbe[nome], precedente, corrente, atteso,
                                     len(giorni_archiviati), residui, pos, totali, giornate_curva)))
        dl = cbe[nome]["tot"]["delta"]
        bottoni.append('<button type="button" data-i="%d" aria-pressed="%s"%s>%s%s</button>'
                       % (i, "true" if i == 0 else "false",
                          ' class="orphan"' if nome in (SENZA_CBE, DA_ASSEGNARE) else "",
                          e(nome_proprio(nome)), "<b>+%d</b>" % dl if dl else ""))

    # ---- avvisi
    avvisi = []
    if not t["tgt"]:
        avvisi.append('<div class="note flag"><h2>Target non ancora inseriti</h2><p>Le colonne <strong>TGT, '
                      'Mancano, Serve/gg</strong> e la proiezione confrontata restano vuote finche\' non compili '
                      '<strong>obiettivi PDV.csv</strong>. Il file contiene gia\' i %d negozi del perimetro: '
                      'ti restano da riempire solo le tre colonne dei target. Nessun numero viene inventato al '
                      'posto tuo.</p></div>' % len(perimetro))
    if buchi:
        avvisi.append('<div class="note alert"><h2>Mancano i giorni %s</h2><p>Quei pezzi non sono contati da '
                      'nessuna parte e il passo tenuto risulta piu\' basso del vero. Se hai ancora le mail, '
                      'scarica un invio di quelle giornate.</p></div>' % ", ".join(buchi))
    if d["perse"]:
        avvisi.append('<div class="note flag"><h2>Il giorno %s non e&rsquo; mai arrivato</h2>'
                      '<p>Il report di quella giornata non e&rsquo; stato spedito e non lo sara&rsquo;: '
                      'e&rsquo; stata dichiarata <strong>giornata persa</strong>. I pezzi venduti quel '
                      'giorno <strong>non sono conteggiati</strong> da nessuna parte, quindi il '
                      'progressivo del mese e le proiezioni qui sopra sono piu&rsquo; bassi del vero '
                      'di circa una giornata.</p></div>' % ", ".join(d["perse"]))
    if attesa:
        tgt_attesa = sum(v["tot"]["tgt"] or 0 for v in attesa.values())
        avvisi.append('<div class="note alert"><h2>%d negozi senza CBE assegnato</h2>'
                      '<p>Sono nel perimetro e hanno un target — <strong>%d pezzi</strong> — ma non '
                      'sono attribuiti a nessuno, quindi restano fuori dai totali per CBE qui sopra. '
                      'Li trovi nel pulsante <em>Da assegnare</em>. Si sistemano scrivendo il nome '
                      'nella colonna CBE di <strong>obiettivi PDV.csv</strong>.</p></div>'
                      % (len(attesa), tgt_attesa))
    if fuori:
        avvisi.append('<div class="note flag"><h2>%d pezzi fuori perimetro</h2><p>Arrivano da negozi che il '
                      'report manda ma che non sono fra i tuoi %d. Sono esclusi da tutti i totali qui sopra.</p>'
                      '</div>' % (fuori, len(perimetro) if perimetro else 0))
    if not perimetro:
        avvisi.append('<div class="note alert"><h2>Manca il file del perimetro</h2><p>Senza '
                      '<strong>obiettivi PDV.csv</strong> vedi solo i negozi che hanno prodotto qualcosa: '
                      'quelli fermi a zero non compaiono nel report e restano invisibili.</p></div>')

    import json
    st = leggi_stato()
    controllo = st.get("ultimo_controllo") or datetime.now().replace(microsecond=0).isoformat()
    stato_pagina = json.dumps({
        "dato": corrente.isoformat(),
        "oraDato": corrente.strftime("%H:%M del %d/%m"),
        "controllo": controllo,
        "oraControllo": datetime.fromisoformat(controllo).strftime("%H:%M del %d/%m"),
        "oraPagina": datetime.now().strftime("%H:%M"),
        "esito": st.get("esito", "ok"),
        "messaggio": st.get("messaggio", ""),
        "automatico": automatico_attivo(),
    }, ensure_ascii=False).replace("<", "\\u003c")   # nessun messaggio puo' chiudere lo <script>

    pagina = """<!doctype html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cruscotto TIM Retail - %s</title><style>%s</style></head><body>
<div class="wrap">
<header class="masthead"><p class="eyebrow">Commessa TIM Retail &middot; %s &middot; giorno %d di %d</p>
<h1>Quanto manca, e dove</h1>
<p>Il progressivo del mese, <strong>pista per pista</strong>. I numeri escludono il doppio conteggio di
&laquo;TOT GAS e LUCE&raquo; e i negozi fuori perimetro.</p></header>
<div class="allarme-fascia" id="allarme" role="status"></div>
%s
<div class="pannello" style="padding:16px 18px 18px">%s</div>
<div class="avvisi">%s</div>
<div class="tabs" role="tablist">
  <button role="tab" aria-selected="true" data-v="negozi">Negozi</button>
  <button role="tab" aria-selected="false" data-v="cbe">CBE</button>
</div>
<div id="v-negozi">
  <div class="pannello"><h3>Dove intervenire</h3>
  <p class="spiega">Ordinati dal piu' indietro: in cima i negozi la cui proiezione a fine mese sta piu' sotto
  il target. Il primo menu sceglie <strong>su quale pista</strong> guardare: le colonne cambiano di conseguenza,
  e cosi' l'ordinamento e lo stato. Clicca le intestazioni per riordinare.</p>
  %s%s</div>
</div>
<div id="v-cbe" hidden>
  <div class="pannello"><h3>Scheda per CBE</h3><nav class="picker" id="picker">%s</nav></div>
  <div style="margin-top:18px">%s</div>
</div>
<footer>Generato il %s da %d invii in <code>report ricevuti\\realtime</code>. L'ora di estrazione e' letta dai
dati interni di ciascun file, non da una cella. Ogni report contiene solo la propria giornata: il progressivo
del mese e' ricostruito dall'archivio, che quindi va conservato. Il passo e' calcolato su giorni di
calendario, non lavorativi.</footer>
</div>
<script>
function q(s,r){return (r||document).querySelector(s)}
function qa(s,r){return [].slice.call((r||document).querySelectorAll(s))}

/* ---- l'allarme.
   Il controllo gira QUI, nel browser, ogni volta che apri o guardi la pagina.
   Se girasse quando la pagina viene generata non servirebbe a niente: proprio
   nel caso peggiore - lo strumento che ha smesso di funzionare - la pagina non
   verrebbe piu' riscritta, e l'allarme non comparirebbe mai. */
var STATO = %s;
function ore(da){ return (Date.now() - new Date(da).getTime()) / 3600000 }
function quanto(h){
  if (h < 1) return Math.round(h*60) + ' minuti';
  if (h < 24) return (h < 2 ? '1 ora' : Math.round(h) + ' ore');
  var g = Math.floor(h/24); return g === 1 ? '1 giorno' : g + ' giorni';
}
function controllaAllarme(){
  var f = q('#allarme'); if (!f) return;
  var hDato = ore(STATO.dato), adesso = new Date().getHours(), livello = '', titolo = '', testo = '';

  if (STATO.esito === 'errore'){
    livello = 'rosso';
    titolo = "Lo scaricamento dalla posta non funziona";
    testo = STATO.messaggio + ". I numeri qui sotto sono fermi all'invio delle "
          + STATO.oraDato + ', ' + quanto(hDato) + ' fa.';
  } else if (hDato > 14){
    livello = 'rosso';
    titolo = 'Nessun report nuovo da ' + quanto(hDato);
    testo = "L'ultimo e' quello delle " + STATO.oraDato + ". Di norma ne arrivano cinque al giorno, "
          + "l'ultimo verso le 22:30: mancarne cosi' tanti non e' normale.";
  } else if (hDato > 4 && adesso >= 12){
    livello = 'giallo';
    titolo = 'Manca un invio';
    testo = "L'ultimo report e' delle " + STATO.oraDato + ', ' + quanto(hDato) + ' fa. '
          + "A quest'ora dovrebbe esserne arrivato un altro.";
  } else if (STATO.automatico && ore(STATO.controllo) > 2){
    livello = 'giallo';
    titolo = 'Il controllo automatico non gira da ' + quanto(ore(STATO.controllo));
    testo = "Dovrebbe partire ogni 30 minuti. Il PC e' rimasto acceso?";
  }

  if (!livello){ f.className = 'allarme-fascia'; return }
  f.className = 'allarme-fascia mostra ' + livello;
  f.innerHTML = '<b>' + titolo + '</b>' + testo
    + '<small>Ultimo controllo della posta: ' + STATO.oraControllo
    + '. Pagina generata alle ' + STATO.oraPagina
    + ' &mdash; ricaricala per aggiornare i numeri.</small>';
}
controllaAllarme();
setInterval(controllaAllarme, 60000);

qa('.tabs button').forEach(function(b){b.addEventListener('click',function(){
  qa('.tabs button').forEach(function(x){x.setAttribute('aria-selected', x===b?'true':'false')});
  q('#v-negozi').hidden = b.dataset.v!=='negozi';
  q('#v-cbe').hidden    = b.dataset.v!=='cbe';
})});

var pk=q('#picker');
if(pk) pk.addEventListener('click',function(ev){
  var b=ev.target.closest('button'); if(!b)return;
  qa('button',pk).forEach(function(x){x.setAttribute('aria-pressed', x===b?'true':'false')});
  qa('#v-cbe .card').forEach(function(c){c.hidden = c.id!=='c'+b.dataset.i});
});

var tb=q('#tab-negozi');
if(tb){
  var corpo=q('tbody',tb), tutte=qa('tr',corpo);
  // la pista scelta decide quali colonne si vedono e quale stato si filtra
  function pista(){var v=q('#ffam').value; return v==='tutte'?'tot':v}
  function filtra(){
    var t=q('#cerca').value.trim().toLowerCase(), c=q('#fcbe').value, s=q('#fstato').value,
        p=pista(), n=0;
    // la stessa pista vale anche per le tabelle dentro le schede CBE
    qa('table.v-tot,table.v-tutte,table.v-fisso,table.v-mobile,table.v-energia')
      .forEach(function(x){x.className='v-'+q('#ffam').value});
    tutte.forEach(function(r){
      var ok=(!t||r.dataset.cerca.indexOf(t)>=0)&&(!c||r.dataset.cbe===c)&&
             (!s||r.getAttribute('data-stato-'+p)===s);
      r.hidden=!ok; if(ok)n++;
    });
    q('#conteggio').textContent=n+' negozi su '+tutte.length;
    q('#nessuno').hidden=n>0;
  }
  ['#ffam','#cerca','#fcbe','#fstato'].forEach(function(s){
    q(s).addEventListener('input',filtra); q(s).addEventListener('change',filtra);
  });
  // cambiando pista, riordina sulla colonna Mancano di quella pista
  q('#ffam').addEventListener('change',function(){
    var th=q('th.sortable.c-'+pista()+'.k-mancano',tb);
    if(th) th.click();
  });
  filtra();
  qa('th.sortable',tb).forEach(function(th){th.addEventListener('click',function(){
    var i=+th.dataset.col, numerico=th.dataset.n==='1',
        cre=th.getAttribute('aria-sort')!=='ascending';
    qa('th',tb).forEach(function(x){x.removeAttribute('aria-sort')});
    th.setAttribute('aria-sort', cre?'ascending':'descending');
    function val(r){
      var c=r.children[i], x=(c.dataset&&c.dataset.v)||c.textContent.trim();
      if(!numerico) return x.toLowerCase();
      var m=x.replace(/\\./g,'').replace(',','.').match(/-?\\d+(\\.\\d+)?/);
      return m?parseFloat(m[0]):-1e12;
    }
    tutte.slice().sort(function(a,b){
      var x=val(a),y=val(b);
      return (x<y?-1:x>y?1:0)*(cre?1:-1);
    }).forEach(function(r){corpo.appendChild(r)});
  })});
}
</script></body></html>""" % (
        corrente.strftime("%d-%m %H%M"), CSS, e(mese.split("-")[-1]), trascorsi, giorni_mese,
        sommario, blocco_piste(globale, atteso, residui, precedente, giornate_curva),
        "".join(avvisi), filtri, tabella_negozi(dentro, atteso, tid="tab-negozi"),
        "".join(bottoni), "".join(blocchi),
        datetime.now().strftime("%d/%m/%Y alle %H:%M"), len(istanti), stato_pagina)

    uscita = BASE / "cruscotto.html"
    uscita.write_text(pagina, encoding="utf-8")

    # ---- una pagina per CBE
    CARTELLA_PAGINE.mkdir(exist_ok=True)
    for nome in ordine:
        pos = reali.index(nome) + 1 if nome in reali else 0
        corpo = blocco_cbe(nome, cbe[nome], precedente, corrente, atteso,
                           len(giorni_archiviati), residui, pos, totali, giornate_curva)
        sicuro = re.sub(r"[^A-Za-z0-9 ']", "", nome_proprio(nome)).strip() or "senza nome"
        (CARTELLA_PAGINE / ("%s.html" % sicuro)).write_text(
            """<!doctype html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s - %s</title><style>%s</style></head><body>
<div class="wrap" style="max-width:900px"><div class="card">%s</div></div></body></html>"""
            % (e(nome_proprio(nome)), corrente.strftime("%d/%m %H:%M"), CSS, corpo), encoding="utf-8")

    # ---- i messaggi da incollare
    messaggi = []
    for nome in reali:
        c = cbe[nome]
        ct = c["tot"]
        muove = sorted([p for p in c["pdv"].values() if p["tot"]["oggi"] > 0],
                       key=lambda p: -p["tot"]["oggi"])[:3]
        fermi_c = [p for p in c["pdv"].values() if p["tot"]["mese"] == 0 and not p["chiuso"]]
        righe_msg = ["-- %s -- %s ore %s" % (nome_proprio(nome).upper(),
                                             corrente.strftime("%d/%m"), corrente.strftime("%H:%M"))]
        coda = ("  (%+d dalle %s)" % (ct["delta"], precedente.strftime("%H:%M"))
                if ct["delta"] is not None else "")
        righe_msg.append("Oggi: %d pz%s   Mese: %d pz in %d giornate"
                         % (ct["oggi"], coda, ct["mese"], len(giorni_archiviati)))
        # lo spaccato, una riga per pista
        for f in FAMIGLIE:
            v = c[f]
            if v["tgt"]:
                righe_msg.append("%-8s %3d/%-4d (%2d%%)  mancano %4d  ->  %s pz/gg   [tieni %s]"
                                 % (f, v["mese"], v["tgt"], v["perc"], v["mancano"],
                                    dec(v["serve"]), dec(v["passo"])))
            else:
                righe_msg.append("%-8s %3d pz   oggi %d   tieni %s pz/gg"
                                 % (f, v["mese"], v["oggi"], dec(v["passo"])))
        if ct["tgt"]:
            righe_msg.append("TOTALE   %3d/%-4d (%2d%%)  mancano %4d  ->  %s pz/gg per %d giorni"
                             % (ct["mese"], ct["tgt"], ct["perc"], ct["mancano"], dec(ct["serve"]), residui))
        if muove:
            righe_msg.append("Muove oggi: " + "  ".join("%s %d" % (p["loc"].title(), p["tot"]["oggi"])
                                                        for p in muove))
        # i negozi piu' indietro, e su quale pista lo sono
        indietro = []
        for p in c["pdv"].values():
            peggio = [(p[f]["scarto"], f) for f in FAMIGLIE if p[f]["scarto"] is not None and p[f]["scarto"] < 0]
            if peggio:
                s, f = min(peggio)
                indietro.append((s, "%s %s %d" % (p["loc"].title(), f, s)))
        if indietro:
            righe_msg.append("Da recuperare: " + "  |  ".join(x[1] for x in sorted(indietro)[:3]))
        if fermi_c:
            righe_msg.append("Fermi a zero (%d): %s" % (len(fermi_c),
                                                        "  ".join(p["loc"].title() for p in fermi_c[:4])))
        messaggi.append("\n".join(righe_msg))
    (BASE / "messaggi.txt").write_text("\n\n".join(messaggi) + "\n", encoding="utf-8")

    # ---- riepilogo a video
    print("  " + "-" * 62)
    print("  Invii letti      %d   (il piu' recente: %s)" % (len(istanti), corrente.strftime("%d/%m alle %H:%M")))
    print("  Confronto        %s" % ("con le %s di oggi" % precedente.strftime("%H:%M") if precedente
                                     else "non disponibile, e' il primo invio di oggi"))
    print("  Giornate         %d su %d%s" % (len(giorni_archiviati), trascorsi,
                                             "   ATTENZIONE: mancano i giorni %s" % ", ".join(buchi) if buchi else ""))
    print("  Perimetro        %d negozi%s" % (len(perimetro), "" if perimetro else "   ATTENZIONE: file assente"))
    if fuori:
        print("  Fuori perimetro  %d pezzi, esclusi dai totali" % fuori)
    print("  Negozi fermi     %d su %d%s" % (len(fermi), len(attivi),
                                             "   (%d chiusi esclusi)" % len(chiusi) if chiusi else ""))
    print()
    print("  %-9s %6s %6s %6s %8s %7s %9s %9s %8s" %
          ("PISTA", "TGT", "FATTO", "OGGI", "STASERA", "MANCANO", "SERVE/gg", "TENUTO/gg", "PROIEZ."))
    for k in FAMIGLIE + ["tot"]:
        v = globale[k]
        print("  %-9s %6s %6d %6d %8s %7s %9s %9s %8d"
              % ("TOTALE" if k == "tot" else k.upper(),
                 v["tgt"] if v["tgt"] else "-", v["mese"], v["oggi"],
                 ("~%d" % v["oggi_stima"]) if v["oggi_stima"] is not None else "chiusa",
                 v["mancano"] if v["mancano"] is not None else "-",
                 dec(v["serve"]) if v["serve"] is not None else "-",
                 dec(v["passo"]), v["proiezione"]))
    if globale["tot"]["oggi_stima"] is not None:
        print("\n  'STASERA' e' la stima di come chiude oggi, su %d giornate in archivio." % giornate_curva)
    if not t["tgt"]:
        print("\n  Le colonne TGT/MANCANO/SERVE si accendono compilando 'obiettivi PDV.csv'")
    else:
        print("\n  Passo da tenere calcolato sui %d giorni che restano." % residui)
    print()
    print("  Creati:  cruscotto.html   |   pagine per CBE\\ (%d)   |   messaggi.txt" % len(ordine))
    print()

    if apri:
        try:
            os.startfile(uscita)
        except Exception:
            print("  Apri a mano: %s" % uscita)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(apri="--non-aprire" not in sys.argv))
    except Exception as errore:
        print("\n  Si e' fermato per un errore imprevisto:")
        print("  %s: %s\n" % (type(errore).__name__, errore))
        input("  Premi Invio per chiudere.")
        sys.exit(1)
