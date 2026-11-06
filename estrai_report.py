# -*- coding: utf-8 -*-
"""
Salva in "report" gli allegati Excel dei messaggi Gmail scaricati in formato RAW.
=================================================================================

Nel cloud i report realtime non li scarica Windows: li legge Claude con il
connettore Gmail (get_message con messageFormat RAW) e salva ogni risposta come
file JSON. Questo script apre quei file, tira fuori l'allegato .xlsx e lo salva
come "AAAA-MM-GG HHMM - Realtime CBE.xlsx", con l'ora letta dentro il file.

Per ogni giorno tiene solo il report piu' recente: ogni report contiene solo il
proprio giorno, quindi l'ultimo della giornata e' quello che conta.

Uso:  python estrai_report.py <file1.json> [<file2.json> ...]
"""

import base64
import email
import json
import sys
import tempfile
from pathlib import Path

import cruscotto as C

REPORT = Path(__file__).resolve().parent / "report"


def allegati(percorso):
    dati = json.loads(Path(percorso).read_text(encoding="utf-8"))
    raw = dati["raw"]
    try:
        grezzo = base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4))
    except Exception:
        grezzo = raw.encode("utf-8")
    for parte in email.message_from_bytes(grezzo).walk():
        nome = parte.get_filename() or ""
        if nome.lower().endswith(".xlsx"):
            yield parte.get_payload(decode=True)


def main(file_json):
    REPORT.mkdir(exist_ok=True)
    esistenti = {}
    for p in REPORT.glob("*.xlsx"):
        esistenti.setdefault(p.name[:10], []).append(p)
    nuovi = 0
    for percorso in file_json:
        for contenuto in allegati(percorso):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as t:
                t.write(contenuto)
            ora = C.ora_estrazione(t.name)
            Path(t.name).unlink()
            nome = ora.strftime("%Y-%m-%d %H%M") + " - Realtime CBE.xlsx"
            giorno = nome[:10]
            stesso = esistenti.get(giorno, [])
            if any(p.name >= nome for p in stesso):
                print("  gia' presente o piu' vecchio:", nome)
                continue
            for p in stesso:        # il nuovo sostituisce il report precedente dello stesso giorno
                p.unlink()
            (REPORT / nome).write_bytes(contenuto)
            esistenti[giorno] = [REPORT / nome]
            nuovi += 1
            print("  salvato:", nome)
    print("Report nuovi: %d" % nuovi)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
