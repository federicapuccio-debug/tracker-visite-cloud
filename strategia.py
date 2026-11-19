# -*- coding: utf-8 -*-
"""
Numeri per la scheda "Strategia" del tracker visite PDV.
=========================================================

Da lanciare quando arrivano i target del mese nuovo (dopo aver aggiornato
dati/ e visite_dati.py). Confronta i target del mese (in pezzi lordi: fisso
x1,10, mobile x1,17, energia com'e') con quanto gli stessi negozi hanno fatto
il mese prima nei report realtime di report/, e conta dove sono i pezzi che
mancano. Stampa i numeri: il testo della scheda si riscrive a mano.

Uso:  python strategia.py
"""

import json
import math

import visite_dati as V

V.GIORNI_STORICO = 400


def main():
    anno, mese = V.MESE
    prima = "%d-%02d" % ((anno, mese - 1) if mese > 1 else (anno - 1, 12))
    prod, _ = V.leggi_produzione()
    tgt, port = V.leggi_target(), V.leggi_portafogli()
    patti = json.loads(V.FILE_PATTI.read_text(encoding="utf-8"))["cbe"] if V.FILE_PATTI.exists() else {}
    focus = {(c, n["cdc"]): n["pista"] for c, p in patti.items() for n in p.get("negozi", [])}
    T, P, mancano, negozi = {k: 0 for k in "fme"}, [0, 0, 0], {k: [] for k in "fme"}, 0
    for cdc, p in port.items():
        cbe = p["cbe"]
        if not cbe or str(cbe).startswith("Senza"):
            continue
        negozi += 1
        fatto = [0, 0, 0]
        for g, v in prod.get(cdc, {}).items():
            if g.startswith(prima):
                for i in range(3):
                    fatto[i] += v[i]
        for i in range(3):
            P[i] += fatto[i]
        t = tgt.get(cdc)
        if not t:
            continue
        for i, k in enumerate("fme"):
            T[k] += t[k] or 0
            mancano[k].append((max(0, math.ceil((t[k] or 0) * V.MARKUP[k]) - fatto[i]), cbe, cdc))
    TT = {k: round(T[k] * V.MARKUP[k]) for k in "fme"}
    print("Negozi nei portafogli: %d" % negozi)
    print("Target del mese: fisso %d, mobile %d, fisso+mobile %d, energia %d" % (TT["f"], TT["m"], TT["f"] + TT["m"], TT["e"]))
    print("Fatto nel mese prima (%s): fisso %d, mobile %d, fisso+mobile %d, energia %d" % (prima, P[0], P[1], P[0] + P[1], P[2]))
    for i, (k, nome) in enumerate([("f", "fisso"), ("m", "mobile"), ("e", "energia")]):
        if P[i]:
            print("  %s: %+.0f%% rispetto al mese prima" % (nome, (TT[k] / P[i] - 1) * 100))
    for k, nome in [("f", "fisso"), ("m", "mobile"), ("e", "energia")]:
        righe = sorted(mancano[k], reverse=True)
        tot = sum(r[0] for r in righe)
        acc, n = 0, 0
        for r in righe:
            if acc >= tot / 2:
                break
            acc += r[0]
            n += 1
        nel_patto = sum(1 for r in righe[:n] if focus.get((r[1], r[2])) == k)
        print("Pezzi mancanti di %s rispetto al mese prima: %d; metà li fanno %d negozi (%d del patto su questa pista)" % (nome, tot, n, nel_patto))


if __name__ == "__main__":
    main()
