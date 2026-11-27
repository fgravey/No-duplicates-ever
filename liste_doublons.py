#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""liste_doublons.py - Pour UNE racine : liste tous les doublons parfaits (hash identique)
et produit :
  - un CSV detaille (groupe, sha256, taille, keeper, tous les chemins)
  - un fichier texte "supprimer_sauf_keeper.txt" (un chemin par ligne, tous les
    exemplaires SAUF le keeper) pour supprimer en une commande, ex. :
        tr '\\n' '\\0' < supprimer_sauf_keeper.txt | xargs -0 rm -f --
Le script ne supprime RIEN. Le keeper choisi = le 1er fichier hors dossier de tri
(_DOUBLONS_A_TRIER etc.), sinon le 1er par ordre alpha. Verifie le CSV avant.

Usage :
  python3 liste_doublons.py "/Users/francoisgravey/Library/CloudStorage/OneDrive-unicaen.fr(2)"
  python3 liste_doublons.py RACINE --report-dir ~/Downloads/rapports_scan --include-empty
"""

import argparse
import csv
import os
import sys
import time
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dedup_scan import scanner_racine, charger_checkpoint, DOSSIERS_IGNORES


def main():
    ap = argparse.ArgumentParser(description="Liste les doublons parfaits d'une racine (CSV + liste de suppression). Ne supprime rien.")
    ap.add_argument("racine")
    ap.add_argument("--report-dir", default=os.path.expanduser("~/Downloads/rapports_scan"))
    ap.add_argument("--max-content", type=int, default=2 * 1024 * 1024 * 1024)
    ap.add_argument("--include-empty", action="store_true", help="inclure les fichiers vides (hash identique trivially)")
    args = ap.parse_args()

    racine = os.path.abspath(args.racine)
    if not os.path.isdir(racine):
        print("Racine introuvable :", racine, file=sys.stderr)
        return 1
    os.makedirs(args.report_dir, exist_ok=True)
    ckpt_path = os.path.join(args.report_dir, "checkpoint.tsv")
    ckpt = charger_checkpoint(ckpt_path)

    print(f"[*] Scan de {racine} (checkpoint: {len(ckpt)} entrees) ...")
    t0 = time.time()
    rows, erreurs = scanner_racine(racine, args.max_content, ckpt=ckpt, ckpt_fichier=ckpt_path)
    print(f"[*] {len(rows)} fichiers, {len(erreurs)} erreurs, en {time.time()-t0:.0f}s")

    groupes = defaultdict(list)
    for r in rows:
        if r["taille"] == 0 and not args.include_empty:
            continue
        groupes[r["sha256"]].append(r["chemin"])

    def est_tri(chemin):
        return any(seg in DOSSIERS_IGNORES for seg in chemin.split(os.sep))

    ts = time.strftime("%Y%m%d_%H%M%S")
    base = os.path.basename(racine.rstrip(os.sep)).replace(" ", "_")
    csv_path = os.path.join(args.report_dir, f"doublons_{base}_{ts}.csv")
    txt_path = os.path.join(args.report_dir, f"supprimer_sauf_keeper_{base}_{ts}.txt")

    n_groupes = n_copies = 0
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as fcsv, open(txt_path, "w", encoding="utf-8") as ftxt:
        w = csv.writer(fcsv)
        w.writerow(["groupe", "sha256", "taille", "nb_exemplaires", "keeper", "exemplaires"])
        for i, (h, chemins) in enumerate(sorted(groupes.items()), 1):
            if len(chemins) < 2:
                continue
            chemins.sort()
            hors_tri = [c for c in chemins if not est_tri(c)]
            keeper = hors_tri[0] if hors_tri else chemins[0]
            copies = [c for c in chemins if c != keeper]
            taille = next((r["taille"] for r in rows if r["chemin"] == keeper), 0)
            w.writerow([i, h, taille, len(chemins), keeper, " | ".join(chemins)])
            for c in copies:
                ftxt.write(c + "\n")
            n_groupes += 1
            n_copies += len(copies)

    print(f"[*] {n_groupes} groupes de doublons parfaits, {n_copies} exemplaires en trop")
    print(f"[*] CSV detaille   : {csv_path}")
    print(f"[*] Liste a supprimer (tous sauf keeper) : {txt_path}")
    print("[*] Apercu de la commande de suppression (A VERIFIER AVANT, le script ne supprime rien) :")
    print(f"      tr '\\n' '\\0' < '{txt_path}' | xargs -0 rm -f --")
    return 0


if __name__ == "__main__":
    sys.exit(main())
