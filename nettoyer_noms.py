#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""nettoyer_noms.py - corrige les noms de fichiers/dossiers incompatibles OneDrive/SharePoint :
  - caracteres interdits (~ " # % & * : < > ? / \\ |), noms reserves (CON, PRN...), point/espace final
  - chemins trop longs (limite OneDrive ~400 caracteres) : raccourcit les segments les plus longs
  - conflit de casse dans un meme dossier (Rapport.docx vs rapport.docx) : suffixe __2, __3...
Dry-run par defaut (plan CSV uniquement) ; --apply pour renommer. Aucune suppression.

Usage :
  python3 nettoyer_noms.py RACINE [RACINE ...]                 # plan
  python3 nettoyer_noms.py --apply RACINE [RACINE ...]         # applique
  python3 nettoyer_noms.py --auto                              # detecte les racines OneDrive/Proton/GoogleDrive/iCloud
"""

import argparse
import csv
import glob
import os
import re
import sys
import time

LIMITE_CHEMIN = 380  # marge sous la limite OneDrive de 400 caracteres
CARACTÈRES_INTERDITS = re.compile(r'[~"#%&*:<>?\\/|]')
NOMS_RESERVES = {"con", "prn", "aux", "nul"} | {f"com{i}" for i in range(1, 10)} | {f"lpt{i}" for i in range(1, 10)}


def nettoyer_nom(nom):
    """Retourne (nom corrige, modifie ?)."""
    change = False
    n = CARACTÈRES_INTERDITS.sub("_", nom)
    n = n.rstrip(" .")
    if not n or n == "." or n == "..":
        n = "_vide"
        change = True
    base = os.path.splitext(n)[0].strip().lower()
    if base in NOMS_RESERVES:
        n = "_" + n
        change = True
    if n != nom:
        change = True
    return n, change


def chemin_trop_long(chemin, limite):
    return len(chemin) > limite


def raccourcir_un_segment(chemin, limite):
    """Raccourcit le DERNIER segment du chemin (fichier ou dossier lui-meme) de 20 caracteres max.
    Le parent de destination existe donc toujours. Un passage = un renommage ;
    relancer le script fait converger tous les chemins sous la limite."""
    if len(chemin) <= limite:
        return None
    parties = chemin.split(os.sep)
    seg = parties[-1]
    base, ext = os.path.splitext(seg)
    cut = min(20, max(0, len(base) - 8))
    if cut == 0:
        return None
    parties[-1] = base[: len(base) - cut] + ext
    return os.sep.join(parties)


def traduire(p, mapping):
    """Remplace le prefix le plus long deja renomme."""
    best = None
    for old, new in mapping.items():
        if p == old or p.startswith(old + os.sep):
            if best is None or len(old) > len(best[0]):
                best = (old, new)
    if best:
        old, new = best
        return new + p[len(old):]
    return p


def unique_casse_insensible(dest):
    """Ajoute __1, __2... si un nom existe deja (meme en ignorant la casse)."""
    d, nom = os.path.split(dest)

    def existe_ci(p):
        if os.path.lexists(p):
            return True
        if os.path.isdir(d):
            cf = os.path.basename(p).casefold()
            return any(e.casefold() == cf for e in os.listdir(d))
        return False

    if not existe_ci(dest):
        return dest
    base, ext = os.path.splitext(nom)
    i = 1
    while existe_ci(os.path.join(d, f"{base}__{i}{ext}")):
        i += 1
    return os.path.join(d, f"{base}__{i}{ext}")


def detecter_racines_auto():
    racines = []
    for prefix in ("OneDrive", "ProtonDrive", "Proton", "GoogleDrive", "iCloud Drive", "iCloudDrive"):
        racines += sorted(glob.glob(os.path.expanduser(f"~/Library/CloudStorage/{prefix}*")))
        racines += sorted(glob.glob(os.path.expanduser(f"~/{prefix}*")))
    if os.name == "nt":
        up = os.path.expanduser("~")
        racines += sorted(glob.glob(os.path.join(up, "OneDrive*")))
    return [r for r in racines if os.path.isdir(r)]


def planifier(racine, limite, rows):
    racine = os.path.abspath(racine)
    # 1) conflits de casse par dossier
    for dirpath, dirnames, filenames in os.walk(racine):
        for liste in (filenames, dirnames):
            groupes = {}
            for e in liste:
                groupes.setdefault(e.casefold(), []).append(e)
            for cf, membres in groupes.items():
                if len(membres) > 1:
                    membres.sort()
                    for m in membres[1:]:
                        ancien = os.path.join(dirpath, m)
                        base, ext = os.path.splitext(m)
                        nouveau = unique_casse_insensible(os.path.join(dirpath, f"{base}__2{ext}"))
                        rows.append(["conflit_casse", ancien, nouveau])
        # 2) noms invalides + chemins trop longs (fichiers d'abord, puis dossiers)
        for e in filenames:
            p = os.path.join(dirpath, e)
            n, modifie = nettoyer_nom(e)
            if modifie:
                rows.append(["nom_invalide", p, os.path.join(dirpath, n)])
        for e in filenames:
            p = os.path.join(dirpath, e)
            if chemin_trop_long(p, limite):
                court = raccourcir_un_segment(p, limite)
                if court:
                    rows.append(["chemin_trop_long", p, court])
        for e in dirnames:
            p = os.path.join(dirpath, e)
            n, modifie = nettoyer_nom(e)
            if modifie:
                rows.append(["nom_invalide", p, os.path.join(dirpath, n)])
            if chemin_trop_long(p + os.sep + "x", limite):
                court = raccourcir_un_segment(p, limite)
                if court:
                    rows.append(["chemin_trop_long", p, court])


def main():
    ap = argparse.ArgumentParser(description="Corrige noms trop longs / invalides / conflits de casse (OneDrive). Dry-run par defaut.")
    ap.add_argument("racines", nargs="*")
    ap.add_argument("--auto", action="store_true", help="detecte les racines cloud automatiquement")
    ap.add_argument("--limite", type=int, default=LIMITE_CHEMIN, help="longueur max de chemin (defaut 380)")
    ap.add_argument("--report-dir", default=os.path.expanduser("~/Downloads/rapports_scan"))
    ap.add_argument("--apply", action="store_true", help="applique les renommages")
    args = ap.parse_args()

    racines = list(args.racines)
    if args.auto or not racines:
        racines = detecter_racines_auto()
    if not racines:
        print("Aucune racine. Indique des chemins ou --auto.", file=sys.stderr)
        return 1
    print("[*] Racines :")
    for r in racines:
        print("    " + r)

    rows = []
    for r in racines:
        print(f"[*] Analyse de {r} ...")
        planifier(r, args.limite, rows)

    ts = time.strftime("%Y%m%d_%H%M%S")
    os.makedirs(args.report_dir, exist_ok=True)
    plan_path = os.path.join(args.report_dir, f"plan_nettoyage_noms_{ts}.csv")
    with open(plan_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["type", "ancien", "nouveau", "resultat"])
        for r in rows:
            w.writerow(r + [""])
    from collections import Counter
    print(f"[*] {len(rows)} renommages prevus -> {plan_path}")
    for k, v in Counter(r[0] for r in rows).most_common():
        print(f"    {v:>7}  {k}")

    if args.apply:
        # de l'exterieur vers l'interieur, avec traduction des prefixes deja renommes
        rows.sort(key=lambda r: r[1].count(os.sep))
        mapping = {}
        ok = echec = 0
        for r in rows:
            type_, ancien, nouveau = r
            ancien2 = traduire(ancien, mapping)
            nouveau2 = traduire(nouveau, mapping)
            try:
                if os.path.lexists(nouveau2):
                    raise OSError("destination existe deja")
                os.rename(ancien2, nouveau2)
                mapping[ancien] = nouveau
                ok += 1
            except OSError as e:
                echec += 1
                print(f"    ECHEC {ancien2} -> {nouveau2} : {e}", file=sys.stderr)
        with open(plan_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["type", "ancien", "nouveau", "resultat"])
            for r in rows:
                type_, ancien, nouveau = r
                res = "OK" if os.path.lexists(traduire(nouveau, mapping)) else "ECHEC"
                w.writerow([type_, ancien, nouveau, res])
        print(f"[*] Renommages : {ok} OK, {echec} ECHEC")
    else:
        print("[*] Dry-run : aucun renommage. Relance avec --apply.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
