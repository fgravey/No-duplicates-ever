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
import unicodedata

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


def raccourcir_un_segment(chemin, limite, longueur=None):
    """Raccourcit le DERNIER segment du chemin (fichier ou dossier lui-meme).
    Le parent de destination existe donc toujours. Un passage = un renommage ;
    relancer le script fait converger tous les chemins sous la limite.
    longueur : longueur effective a considerer (ex. chemin + '/x' pour un dossier)."""
    L = len(chemin) if longueur is None else longueur
    if L <= limite:
        return None
    parties = chemin.split(os.sep)
    seg = parties[-1]
    base, ext = os.path.splitext(seg)
    cut = min(40, max(0, len(base) - 8))
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


def existe_ci(p):
    """True si p existe, ou si un nom identique en ignorant la casse existe dans le dossier parent."""
    if os.path.lexists(p):
        return True
    d = os.path.dirname(p)
    if os.path.isdir(d):
        cf = os.path.basename(p).casefold()
        return any(e.casefold() == cf for e in os.listdir(d))
    return False


def unique_casse_insensible(dest):
    """Ajoute __1, __2... si un nom existe deja (meme en ignorant la casse)."""
    if not existe_ci(dest):
        return dest
    d, nom = os.path.split(dest)
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
    return list(dict.fromkeys(r for r in racines if os.path.isdir(r)))


def cible_finale(dirpath, nom, limite, proxy=False):
    """Calcule le renommage unique d'un fichier/dossier : nettoyage du nom ET/OU
    raccourcissement si le chemin depasse la limite. Retourne (ancien, nouveau, raison) ou None."""
    p = os.path.join(dirpath, nom)
    n, modifie = nettoyer_nom(nom)
    p2 = os.path.join(dirpath, n) if modifie else p
    raison = "nom_invalide" if modifie else None
    if proxy:
        # marge de 40 caracteres pour un eventuel nom de fichier enfant :
        # sinon un dossier long avec des fichiers a noms courts n'est jamais raccourci
        longueur = len(p2) + 1 + 40
    else:
        longueur = len(p2)
    if longueur > limite:
        court = raccourcir_un_segment(p2, limite, longueur=longueur)
        if court:
            b = os.path.basename(court)
            b2, _ = nettoyer_nom(b)
            p2 = os.path.join(os.path.dirname(court), b2)
            raison = (raison + "+chemin_trop_long") if raison else "chemin_trop_long"
    if p2 == p:
        return None
    return p, unique_casse_insensible(p2), raison


def planifier(racine, limite, rows):
    racine = os.path.abspath(racine)
    for dirpath, dirnames, filenames in os.walk(racine):
        # 1) conflits de casse/Unicode par dossier : meme nom en ignorant la casse
        #    ET la forme de normalisation (NFC vs NFD) — macOS stocke en NFD, le
        #    serveur en NFC, d'ou les "meme nom ou majuscules differentes" fantomes
        for liste in (filenames, dirnames):
            groupes = {}
            for e in liste:
                groupes.setdefault(unicodedata.normalize("NFC", e).casefold(), []).append(e)
            for cf, membres in groupes.items():
                if len(membres) > 1:
                    membres.sort()
                    for m in membres[1:]:
                        ancien = os.path.join(dirpath, m)
                        base, ext = os.path.splitext(m)
                        nouveau = unique_casse_insensible(os.path.join(dirpath, f"{base}__2{ext}"))
                        rows.append(["conflit_casse", ancien, nouveau])
        # 2) une seule ligne par entree : nom invalide et/ou chemin trop long
        for e in filenames:
            r = cible_finale(dirpath, e, limite)
            if r:
                rows.append([r[2], r[0], r[1]])
        for e in dirnames:
            r = cible_finale(dirpath, e, limite, proxy=True)
            if r:
                rows.append([r[2], r[0], r[1]])


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
    if rows:
        print("[*] Plus longs chemins restants :")
        for r in sorted(rows, key=lambda r: -len(r[1]))[:5]:
            print(f"    {len(r[1]):>4}  {r[1]}")

    if args.apply:
        # de l'exterieur vers l'interieur, avec traduction des prefixes deja renommes.
        # Re-executable : une source absente = deja renommee ou disparue -> IGNORE.
        rows.sort(key=lambda r: r[1].count(os.sep))
        mapping = {}
        ok = echec = ignore = 0
        for r in rows:
            type_, ancien, nouveau = r
            ancien2 = traduire(ancien, mapping)
            if not os.path.lexists(ancien2):
                r.append("IGNORE (absent ou deja renomme)")
                ignore += 1
                continue
            nouveau2 = traduire(nouveau, mapping)
            if existe_ci(nouveau2):  # destination prise entre-temps -> nom unique frais
                nouveau2 = unique_casse_insensible(nouveau2)
            try:
                os.rename(ancien2, nouveau2)
                mapping[ancien] = nouveau
                r.append("OK")
                ok += 1
            except OSError as e:
                r.append(f"ECHEC: {e}")
                echec += 1
                print(f"    ECHEC {ancien2} -> {nouveau2} : {e}", file=sys.stderr)
        with open(plan_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["type", "ancien", "nouveau", "resultat"])
            for r in rows:
                w.writerow(r)
        print(f"[*] Renommages : {ok} OK, {echec} ECHEC, {ignore} ignores (deja faits)")
        if echec:
            print(f"[*] ECHEC detailles dans : {plan_path} (colonne resultat)")
    else:
        print("[*] Dry-run : aucun renommage. Relance avec --apply.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
