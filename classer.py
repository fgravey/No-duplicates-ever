#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""classer.py - Propose puis applique un classement des fichiers issus du scan :
  université -> OneDrive, personnel -> Proton, 100% CHU -> reste en local.

Securite :
  - dry-run par defaut (plan uniquement) ; --apply pour deplacer reellement
  - AUCUNE suppression
  - regle dure : toute donnee CHU/patients detectee dans un cloud n'est JAMAIS
    envoyee vers OneDrive/Proton -> signalee "A_TRAITER_MANUELLEMENT"
  - les dossiers de tri (_A_METTRE_DE_COTE, _DOUBLONS_A_TRIER, ...) sont ignores

Usage :
  python3 classer.py                          # plan (dry-run) depuis le dernier inventaire
  python3 classer.py --apply                  # applique les deplacements
  python3 classer.py --default proton         # tout ce qui ne matche aucune regle -> Proton
  python3 classer.py --regles mes_regles.tsv --inventaire .../inventaire_XXX.tsv
"""

import argparse
import csv
import glob
import os
import re
import shutil
import sys
import time
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dedup_scan import extraire_texte, EXT_TEXT, EXT_ZIP_XML, EXT_PDF  # reutilise l'extraction de texte

EXT_TEXTE_CLASSEMENT = EXT_TEXT | EXT_ZIP_XML | EXT_PDF
MAX_CONTENU = 5 * 1024 * 1024  # 5 Mo max lus par fichier pour les regles "contenu"

DOSSIERS_TRI = {"_a_mettre_de_cote", "_doublons_a_trier", "_dossiers_en_double", "_rapport_scan", "_tri_dedup_scan"}

# Donnees CHU/patients : ne jamais monter vers un cloud
REGLE_CHU = re.compile(r"(?<![a-z0-9])(chu|chucaen|chu de caen|patients?|dossier medical|donnees patients|hospitalier)(?![a-z0-9])")

DEST_RACINES = {"ONEDRIVE": "onedrive", "PROTON": "proton", "LOCAL_CHU": None, "REVUE": None}


def normaliser(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower()


def est_dans_cloud(chemin):
    n = normaliser(chemin)
    return ("cloudstorage" in n or "onedrive" in n or "proton" in n
            or "icloud" in n or "google drive" in n or n.startswith("/volumes/"))


def detecter_racine(prefix):
    cands = sorted(glob.glob(os.path.expanduser(f"~/Library/CloudStorage/{prefix}*"))) + \
            sorted(glob.glob(os.path.expanduser(f"~/{prefix}*")))
    return cands[0] if cands else None


def charger_regles(path):
    regles = []
    with open(path, encoding="utf-8-sig") as f:
        for ligne in f:
            ligne = ligne.strip()
            if not ligne or ligne.startswith("#"):
                continue
            p = ligne.split("\t")
            if len(p) >= 4:
                regles.append((p[0].strip(), p[1].strip().lower(), p[2].strip(), p[3].strip().upper()))
    return regles


def motif_regex(motif):
    # mot entier, insensible casse/accents, pluriel en s toléré :
    # "cours" ne matche pas "parcours", "impot" matche "impots"
    return re.compile(r"(?<![a-z0-9])" + re.escape(normaliser(motif)) + r"s?(?![a-z0-9])")


def classer(chemin, racine, regles, defaut):
    segments = {normaliser(s) for s in chemin.split(os.sep)}
    if segments & DOSSIERS_TRI:
        return "DEJA_TRIE", "-", "RESTE_EN_PLACE"
    n = normaliser(chemin)
    if REGLE_CHU.search(n):
        if est_dans_cloud(chemin):
            return ("CHU", "securite CHU", "A_TRAITER_MANUELLEMENT (donnees CHU/patients dans un cloud : a remettre sur le PC CHU, jamais vers Proton/OneDrive perso)")
        return "CHU", "securite CHU", "RESTE_EN_PLACE (local)"
    nom = normaliser(os.path.basename(chemin))
    dossiers = [normaliser(s) for s in chemin.split(os.sep)[:-1]]
    texte = None
    for cat, type_, motif, dest in regles:
        if type_ == "contenu":
            if texte is None:
                ext = os.path.splitext(chemin)[1].lower()
                try:
                    taille = os.path.getsize(chemin)
                except OSError:
                    taille = 0
                if ext in EXT_TEXTE_CLASSEMENT and 0 < taille <= MAX_CONTENU:
                    texte = normaliser(extraire_texte(chemin, ext, MAX_CONTENU) or "")
                else:
                    texte = ""
            if texte and motif_regex(motif).search(texte):
                return cat, "contenu:" + motif, dest
            continue
        rx = motif_regex(motif)
        if type_ == "chemin" and rx.search(n):
            return cat, motif, dest
        if type_ == "dossier" and any(rx.search(d) for d in dossiers):
            return cat, "dossier:" + motif, dest
        if type_ == "nom" and rx.search(nom):
            return cat, motif, dest
        if type_ == "ext" and chemin.lower().endswith(normaliser(motif)):
            return cat, motif, dest
    return "PAR_DEFAUT", "-", defaut


def chemin_unique(dest):
    if not os.path.exists(dest):
        return dest
    base, ext = os.path.splitext(dest)
    i = 1
    while os.path.exists(f"{base}__{i}{ext}"):
        i += 1
    return f"{base}__{i}{ext}"


def main():
    ap = argparse.ArgumentParser(description="Planifie/applique le classement univ->OneDrive, perso->Proton, CHU->local. Dry-run par defaut, aucune suppression.")
    ap.add_argument("--report-dir", default=os.path.expanduser("~/Downloads/rapports_scan"))
    ap.add_argument("--inventaire", default=None, help="inventaire TSV (defaut: le plus recent du report-dir)")
    ap.add_argument("--regles", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "regles_classement.tsv"))
    ap.add_argument("--onedrive", default=None, help="racine OneDrive cible (auto-detecte sinon)")
    ap.add_argument("--proton", default=None, help="racine Proton Drive cible (auto-detecte sinon)")
    ap.add_argument("--default", choices=["REVUE", "PROTON", "ONEDRIVE"], default="REVUE",
                    help="destination si aucune regle ne matche (defaut REVUE = signale seulement)")
    ap.add_argument("--protect", action="append", default=[], help="sous-chaine : chemin en lecture seule (plan uniquement)")
    ap.add_argument("--apply", action="store_true", help="applique les deplacements (sinon dry-run)")
    args = ap.parse_args()

    inv = args.inventaire
    if not inv:
        cands = sorted(glob.glob(os.path.join(args.report_dir, "inventaire_*.tsv")))
        if not cands:
            print("Aucun inventaire trouve dans", args.report_dir, file=sys.stderr)
            return 1
        inv = cands[-1]
    print(f"[*] Inventaire : {inv}")

    regles = charger_regles(args.regles)
    print(f"[*] {len(regles)} regles chargees depuis {args.regles}")

    onedrive = args.onedrive or detecter_racine("OneDrive")
    proton = args.proton or detecter_racine("ProtonDrive") or detecter_racine("Proton")
    racines_dest = {"ONEDRIVE": onedrive, "PROTON": proton}
    print(f"[*] OneDrive cible : {onedrive}")
    print(f"[*] Proton cible   : {proton}")

    ts = time.strftime("%Y%m%d_%H%M%S")
    plan_path = os.path.join(args.report_dir, f"plan_classement_{ts}.csv")
    lignes = []
    with open(inv, encoding="utf-8-sig") as f:
        next(f, None)
        for ligne in f:
            p = ligne.rstrip("\n").split("\t")
            if len(p) < 2:
                continue
            racine, chemin = p[0], p[1]
            if not os.path.exists(chemin):
                continue
            cat, regle, dest = classer(chemin, racine, regles, args.default)
            if dest.startswith(("A_TRAITER", "RESTE")):
                action_txt = dest
                dest_path = ""
            elif dest in ("LOCAL_CHU", "REVUE"):
                action_txt = "RESTE_EN_PLACE"
                dest_path = ""
            else:
                cible = racines_dest.get(dest)
                if not cible:
                    action_txt = "RESTE_EN_PLACE"
                    dest_path = ""
                elif os.path.abspath(chemin).startswith(os.path.abspath(cible) + os.sep):
                    action_txt = "RESTE_EN_PLACE (deja dans la cible)"
                    dest_path = ""
                else:
                    rel = os.path.relpath(chemin, racine)
                    dest_path = chemin_unique(os.path.join(cible, rel))
                    action_txt = "DEPLACER"
            if any(normaliser(p_) in normaliser(chemin) for p_ in args.protect):
                action_txt = "PLAN_SEULEMENT (protege)"
            lignes.append([chemin, racine, cat, regle, action_txt, dest_path])

    with open(plan_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["chemin", "racine", "categorie", "regle", "action", "destination", "resultat"])
        for l in lignes:
            w.writerow(l + [""])

    from collections import Counter
    compte = Counter(l[4] for l in lignes)
    print(f"[*] Plan : {len(lignes)} fichiers -> {plan_path}")
    for k, v in compte.most_common():
        print(f"    {v:>7}  {k}")

    if args.apply:
        print("[*] Application des deplacements...")
        ok = echec = 0
        for l in lignes:
            if l[4] != "DEPLACER":
                continue
            try:
                os.makedirs(os.path.dirname(l[5]), exist_ok=True)
                shutil.move(l[0], l[5])
                l.append("OK")
                ok += 1
            except OSError as e:
                l.append(f"ECHEC: {e}")
                echec += 1
        with open(plan_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["chemin", "racine", "categorie", "regle", "action", "destination", "resultat"])
            for l in lignes:
                w.writerow(l)
        print(f"[*] Deplacements : {ok} OK, {echec} ECHEC")
        hist = os.path.join(args.report_dir, "historique_classement.tsv")
        nouveau = not os.path.exists(hist)
        with open(hist, "a", encoding="utf-8") as f:
            if nouveau:
                f.write("ts\tinventaire\tfichiers\tdeplacer_ok\tdeplacer_echec\ta_traiter_manuellement\tprotege\n")
            f.write(f"{ts}\t{inv}\t{len(lignes)}\t{ok}\t{echec}\t"
                    f"{sum(1 for l in lignes if l[4].startswith('A_TRAITER'))}\t"
                    f"{sum(1 for l in lignes if l[4].startswith('PLAN_SEULEMENT'))}\n")
    else:
        print("[*] Dry-run : aucun deplacement effectue. Relance avec --apply pour appliquer.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
