#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dedup_scan.py - Scan local de doublons, versions et fichiers a mettre de cote.

100 % local, bibliotheque standard uniquement. Aucun envoi reseau, aucune suppression.

Usage :
    python3 dedup_scan.py ROOT [ROOT ...] [options]
    python3 dedup_scan.py --auto [options]   # scan total : tous les disques / home + /Volumes

Options :
    --protect SUBSTR     racine "protegee" si son chemin contient SUBSTR : rapport seul, aucun deplacement
    --dry-run            simule les actions, ne touche a rien
    --report-dir DIR     dossier des rapports (defaut : <premiere racine>/_RAPPORT_SCAN)
    --max-content N      taille max (octets) lue pour la recherche dans le contenu (defaut 20 Mo)

Actions : fichiers tagues -> _A_METTRE_DE_COTE (deplacement, jamais de suppression),
dossiers entierement dupliques -> _DOSSIERS_EN_DOUBLE, doublons fichier -> _DOUBLONS_A_TRIER (copie).
L'exemplaire conserve de chaque groupe est marque "CONSERVE EN PLACE" dans le rapport.

Sorties dans le dossier de rapport :
- inventaire_<ts>.tsv : ecrit en continu pendant le scan (resiste a un crash)
- checkpoint.tsv      : hash/taille/date des fichiers deja vus ; une relance les reutilise
                        (fichier inchange = pas de rehash). Supprime-le pour tout recalculer.
- historique_runs.tsv : une ligne par execution (le "chapitrage" des scans successifs)
"""

import argparse
import csv
import fnmatch
import hashlib
import html
import os
import re
import shutil
import sys
import time
import unicodedata
import zipfile
import zlib
from datetime import datetime

NOMS_EX = ["elsa", "eric", "chantal", "laetitia", "estelle", "josy", "michelle"]
DOSSIERS_IGNORES = {"_A_METTRE_DE_COTE", "_DOUBLONS_A_TRIER", "_DOSSIERS_EN_DOUBLE", "_RAPPORT_SCAN"}
DOSSIERS_SYSTEME = {
    "Windows", "$Recycle.Bin", "System Volume Information", "Program Files",
    "Program Files (x86)", "ProgramData", "MSOCache", "PerfLogs",
    "node_modules", ".git", "__pycache__", ".Trash", "Caches",
}
FICHIERS_SYSTEME = {"pagefile.sys", "hiberfil.sys", "swapfile.sys", "dumpfile.sys", "Thumbs.db", ".DS_Store"}
NOM_RAPPORT_PREFIX = "rapport_"

EXT_TEXT = {
    ".txt", ".md", ".csv", ".tsv", ".log", ".eml", ".html", ".htm", ".xml",
    ".json", ".ini", ".cfg", ".py", ".r", ".tex", ".srt",
}
EXT_ZIP_XML = {".docx", ".pptx", ".xlsx", ".odt", ".ods", ".epub"}
EXT_PDF = {".pdf"}

DEFAULT_MAX_CONTENT = 20 * 1024 * 1024


def normaliser(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower()


NOMS_NORM = [normaliser(n) for n in NOMS_EX]
PATTERN_NOMS = re.compile(
    r"(?<![a-z0-9])(" + "|".join(sorted(NOMS_NORM, key=len, reverse=True)) + r")(?![a-z0-9])"
)

PATTERN_VARIANTE = re.compile(
    r"[\s_\-.]*(v\s?\d+|final(?:e)?|def(?:initif(?:ve)?)?|copie|copy|bis|ter|quater|"
    r"\(\s*\d+\s*\)|-\d+|\d{4}-\d{2}-\d{2}|\d{8}|new|old|brouillon|draft)"
    r"[\s_\-.]*$"
)


def nettoyer_tige(tige):
    s = normaliser(tige)
    precedent = None
    while precedent != s:
        precedent = s
        s = PATTERN_VARIANTE.sub("", s)
    return s.strip(" _-.")


def sha256_fichier(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            bloc = f.read(1024 * 1024)
            if not bloc:
                break
            h.update(bloc)
    return h.hexdigest()


def extraire_texte(path, ext, max_content):
    try:
        taille = os.path.getsize(path)
        if taille > max_content:
            return None
        if ext in EXT_TEXT:
            with open(path, "rb") as f:
                return f.read(max_content).decode("utf-8", "ignore")
        if ext in EXT_ZIP_XML:
            morceaux = []
            with zipfile.ZipFile(path) as z:
                total = 0
                for info in z.infolist():
                    if not info.filename.endswith(".xml"):
                        continue
                    if total > max_content:
                        break
                    brut = z.read(info)[: max_content - total]
                    total += len(brut)
                    morceaux.append(re.sub(r"<[^>]+>", " ", brut.decode("utf-8", "ignore")))
            return " ".join(morceaux)
        if ext in EXT_PDF:
            with open(path, "rb") as f:
                data = f.read(max_content)
            morceaux = [data.decode("latin-1", "ignore")]
            for m in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.S):
                bloc = m.group(1).strip(b"\r\n")
                for candidat in (bloc, bloc[2:] if len(bloc) > 2 else bloc):
                    try:
                        morceaux.append(zlib.decompress(candidat).decode("latin-1", "ignore"))
                        break
                    except Exception:
                        continue
            return " ".join(morceaux)
    except Exception:
        return None
    return None


def decouvrir_racines_auto():
    racines = []
    if os.name == "nt":
        import string
        for lettre in string.ascii_uppercase:
            d = f"{lettre}:\\"
            if os.path.isdir(d):
                racines.append(d)
    else:
        racines.append(os.path.expanduser("~"))
        volumes = "/Volumes"
        if os.path.isdir(volumes):
            for v in sorted(os.listdir(volumes)):
                if v.startswith(".") or "Macintosh HD" in v:
                    continue
                p = os.path.join(volumes, v)
                if os.path.isdir(p):
                    racines.append(p)
    return racines


def charger_checkpoint(path):
    ckpt = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig", errors="ignore") as f:
            for ligne in f:
                parties = ligne.rstrip("\n").split("\t", 5)
                if len(parties) == 6:
                    sha, taille, mtime, tn, tc, chemin = parties
                    try:
                        ckpt[chemin] = {
                            "taille": int(taille), "mtime": int(float(mtime)), "sha256": sha,
                            "touche_nom": tn == "1", "touche_contenu": tc == "1",
                        }
                    except ValueError:
                        continue
    return ckpt


def scanner_racine(racine, max_content, ckpt=None, inventaire=None, ckpt_fichier=None):
    enregistrements = []
    erreurs = []
    ckpt = ckpt if ckpt is not None else {}
    tampon_inv = []
    tampon_ck = []
    racine = os.path.abspath(racine)
    for dirpath, dirnames, filenames in os.walk(racine, followlinks=False):
        dirnames[:] = [d for d in dirnames if d not in DOSSIERS_IGNORES and d not in DOSSIERS_SYSTEME]
        for nom in sorted(filenames):
            if nom.startswith(NOM_RAPPORT_PREFIX) or nom in FICHIERS_SYSTEME or fnmatch.fnmatch(nom.lower(), "*.log.tmp"):
                continue
            chemin = os.path.join(dirpath, nom)
            try:
                st = os.stat(chemin)
                if not os.path.isfile(chemin):
                    continue
                ext = os.path.splitext(nom)[1].lower()
                entree = ckpt.get(chemin)
                if entree and entree["taille"] == st.st_size and entree["mtime"] == int(st.st_mtime):
                    h = entree["sha256"]
                    tn = entree["touche_nom"]
                    tc = entree["touche_contenu"]
                else:
                    h = sha256_fichier(chemin)
                    texte = extraire_texte(chemin, ext, max_content) if ext in (EXT_TEXT | EXT_ZIP_XML | EXT_PDF) else None
                    tn = bool(PATTERN_NOMS.search(normaliser(nom)))
                    tc = bool(texte and PATTERN_NOMS.search(normaliser(texte)))
                    tampon_ck.append(f'{h}\t{st.st_size}\t{int(st.st_mtime)}\t{1 if tn else 0}\t{1 if tc else 0}\t{chemin}')
                    ckpt[chemin] = {"taille": st.st_size, "mtime": int(st.st_mtime), "sha256": h, "touche_nom": tn, "touche_contenu": tc}
                enregistrements.append({
                    "racine": racine,
                    "chemin": chemin,
                    "relatif": os.path.relpath(chemin, racine),
                    "nom": nom,
                    "ext": ext,
                    "taille": st.st_size,
                    "mtime": st.st_mtime,
                    "sha256": h,
                    "touche_nom": tn,
                    "touche_contenu": tc,
                })
                if inventaire is not None:
                    tampon_inv.append(f'{racine}\t{chemin}\t{st.st_size}\t{int(st.st_mtime)}\t{h}\t{1 if tn else 0}\t{1 if tc else 0}')
                if inventaire is not None and len(tampon_inv) >= 1000:
                    inventaire.write("\n".join(tampon_inv) + "\n")
                    inventaire.flush()
                    tampon_inv = []
                if ckpt_fichier is not None and len(tampon_ck) >= 1000:
                    ckpt_fichier.write("\n".join(tampon_ck) + "\n")
                    ckpt_fichier.flush()
                    tampon_ck = []
            except OSError as e:
                erreurs.append((chemin, str(e)))
            if len(enregistrements) % 5000 == 0:
                print(f"    ... {len(enregistrements)} fichiers analyses")
    if inventaire is not None and tampon_inv:
        inventaire.write("\n".join(tampon_inv) + "\n")
        inventaire.flush()
    if ckpt_fichier is not None and tampon_ck:
        ckpt_fichier.write("\n".join(tampon_ck) + "\n")
        ckpt_fichier.flush()
    return enregistrements, erreurs


def groupe_par(records, cle):
    groupes = {}
    for r in records:
        groupes.setdefault(cle(r), []).append(r)
    return groupes


def chemin_unique(dest):
    if not os.path.exists(dest):
        return dest
    base, ext = os.path.splitext(dest)
    i = 1
    while os.path.exists(f"{base}__{i}{ext}"):
        i += 1
    return f"{base}__{i}{ext}"


def bases_action(racine, dossier_cible, dry_run):
    """Dossier d'action : racine si inscriptible, sinon repli ~/_TRI_DEDUP_SCAN (ex. racine cloud File Provider)."""
    candidats = [os.path.join(racine, dossier_cible)]
    if not dry_run:
        candidats.append(os.path.join(os.path.expanduser("~"), "_TRI_DEDUP_SCAN", dossier_cible))
    return candidats


def action_fs(racine, dossier_cible, relatif, dry_run, operation, src):
    for base in bases_action(racine, dossier_cible, dry_run):
        dest = chemin_unique(os.path.join(base, relatif))
        if dry_run:
            return dest, None
        try:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            operation(src, dest)
            return dest, None
        except OSError:
            continue
    return None, "ecriture refusee dans la racine et dans ~/_TRI_DEDUP_SCAN"


def deplacer_vers(racine, r, dossier_cible, actions, dry_run, raison):
    dest, err = action_fs(racine, dossier_cible, r["relatif"], dry_run, shutil.move, r["chemin"])
    actions.append({
        "raison": raison, "source": r["chemin"], "destination": dest or "",
        "sha256": r["sha256"], "taille": r["taille"], "simulation": dry_run,
        "statut": "ECHEC" if err else "OK", "erreur": err or "",
    })
    return dest


def copier_vers(racine, r, dossier_cible, actions, dry_run, raison):
    dest, err = action_fs(racine, dossier_cible, r["relatif"], dry_run, shutil.copy2, r["chemin"])
    actions.append({
        "raison": raison, "source": r["chemin"], "destination": dest or "",
        "sha256": r["sha256"], "taille": r["taille"], "simulation": dry_run,
        "statut": "ECHEC" if err else "OK", "erreur": err or "",
    })
    return dest


def chemin_protege(chemin, protects):
    n = normaliser(chemin)
    return any(normaliser(p) in n for p in protects)


def lire_manifeste(racine):
    p = os.path.join(racine, "_DOUBLONS_A_TRIER", ".deja_copies.txt")
    if os.path.exists(p):
        with open(p, encoding="utf-8", errors="ignore") as f:
            return set(f.read().splitlines())
    return set()


def ajouter_manifeste(racine, cle):
    p = os.path.join(racine, "_DOUBLONS_A_TRIER", ".deja_copies.txt")
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "a", encoding="utf-8") as f:
        f.write(cle + "\n")


def produire_rapport(repertoire, ts, sections, actions, erreurs, racines, dry_run, protects, conserves=None):
    conserves = conserves or set()
    actions_par_source = {}
    for a in actions:
        actions_par_source.setdefault(a["source"], a)
    os.makedirs(repertoire, exist_ok=True)
    csv_path = os.path.join(repertoire, f"rapport_{ts}.csv")
    html_path = os.path.join(repertoire, f"rapport_{ts}.html")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["categorie", "groupe", "chemin", "taille_octets", "modifie_le", "sha256", "action", "destination"])
        for categorie, groupes in sections.items():
            for gname, fichiers in groupes:
                for r in fichiers:
                    act = actions_par_source.get(r["chemin"])
                    if act:
                        action_txt = act["raison"] + (" (SIMULATION)" if act["simulation"] else "")
                        if act.get("statut") == "ECHEC":
                            action_txt = "ECHEC: " + act.get("erreur", "") + " | " + action_txt
                        dest_txt = act["destination"]
                    elif r["chemin"] in conserves:
                        action_txt = "CONSERVÉ EN PLACE (exemplaire unique gardé)"
                        dest_txt = ""
                    else:
                        action_txt = ""
                        dest_txt = ""
                    w.writerow([
                        categorie, gname, r["chemin"], r["taille"],
                        datetime.fromtimestamp(r["mtime"]).isoformat(timespec="seconds"),
                        r["sha256"], action_txt, dest_txt,
                    ])
        for e in erreurs:
            w.writerow(["ERREUR", "", e[0], "", "", "", e[1], ""])

    def ligne(r):
        return ("<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>").format(
            html.escape(r["chemin"]), r["taille"],
            datetime.fromtimestamp(r["mtime"]).strftime("%Y-%m-%d %H:%M:%S"),
            html.escape(r["sha256"][:16]) + "…")

    parties = [
        "<html><head><meta charset='utf-8'><title>Rapport scan " + ts + "</title>",
        "<style>body{font-family:sans-serif;margin:2em} h1,h2{color:#234} table{border-collapse:collapse;margin-bottom:2em} "
        "td,th{border:1px solid #bbb;padding:4px 8px;font-size:13px} th{background:#e8eef7}</style></head><body>",
        f"<h1>Rapport de scan du {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</h1>",
        f"<p>Racines scannees : {html.escape(', '.join(racines))}<br>"
        f"Mode : {'SIMULATION (dry-run)' if dry_run else 'actions appliquees'} - "
        f"protections : {html.escape(', '.join(protects)) if protects else 'aucune'} - "
        "aucun fichier n'est supprime.</p>",
    ]
    for categorie, groupes in sections.items():
        parties.append(f"<h2>{html.escape(categorie)} ({len(groupes)} groupe(s))</h2>")
        for gname, fichiers in groupes:
            parties.append(f"<h3>{html.escape(gname)}</h3><table>")
            parties.append("<tr><th>Chemin</th><th>Taille</th><th>Modifie le</th><th>SHA-256</th></tr>")
            for r in fichiers:
                parties.append(ligne(r))
            parties.append("</table>")
    if actions:
        parties.append(f"<h2>Actions effectuees ({len(actions)})</h2><table>")
        parties.append("<tr><th>Raison</th><th>Source</th><th>Destination</th></tr>")
        for a in actions:
            parties.append("<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
                html.escape(a["raison"] + (" (SIMULATION)" if a["simulation"] else "")),
                html.escape(a["source"]), html.escape(a["destination"])))
        parties.append("</table>")
    if erreurs:
        parties.append("<h2>Erreurs de lecture</h2><ul>")
        for chemin, msg in erreurs:
            parties.append(f"<li>{html.escape(chemin)} : {html.escape(msg)}</li>")
        parties.append("</ul>")
    parties.append("</body></html>")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parties))
    return csv_path, html_path


def main():
    ap = argparse.ArgumentParser(description="Scan local : doublons, versions, fichiers a mettre de cote. 100% local, rien n'est supprime.")
    ap.add_argument("racines", nargs="*", help="Arborescences a scanner (ex. /Volumes/MonDisque)")
    ap.add_argument("--auto", action="store_true", help="Scan total automatique : tous les disques (Windows) ou home + /Volumes (Mac/Linux)")
    ap.add_argument("--protect", action="append", default=[], help="Sous-chaine : racine en lecture seule (rapport seul)")
    ap.add_argument("--dry-run", action="store_true", help="Simule sans rien toucher")
    ap.add_argument("--report-dir", default=None, help="Dossier des rapports (defaut : <racine1>/_RAPPORT_SCAN)")
    ap.add_argument("--max-content", type=int, default=DEFAULT_MAX_CONTENT, help="Taille max lue pour le contenu (octets)")
    args = ap.parse_args()

    if args.auto:
        auto = decouvrir_racines_auto()
        print("[*] Mode --auto, racines decouvertes :")
        for r in auto:
            print("    " + r)
        args.racines = list(args.racines) + auto
    if not args.racines:
        ap.error("Indique au moins une racine (ex. ~/Documents) ou --auto")

    ts = time.strftime("%Y%m%d_%H%M%S")
    report_dir = args.report_dir or os.path.join(os.path.abspath(args.racines[0]), "_RAPPORT_SCAN")
    os.makedirs(report_dir, exist_ok=True)
    ckpt_path = os.path.join(report_dir, "checkpoint.tsv")
    ckpt = charger_checkpoint(ckpt_path)
    print(f"[*] Checkpoint : {len(ckpt)} fichiers deja connus (hash reutilise si taille+date identiques)")
    inv_path = os.path.join(report_dir, f"inventaire_{ts}.tsv")
    inventaire = open(inv_path, "a", encoding="utf-8")
    if inventaire.tell() == 0:
        inventaire.write("racine\tchemin\ttaille\tmtime\tsha256\ttouche_nom\ttouche_contenu\n")
    ckpt_fichier = open(ckpt_path, "a", encoding="utf-8")
    tous = []
    toutes_erreurs = []
    for racine in args.racines:
        if not os.path.isdir(racine):
            print(f"[!] Racine ignoree (introuvable) : {racine}", file=sys.stderr)
            continue
        print(f"[*] Scan de {racine} ...")
        recs, errs = scanner_racine(racine, args.max_content, ckpt, inventaire, ckpt_fichier)
        tous.extend(recs)
        toutes_erreurs.extend(errs)
        print(f"    {len(recs)} fichiers, {len(errs)} erreur(s)")
    inventaire.close()
    ckpt_fichier.close()

    if not tous:
        print("Aucun fichier trouve.")
        return 1

    print("[*] Analyse : regroupement par hash, nom, variantes, dossiers dupliques...")
    sections = {}

    par_hash = groupe_par(tous, lambda r: r["sha256"])
    doublons = [(h, rs) for h, rs in par_hash.items() if len(rs) > 1]
    doublons.sort(key=lambda g: -sum(r["taille"] for r in g[1]))
    sections["DOUBLONS EXACTS (meme contenu, hash identique)"] = [
        (f"{h[:12]}… ({len(rs)} fichiers)", sorted(rs, key=lambda r: r["chemin"])) for h, rs in doublons
    ]

    par_nom = groupe_par(tous, lambda r: normaliser(r["nom"]))
    memes_noms = [
        (nom, rs) for nom, rs in par_nom.items()
        if len(rs) > 1 and len({r["sha256"] for r in rs}) > 1
    ]
    memes_noms.sort(key=lambda g: g[0])
    sections["MEMES NOMS, CONTENUS DIFFERENTS (versions)"] = [
        (nom, sorted(rs, key=lambda r: r["chemin"])) for nom, rs in memes_noms
    ]

    par_tige = groupe_par(tous, lambda r: (nettoyer_tige(os.path.splitext(r["nom"])[0]), r["ext"]))
    variantes = [
        (cle, rs) for cle, rs in par_tige.items()
        if len(rs) > 1 and len({r["sha256"] for r in rs}) > 1
    ]
    variantes.sort(key=lambda g: g[0])
    sections["VARIANTES DE NOM (v2, final, copie, (1), dates...)"] = [
        (f"{cle[0]}{cle[1]} ({len(rs)} fichiers)", sorted(rs, key=lambda r: r["chemin"])) for cle, rs in variantes
    ]

    touches = [r for r in tous if r["touche_nom"] or r["touche_contenu"]]
    touches.sort(key=lambda r: r["chemin"])
    sections["FICHIERS A METTRE DE COTE (noms : " + ", ".join(NOMS_EX) + ")"] = [
        (f"{'nom' if r['touche_nom'] else ''}{'+' if r['touche_nom'] and r['touche_contenu'] else ''}{'contenu' if r['touche_contenu'] else ''} : {r['nom']}", [r]) for r in touches
    ]

    def est_sous(chemin, racine):
        return chemin == racine or chemin.startswith(racine + os.sep)

    def est_couvert(chemin, deplaces_set):
        p = chemin
        while True:
            if p in deplaces_set:
                return True
            parent = os.path.dirname(p)
            if parent == p:
                return False
            p = parent

    par_dossier = {}
    for r in tous:
        par_dossier.setdefault(os.path.dirname(r["chemin"]), []).append(r)

    def sig_dossier(d, recs):
        lignes = sorted(f'{os.path.relpath(r["chemin"], d)}|{r["sha256"]}' for r in recs)
        h = hashlib.sha256()
        for l in lignes:
            h.update(l.encode("utf-8"))
        return h.hexdigest()

    sigs = {}
    for d, recs in par_dossier.items():
        sigs.setdefault(sig_dossier(d, recs), []).append((d, recs))

    dossiers_doublons = []
    deplaces = []
    deplaces_set = set()
    for sig, lst in sorted(sigs.items(), key=lambda kv: min(len(d) for d, _ in kv[1])):
        lst = [
            (d, recs) for d, recs in lst
            if not any(r["touche_nom"] or r["touche_contenu"] for r in recs)
            and not est_couvert(d, deplaces_set)
        ]
        if len(lst) < 2:
            continue
        lst.sort(key=lambda item: (-max(r["mtime"] for r in item[1]), item[0]))
        keeper, rest = lst[0], lst[1:]
        dossiers_doublons.append((sig, keeper, rest))
        for d, _ in rest:
            deplaces.append(d)
            deplaces_set.add(d)

    def rec_dossier(d, recs, sig):
        return {
            "chemin": d,
            "taille": sum(r["taille"] for r in recs),
            "mtime": max(r["mtime"] for r in recs),
            "sha256": sig,
        }

    sections["DOSSIERS ENTIEREMENT DUPLIQUES (contenu identique ailleurs)"] = [
        (
            f"exemplaire conservé : {keeper[0]}",
            [rec_dossier(keeper[0], keeper[1], sig)] + [rec_dossier(d, recs, sig) for d, recs in rest],
        )
        for sig, keeper, rest in dossiers_doublons
    ]

    conserves = {sorted(rs, key=lambda r: (-r["mtime"], r["chemin"]))[0]["chemin"] for _, rs in doublons}
    conserves.update(keeper[0] for _, keeper, _ in dossiers_doublons)
    deplaces_fichiers = deplaces_set

    manifestes = {os.path.abspath(racine): lire_manifeste(os.path.abspath(racine)) for racine in args.racines}

    print("[*] Actions : " + ("simulation (dry-run)" if args.dry_run else "application reelle") + "...")
    actions = []
    for r in touches:
        if chemin_protege(r["chemin"], args.protect):
            continue
        deplacer_vers(r["racine"], r, "_A_METTRE_DE_COTE", actions, args.dry_run, "nom/contenu ex-partenaire")

    for sig, keeper, rest in dossiers_doublons:
        for d, recs in rest:
            if chemin_protege(d, args.protect):
                continue
            racine = recs[0]["racine"]
            dest, err = action_fs(racine, "_DOSSIERS_EN_DOUBLE", os.path.basename(d), args.dry_run, shutil.move, d)
            actions.append({
                "raison": f"dossier entierement doublon de {keeper[0]}",
                "source": d, "destination": dest or "",
                "sha256": sig, "taille": sum(r["taille"] for r in recs),
                "simulation": args.dry_run,
                "statut": "ECHEC" if err else "OK", "erreur": err or "",
            })

    partis_en_quarantaine = {r["chemin"] for r in touches}
    for h, rs in doublons:
        rs_tries = sorted(rs, key=lambda r: (-r["mtime"], r["chemin"]))
        gardien = rs_tries[0]
        for r in rs_tries[1:]:
            if chemin_protege(r["chemin"], args.protect):
                continue
            if r["chemin"] in partis_en_quarantaine or not os.path.exists(r["chemin"]):
                continue
            if est_couvert(r["chemin"], deplaces_fichiers):
                continue
            cle = f'{r["sha256"]}\t{r["chemin"]}'
            if cle in manifestes.get(r["racine"], set()):
                continue
            copier_vers(r["racine"], r, "_DOUBLONS_A_TRIER", actions, args.dry_run, f"doublon de {gardien['chemin']}")
            if not args.dry_run:
                manifestes.setdefault(r["racine"], set()).add(cle)
                ajouter_manifeste(r["racine"], cle)

    print("[*] Ecriture du rapport...")
    csv_path, html_path = produire_rapport(
        report_dir, ts, sections, actions, toutes_erreurs,
        [os.path.abspath(r) for r in args.racines], args.dry_run, args.protect, conserves,
    )

    print()
    print(f"Fichiers scannes        : {len(tous)}")
    print(f"Groupes de doublons     : {len(doublons)} ({sum(len(rs)-1 for _, rs in doublons)} copie(s) prevue(s))")
    print(f"Dossiers tout doublons  : {len(dossiers_doublons)} ({len(deplaces)} dossier(s) a deplacer)")
    print(f"Groupes memes noms      : {len(memes_noms)}")
    print(f"Groupes variantes       : {len(variantes)}")
    print(f"Fichiers a mettre de cote : {len(touches)}")
    print(f"Actions                 : {len(actions)} {'(SIMULATION)' if args.dry_run else ''} dont {sum(1 for a in actions if a.get('statut') == 'ECHEC')} en ECHEC")
    print(f"Erreurs de lecture      : {len(toutes_erreurs)}")
    hist_path = os.path.join(report_dir, "historique_runs.tsv")
    nouveau_hist = not os.path.exists(hist_path) or os.path.getsize(hist_path) == 0
    with open(hist_path, "a", encoding="utf-8") as f:
        if nouveau_hist:
            f.write("ts\tracines\tfichiers\tgroupes_doublons\tdossiers_doublons\tmemes_noms\tvariantes\tfichiers_a_cote\tactions\terreurs\tdry_run\n")
        f.write(f"{ts}\t{';'.join(os.path.abspath(r) for r in args.racines)}\t{len(tous)}\t{len(doublons)}\t{len(dossiers_doublons)}\t{len(memes_noms)}\t{len(variantes)}\t{len(touches)}\t{len(actions)}\t{len(toutes_erreurs)}\t{args.dry_run}\n")
    print(f"Rapport CSV             : {csv_path}")
    print(f"Rapport HTML            : {html_path}")
    print(f"Inventaire (continu)    : {inv_path}")
    print(f"Checkpoint (reprise)    : {ckpt_path}")
    print(f"Historique des runs     : {hist_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
