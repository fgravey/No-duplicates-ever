# dedup-scan — scanner local de doublons, versions et fichiers « ex-partenaire »

Outil 100 % **local** : aucune installation, aucun envoi vers un cloud, **aucun fichier supprimé**.
Deux moteurs équivalents :

| Fichier | Où l'utiliser |
|---|---|
| `dedup_scan.ps1` | PC CHU (Windows), en PowerShell natif — zéro installation, compatible DSI |
| `dedup_scan.py` | Mac / Linux — `python3` (bibliothèque standard seule, rien à installer) |

## Ce que le scan produit

Rapport **CSV** (s'ouvre dans Excel, accents OK) + **HTML**, écrits dans `_RAPPORT_SCAN/` :

1. **Doublons exacts** — comparaison par hash SHA-256 (pas seulement nom/taille).
   Les doublons (sauf un exemplaire conservé, le plus récent) sont **copiés** dans `_DOUBLONS_A_TRIER/`.
   Dans le rapport CSV, l'exemplaire qui **reste en place** est marqué `CONSERVE EN PLACE`,
   et chaque copie indique `doublon de <chemin de l'exemplaire conservé>` : tu sais donc
   toujours dans quel dossier se trouve la copie unique.
2. **Dossiers entièrement dupliqués** — si tout le contenu d'un dossier (mêmes chemins
   relatifs, mêmes hashes) existe à l'identique ailleurs, le dossier entier est **déplacé**
   dans `_DOSSIERS_EN_DOUBLE/` (jamais supprimé). Exception : un dossier contenant des
   fichiers tagués « ex-partenaire » n'est pas déplacé (ses fichiers partent en quarantaine
   individuellement et le dossier est signalé dans le rapport).
2. **Fichiers de même nom, contenus différents** (ex. `rapport.docx` à plusieurs endroits).
3. **Variantes de nom** : `rapport_v2.docx`, `rapport final.docx`, `rapport (1).docx`, dates…
4. **Fichiers à mettre de côté** : nom de fichier **ou** contenu contenant, en mot entier
   (insensible casse/accents, donc « America » ne déclenche pas « eric ») :
   `elsa, eric, chantal, laetitia, estelle, josy, michelle`.
   Recherche dans les noms **et** dans le texte des `.txt .md .csv .eml .html .docx .pptx .xlsx .odt .pdf`.
   Ces fichiers sont **déplacés sans suppression** dans `_A_METTRE_DE_COTE/` (chemin relatif conservé).

## Utilisation

PC CHU (PowerShell, lancer depuis le dossier du script) — **d'abord en simulation** :

```powershell
powershell -ExecutionPolicy Bypass -File .\dedup_scan.ps1 -Roots "D:\MesDocuments" -DryRun
powershell -ExecutionPolicy Bypass -File .\dedup_scan.ps1 -Roots "D:\MesDocuments","C:\Users\MOI\OneDrive","C:\Users\MOI\Proton Drive","E:\DisqueExterne" -Protect "CHU"
```

Mac :

```bash
python3 dedup_scan.py ~/Documents ~/OneDrive ~/ProtonDrive --dry-run
python3 dedup_scan.py ~/Documents ~/OneDrive ~/ProtonDrive --protect CHU
# disques externes / clouds visibles dans le Finder (points de montage /Volumes) :
python3 dedup_scan.py /Volumes/MonDisqueUSB "/Volumes/Proton Drive" ~/Library/CloudStorage/OneDrive-Personnel
```

Options communes : `-DryRun` / `--dry-run` (simule, ne touche à rien),
`-Protect "CHU"` / `--protect CHU` (racine en **lecture seule** : rapport uniquement —
recommandé pour toute arborescence contenant des données patients ou internes CHU),
`-ReportDir` / `--report-dir` (dossier de sortie des rapports).

## Règles de sécurité intégrées

- Rien n'est jamais supprimé : déplacements vers `_A_METTRE_DE_COTE`, **copies** vers `_DOUBLONS_A_TRIER`.
- Les dossiers `_A_METTRE_DE_COTE`, `_DOUBLONS_A_TRIER`, `_DOSSIERS_EN_DOUBLE`, `_RAPPORT_SCAN` sont exclus du scan (les relances sont propres).
- Fichiers illisibles : listés dans le rapport (section ERREUR), jamais bloquants.
- Collisions de noms en destination : suffixe `__1`, `__2`…
- Le CSV est en UTF-8 avec BOM pour un affichage correct des accents dans Excel.

## Organisation ensuite

Le rapport sert de base de décision : université → OneDrive, personnel → Proton Drive,
100 % CHU → reste sur le PC CHU (utiliser `--protect CHU` pour garantir qu'aucun
déplacement n'y est tenté). Chaque déplacement reste manuel et validé par toi.
