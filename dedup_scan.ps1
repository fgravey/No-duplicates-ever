<#
.SYNOPSIS
    Scan local : doublons exacts, fichiers de meme nom a etats differents, fichiers "ex-partenaire" a mettre de cote.

.DESCRIPTION
    100 % local, aucune installation requise (Windows PowerShell 5.1+), aucun envoi reseau,
    AUCUNE suppression de fichier.
    - Doublons exacts : hash SHA-256 (Get-FileHash).
    - Memes noms / variantes (v2, final, copie, (1), dates) : groupes par nom nettoye.
    - Noms recherches (noms de fichiers + contenu texte de txt/csv/docx/xlsx/pptx/pdf/eml...) :
      elsa, eric, chantal, laetitia, estelle, josy, michelle (mot entier, insensible casse/accents).
    - Actions : les fichiers tagues sont DEPLACES (jamais supprimes) dans _A_METTRE_DE_COTE,
      les DOSSIERS entierement dupliques sont DEPLACES dans _DOSSIERS_EN_DOUBLE,
      les doublons fichier sont COPIES dans _DOUBLONS_A_TRIER, rapports CSV + HTML dans _RAPPORT_SCAN.
      L'exemplaire conserve de chaque groupe est marque "CONSERVE EN PLACE" dans le rapport CSV.
    - Une racine contenant une sous-chaine de -Protect est en LECTURE SEULE (rapport seul) :
      ideal pour l'arborescence CHU avec donnees patients.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\dedup_scan.ps1 -Auto -Protect "CHU" -DryRun
    powershell -ExecutionPolicy Bypass -File .\dedup_scan.ps1 -Roots "D:\MesDocuments" -Protect "CHU" -DryRun
    powershell -ExecutionPolicy Bypass -File .\dedup_scan.ps1 -Roots "D:\MesDocuments","D:\OneDrive","D:\ProtonDrive"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$Roots = @(),

    [switch]$Auto,

    [string[]]$Protect = @(),

    [switch]$DryRun,

    [string]$ReportDir,

    [long]$MaxContent = 20MB
)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

if ($Auto) {
    $Roots = @([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady -and $_.DriveType -ne [System.IO.DriveType]::CDRom } | ForEach-Object { $_.Name })
    Write-Host "[*] Mode -Auto, lecteurs decouverts : $($Roots -join ', ')"
}
if (-not $Auto -and $Roots.Count -eq 0) { throw 'Indique -Roots <chemins...> ou -Auto' }

$script:NomsEx = @('elsa', 'eric', 'chantal', 'laetitia', 'estelle', 'josy', 'michelle')
$script:DossiersIgnores = @('_A_METTRE_DE_COTE', '_DOUBLONS_A_TRIER', '_DOSSIERS_EN_DOUBLE', '_RAPPORT_SCAN',
    'Windows', '$Recycle.Bin', 'System Volume Information', 'Program Files', 'Program Files (x86)',
    'ProgramData', 'MSOCache', 'PerfLogs', 'node_modules', '.git', '__pycache__', '.Trash', 'Caches')
$script:FichiersIgnores = @('Thumbs.db', '.DS_Store', 'pagefile.sys', 'hiberfil.sys', 'swapfile.sys', 'dumpfile.sys')
$script:ExtText = @('.txt', '.md', '.csv', '.tsv', '.log', '.eml', '.html', '.htm', '.xml', '.json', '.ini', '.cfg', '.tex', '.srt')
$script:ExtZip = @('.docx', '.pptx', '.xlsx', '.odt', '.ods', '.epub')
$script:PatternNoms = '(?<![a-z0-9])(' + (($script:NomsEx | Sort-Object -Descending Length) -join '|') + ')(?![a-z0-9])'
$script:PatternVariante = '[\s_\-.]*(v\s?\d+|final(e)?|def(initif(ve)?)?|copie|copy|bis|ter|quater|\(\s*\d+\s*\)|-\d+|\d{4}-\d{2}-\d{2}|\d{8}|new|old|brouillon|draft)[\s_\-.]*$'

function Normaliser([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $d = $s.Normalize([System.Text.NormalizationForm]::FormD)
    return (($d -replace '\p{M}', '')).ToLowerInvariant()
}

function Nettoyer-Tige([string]$tige) {
    $s = Normaliser $tige
    $prev = $null
    while ($prev -ne $s) {
        $prev = $s
        $s = [regex]::Replace($s, $script:PatternVariante, '')
    }
    return $s.Trim(' ', '_', '-', '.')
}

function EstCheminIgnore([string]$path) {
    $segments = $path -split [regex]::Escape([System.IO.Path]::DirectorySeparatorChar)
    foreach ($d in $script:DossiersIgnores) {
        if ($segments -contains $d) { return $true }
    }
    return $false
}

function CheminProtege([string]$chemin) {
    $n = Normaliser $chemin
    foreach ($p in $Protect) {
        $np = Normaliser $p
        if ($n.Contains($np)) { return $true }
    }
    return $false
}

function Lire-Manifeste([string]$racine) {
    $p = Join-Path $racine '_DOUBLONS_A_TRIER\.deja_copies.txt'
    if (Test-Path -LiteralPath $p) { return @(Get-Content -LiteralPath $p) } 
    return @()
}

function Ajouter-Manifeste([string]$racine, [string]$ligne) {
    $dir = Join-Path $racine '_DOUBLONS_A_TRIER'
    New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
    Add-Content -LiteralPath (Join-Path $dir '.deja_copies.txt') -Value $ligne -Encoding UTF8
}

function Get-TexteFichier([System.IO.FileInfo]$f) {
    $ext = $f.Extension.ToLowerInvariant()
    if ($f.Length -gt $MaxContent) { return $null }
    if ($ext -notin ($script:ExtText + $script:ExtZip + @('.pdf'))) { return $null }
    try {
        if ($ext -in $script:ExtText) {
            return [System.IO.File]::ReadAllText($f.FullName)
        }
        if ($ext -in $script:ExtZip) {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
            try {
                $sb = New-Object System.Text.StringBuilder
                $total = 0L
                foreach ($entry in $zip.Entries) {
                    if (-not $entry.Name.EndsWith('.xml')) { continue }
                    if ($total -gt $MaxContent) { break }
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    $txt = $reader.ReadToEnd()
                    $reader.Close()
                    $total += $txt.Length
                    [void]$sb.Append($txt -replace '<[^>]+>', ' ')
                }
                return $sb.ToString()
            } finally { $zip.Dispose() }
        }
        if ($ext -eq '.pdf') {
            $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
            if ($bytes.Length -gt $MaxContent) { $bytes = $bytes[0..($MaxContent - 1)] }
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.Append([System.Text.Encoding]::Latin1.GetString($bytes))
            $flux = New-Object System.IO.MemoryStream(,$bytes)
            $reader = New-Object System.IO.StreamReader($flux)
            $raw = $reader.ReadToEnd()
            $reader.Close(); $flux.Close()
            foreach ($m in [regex]::Matches($raw, 'stream\r?\n(.*?)endstream', 'Singleline')) {
                $bloc = [System.Text.Encoding]::Latin1.GetBytes($m.Groups[1].Value.Trim())
                $offsets = @(0, 2)
                foreach ($off in $offsets) {
                    if ($bloc.Length -le $off + 4) { continue }
                    try {
                        $ms = New-Object System.IO.MemoryStream(,$bloc[$off..($bloc.Length - 1)])
                        $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Decompress)
                        $sr = New-Object System.IO.StreamReader($ds)
                        [void]$sb.Append($sr.ReadToEnd())
                        $sr.Close(); $ds.Close(); $ms.Close()
                        break
                    } catch { }
                }
            }
            return $sb.ToString()
        }
    } catch { return $null }
    return $null
}

function Get-Relatif([string]$base, [string]$chemin) {
    return $chemin.Substring($base.Length).TrimStart('\', '/')
}

function EstSousDossier([string]$chemin, [string]$racine) {
    return ($chemin -eq $racine) -or $chemin.StartsWith($racine + '\')
}

function New-RecDossier($item, [string]$sig) {
    [pscustomobject]@{
        Chemin = $item.Dossier
        Taille = ($item.Recs | Measure-Object -Property Taille -Sum).Sum
        Mtime  = ($item.Recs | Measure-Object -Property Mtime -Maximum).Maximum
        Sha256 = $sig
    }
}

function CheminUnique([string]$dest) {
    if (-not (Test-Path -LiteralPath $dest)) { return $dest }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($dest)
    $ext = [System.IO.Path]::GetExtension($dest)
    $dir = [System.IO.Path]::GetDirectoryName($dest)
    $i = 1
    while (Test-Path -LiteralPath (Join-Path $dir ("{0}__{1}{2}" -f $base, $i, $ext))) { $i++ }
    return Join-Path $dir ("{0}__{1}{2}" -f $base, $i, $ext)
}

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$tous = @()
$erreurs = @()

foreach ($racine in $Roots) {
    if (-not (Test-Path -LiteralPath $racine -PathType Container)) {
        Write-Warning "Racine ignoree (introuvable) : $racine"
        continue
    }
    $racineAbs = (Resolve-Path -LiteralPath $racine).Path
    Write-Host "[*] Scan de $racineAbs ..."
    $fichiers = Get-ChildItem -LiteralPath $racineAbs -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { -not (EstCheminIgnore $_.FullName) -and $_.Name -notlike 'rapport_*' -and $_.Name -notin $script:FichiersIgnores }
    $n = 0
    foreach ($f in $fichiers) {
        $n++
        try {
            $h = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        } catch {
            $erreurs += [pscustomobject]@{ Chemin = $f.FullName; Erreur = $_.Exception.Message }
            continue
        }
        $texte = Get-TexteFichier $f
        $nomNorm = Normaliser $f.Name
        $tous += [pscustomobject]@{
            Racine = $racineAbs
            Chemin = $f.FullName
            Relatif = $f.FullName.Substring($racineAbs.Length).TrimStart('\', '/')
            Nom = $f.Name
            Ext = $f.Extension.ToLowerInvariant()
            Taille = $f.Length
            Mtime = $f.LastWriteTime
            Sha256 = $h
            ToucheNom = [bool]([regex]::IsMatch($nomNorm, $script:PatternNoms))
            ToucheContenu = ($null -ne $texte -and [bool]([regex]::IsMatch((Normaliser $texte), $script:PatternNoms)))
        }
        if ($n % 500 -eq 0) { Write-Host "    $n fichiers..." }
    }
    Write-Host "    $n fichiers analyses, $($erreurs.Count) erreur(s) au total"
}

if ($tous.Count -eq 0) { Write-Host 'Aucun fichier trouve.'; exit 1 }

$sections = [ordered]@{}

$parHash = $tous | Group-Object Sha256 | Where-Object { $_.Count -gt 1 } | Sort-Object { -($_.Group | Measure-Object Taille -Sum).Sum }
$sections['DOUBLONS EXACTS (meme contenu, hash identique)'] = @(
    foreach ($g in $parHash) {
        [pscustomobject]@{ Groupe = "$($g.Name.Substring(0,12))... ($($g.Count) fichiers)"; Fichiers = ($g.Group | Sort-Object Chemin) }
    })

$parNom = $tous | Group-Object { Normaliser $_.Nom } | Where-Object { $_.Count -gt 1 -and (($_.Group | Select-Object -ExpandProperty Sha256 -Unique).Count -gt 1) }
$sections['MEMES NOMS, CONTENUS DIFFERENTS (versions)'] = @(
    foreach ($g in ($parNom | Sort-Object Name)) {
        [pscustomobject]@{ Groupe = $g.Name; Fichiers = ($g.Group | Sort-Object Chemin) }
    })

$parTige = $tous | Group-Object { (Nettoyer-Tige ([System.IO.Path]::GetFileNameWithoutExtension($_.Nom))) + $_.Ext.ToLowerInvariant() } |
    Where-Object { $_.Count -gt 1 -and (($_.Group | Select-Object -ExpandProperty Sha256 -Unique).Count -gt 1) }
$sections['VARIANTES DE NOM (v2, final, copie, (1), dates...)'] = @(
    foreach ($g in ($parTige | Sort-Object Name)) {
        [pscustomobject]@{ Groupe = "$($g.Name) ($($g.Count) fichiers)"; Fichiers = ($g.Group | Sort-Object Chemin) }
    })

$touches = @($tous | Where-Object { $_.ToucheNom -or $_.ToucheContenu } | Sort-Object Chemin)
$sections['FICHIERS A METTRE DE COTE (noms : ' + ($script:NomsEx -join ', ') + ')'] = @(
    foreach ($r in $touches) {
        $motif = @(); if ($r.ToucheNom) { $motif += 'nom' }; if ($r.ToucheContenu) { $motif += 'contenu' }
        [pscustomobject]@{ Groupe = "$($motif -join '+') : $($r.Nom)"; Fichiers = @($r) }
    })

$actions = @()
$parDossier = $tous | Group-Object { Split-Path -Parent $_.Chemin }
$sigs = @{}
foreach ($g in $parDossier) {
    $d = $g.Name
    $lignes = $g.Group | ForEach-Object { (Get-Relatif $d $_.Chemin) + '|' + $_.Sha256 } | Sort-Object
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(($lignes -join "`n")))
    $sha.Dispose()
    $sig = ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    if (-not $sigs.ContainsKey($sig)) { $sigs[$sig] = @() }
    $sigs[$sig] = $sigs[$sig] + ,(@{ Dossier = $d; Recs = $g.Group })
}

$dossiersDoublons = @()
$deplaces = @()
$clesSig = $sigs.Keys | Sort-Object { ($sigs[$_][0].Dossier -split '\\').Count }
foreach ($sig in $clesSig) {
    $lst = @($sigs[$sig] | Where-Object {
        $item = $_
        -not ($item.Recs | Where-Object { $_.ToucheNom -or $_.ToucheContenu }) -and
        -not ($deplaces | Where-Object { EstSousDossier $item.Dossier $_ })
    })
    if ($lst.Count -lt 2) { continue }
    $lst = @($lst | Sort-Object @{Expression = { -($_.Recs | ForEach-Object { $_.Mtime.Ticks } | Measure-Object -Maximum).Maximum }}, Dossier)
    $keeper = $lst[0]
    $rest = @($lst[1..($lst.Count - 1)])
    $dossiersDoublons += [pscustomobject]@{ Sig = $sig; Keeper = $keeper; Rest = $rest }
    $deplaces += $rest | ForEach-Object { $_.Dossier }
}

$sections['DOSSIERS ENTIEREMENT DUPLIQUES (contenu identique ailleurs)'] = @(
    foreach ($dd in $dossiersDoublons) {
        [pscustomobject]@{
            Groupe = "exemplaire conservé : $($dd.Keeper.Dossier)"
            Fichiers = @(New-RecDossier $dd.Keeper $dd.Sig) + @($dd.Rest | ForEach-Object { New-RecDossier $_ $dd.Sig })
        }
    })

$conserves = @{}
foreach ($g in $parHash) {
    $tries = $g.Group | Sort-Object @{Expression = { - $_.Mtime.Ticks }}, Chemin
    $conserves[$tries[0].Chemin] = $true
}
foreach ($dd in $dossiersDoublons) { $conserves[$dd.Keeper.Dossier] = $true }

$actions = @()
$manifestes = @{}
foreach ($r in $touches) {
    if (CheminProtege $r.Chemin) { continue }
    $destDir = Join-Path $r.Racine '_A_METTRE_DE_COTE'
    $dest = CheminUnique (Join-Path $destDir $r.Relatif)
    if (-not $DryRun) {
        New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($dest)) -Force -ErrorAction SilentlyContinue | Out-Null
        Move-Item -LiteralPath $r.Chemin -Destination $dest -Force
    }
    $actions += [pscustomobject]@{ Raison = 'nom/contenu ex-partenaire' + ($(if ($DryRun) { ' (SIMULATION)' })); Source = $r.Chemin; Destination = $dest; Sha256 = $r.Sha256 }
}

foreach ($dd in $dossiersDoublons) {
    foreach ($item in $dd.Rest) {
        $d = $item.Dossier
        if (CheminProtege $d) { continue }
        $racine = $item.Recs[0].Racine
        $dest = CheminUnique (Join-Path (Join-Path $racine '_DOSSIERS_EN_DOUBLE') (Split-Path -Leaf $d))
        if (-not $DryRun) {
            New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($dest)) -Force -ErrorAction SilentlyContinue | Out-Null
            Move-Item -LiteralPath $d -Destination $dest -Force
        }
        $actions += [pscustomobject]@{ Raison = "dossier entierement doublon de $($dd.Keeper.Dossier)" + ($(if ($DryRun) { ' (SIMULATION)' })); Source = $d; Destination = $dest; Sha256 = $dd.Sig }
    }
}

foreach ($g in $parHash) {
    $tries = $g.Group | Sort-Object @{Expression = { - $_.Mtime.Ticks }}, Chemin
    $gardien = $tries[0]
    foreach ($r in $tries[1..($tries.Count - 1)]) {
        if (CheminProtege $r.Chemin) { continue }
        if ($r.ToucheNom -or $r.ToucheContenu) { continue }
        if (-not (Test-Path -LiteralPath $r.Chemin)) { continue }
        if ($deplaces | Where-Object { $r.Chemin.StartsWith($_ + '\') }) { continue }
        if (-not $manifestes.ContainsKey($r.Racine)) { $manifestes[$r.Racine] = @(Lire-Manifeste $r.Racine) }
        $cle = "$($r.Sha256)`t$($r.Chemin)"
        if ($manifestes[$r.Racine] -contains $cle) { continue }
        $destDir = Join-Path $r.Racine '_DOUBLONS_A_TRIER'
        $dest = CheminUnique (Join-Path $destDir $r.Relatif)
        if (-not $DryRun) {
            New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($dest)) -Force -ErrorAction SilentlyContinue | Out-Null
            Copy-Item -LiteralPath $r.Chemin -Destination $dest -Force
            $manifestes[$r.Racine] += $cle
            Ajouter-Manifeste $r.Racine $cle
        }
        $actions += [pscustomobject]@{ Raison = "doublon de $($gardien.Chemin)" + ($(if ($DryRun) { ' (SIMULATION)' })); Source = $r.Chemin; Destination = $dest; Sha256 = $r.Sha256 }
    }
}

if ([string]::IsNullOrWhiteSpace($ReportDir)) {
    $ReportDir = Join-Path (Resolve-Path -LiteralPath $Roots[0]).Path '_RAPPORT_SCAN'
}
New-Item -ItemType Directory -Path $ReportDir -Force -ErrorAction SilentlyContinue | Out-Null

$lignes = @()
foreach ($categorie in $sections.Keys) {
    foreach ($g in $sections[$categorie]) {
        foreach ($r in $g.Fichiers) {
            $act = $actions | Where-Object { $_.Source -eq $r.Chemin } | Select-Object -First 1
            $lignes += [pscustomobject]@{
                Categorie = $categorie
                Groupe = $g.Groupe
                Chemin = $r.Chemin
                TailleOctets = $r.Taille
                ModifieLe = $r.Mtime.ToString('yyyy-MM-dd HH:mm:ss')
                Sha256 = $r.Sha256
                Action = $(if ($act) { $act.Raison } elseif ($conserves.ContainsKey($r.Chemin)) { 'CONSERVE EN PLACE (exemplaire unique garde)' } else { '' })
                Destination = $(if ($act) { $act.Destination } else { '' })
            }
        }
    }
}
foreach ($e in $erreurs) {
    $lignes += [pscustomobject]@{ Categorie = 'ERREUR'; Groupe = ''; Chemin = $e.Chemin; TailleOctets = ''; ModifieLe = ''; Sha256 = ''; Action = $e.Erreur; Destination = '' }
}
$csvPath = Join-Path $ReportDir "rapport_$ts.csv"
$lignes | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

$htmlPath = Join-Path $ReportDir "rapport_$ts.html"
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("<html><head><meta charset='utf-8'><title>Rapport scan $ts</title><style>body{font-family:sans-serif;margin:2em} table{border-collapse:collapse;margin-bottom:2em} td,th{border:1px solid #bbb;padding:4px 8px;font-size:13px} th{background:#e8eef7}</style></head><body>")
[void]$sb.Append("<h1>Rapport de scan du $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</h1>")
[void]$sb.Append("<p>Racines scannees : $($Roots -join ', ')<br>Mode : $(if ($DryRun) {'SIMULATION (dry-run)'} else {'actions appliquees'}) - protections : $(if ($Protect) { $Protect -join ', ' } else { 'aucune' }) - aucun fichier n'est supprime.</p>")
foreach ($categorie in $sections.Keys) {
    [void]$sb.Append("<h2>$([System.Net.WebUtility]::HtmlEncode($categorie)) ($($sections[$categorie].Count) groupe(s))</h2>")
    foreach ($g in $sections[$categorie]) {
        [void]$sb.Append("<h3>$([System.Net.WebUtility]::HtmlEncode($g.Groupe))</h3><table><tr><th>Chemin</th><th>Taille</th><th>Modifie le</th><th>SHA-256</th></tr>")
        foreach ($r in $g.Fichiers) {
            [void]$sb.Append("<tr><td>$([System.Net.WebUtility]::HtmlEncode($r.Chemin))</td><td>$($r.Taille)</td><td>$($r.Mtime.ToString('yyyy-MM-dd HH:mm:ss'))</td><td>$($r.Sha256.Substring(0,16))…</td></tr>")
        }
        [void]$sb.Append('</table>')
    }
}
if ($actions.Count -gt 0) {
    [void]$sb.Append("<h2>Actions effectuees ($($actions.Count))</h2><table><tr><th>Raison</th><th>Source</th><th>Destination</th></tr>")
    foreach ($a in $actions) {
        [void]$sb.Append("<tr><td>$([System.Net.WebUtility]::HtmlEncode($a.Raison))</td><td>$([System.Net.WebUtility]::HtmlEncode($a.Source))</td><td>$([System.Net.WebUtility]::HtmlEncode($a.Destination))</td></tr>")
    }
    [void]$sb.Append('</table>')
}
[void]$sb.Append('</body></html>')
[System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), [System.Text.Encoding]::UTF8)

Write-Host ''
Write-Host "Fichiers scannes          : $($tous.Count)"
Write-Host "Groupes de doublons       : $($parHash.Count)"
Write-Host "Dossiers tout doublons    : $($dossiersDoublons.Count) ($($deplaces.Count) a deplacer)"
Write-Host "Groupes memes noms        : $($sections['MEMES NOMS, CONTENUS DIFFERENTS (versions)'].Count)"
Write-Host "Groupes variantes         : $($sections['VARIANTES DE NOM (v2, final, copie, (1), dates...)'].Count)"
Write-Host "Fichiers a mettre de cote : $($touches.Count)"
Write-Host "Actions                   : $($actions.Count) $(if ($DryRun) { '(SIMULATION)' })"
Write-Host "Erreurs de lecture        : $($erreurs.Count)"
Write-Host "Rapport CSV               : $csvPath"
Write-Host "Rapport HTML              : $htmlPath"
