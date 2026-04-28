param(
    [Parameter(Mandatory = $true)][ValidateSet('inventory','hash','search','sqlite-probe','diff')] [string]$Mode,
    [string]$Area,
    [string]$AreaA,
    [string]$AreaB,
    [Parameter(Mandatory = $true)] [string]$OutDir,
    [string[]]$Terms = @('favorite','favourites','favorites','fave','chart','table','results','foldername','favename')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-NormPath([string]$p) {
    $normalized = $p -replace '\\','/'
    return (Resolve-Path -LiteralPath $normalized).Path
}

function Ensure-OutDir([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

function Get-AllFiles([string]$root) {
    Get-ChildItem -LiteralPath $root -Recurse -File -Force
}

Ensure-OutDir -path $OutDir

switch ($Mode) {
    'inventory' {
        if (-not $Area) { throw 'Mode inventory requires -Area' }
        $areaPath = Resolve-NormPath $Area
        $files = Get-AllFiles $areaPath

        $files |
            Select-Object @{N='type';E={'f'}}, Length, LastWriteTimeUtc, FullName |
            Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'inventory.tsv')

        $files |
            Group-Object { if ($_.Extension) { $_.Extension.ToLower() } else { '[noext]' } } |
            Sort-Object Count -Descending |
            Select-Object @{N='extension';E={$_.Name}}, Count |
            Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'extension_summary.tsv')

        $candidateExt = @('.db','.sqlite','.sqlite3','.mdb','.accdb','.xml','.json','.ini','.cfg','.txt','.dat','.bin','.zip','.7z','.nx1')
        $files |
            Where-Object { $candidateExt -contains $_.Extension.ToLower() } |
            Select-Object Length, LastWriteTimeUtc, Extension, FullName |
            Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'candidate_files.tsv')

        $files |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object LastWriteTimeUtc, Length, FullName |
            Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'recent_files.tsv')

        $files |
            Where-Object { $_.Length -lt 20480 } |
            Sort-Object Length |
            Select-Object Length, LastWriteTimeUtc, FullName |
            Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'small_files_lt20k.tsv')
    }

    'hash' {
        if (-not $Area) { throw 'Mode hash requires -Area' }
        $areaPath = Resolve-NormPath $Area
        $rows = foreach ($f in (Get-AllFiles $areaPath)) {
            $h = Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256
            [PSCustomObject]@{ sha256 = $h.Hash.ToLower(); size = $f.Length; path = $f.FullName }
        }
        $rows | Export-Csv -Delimiter "`t" -NoTypeInformation -Path (Join-Path $OutDir 'file_hashes_sha256.tsv')
    }

    'search' {
        if (-not $Area) { throw 'Mode search requires -Area' }
        $areaPath = Resolve-NormPath $Area
        $pattern = ($Terms | ForEach-Object { [Regex]::Escape($_) }) -join '|'
        $textOut = Join-Path $OutDir 'text_matches.tsv'
        $binOut = Join-Path $OutDir 'binary_strings_matches.tsv'
        @() | Export-Csv -Delimiter "`t" -NoTypeInformation -Path $textOut
        @() | Export-Csv -Delimiter "`t" -NoTypeInformation -Path $binOut

        foreach ($f in (Get-AllFiles $areaPath)) {
            $ext = $f.Extension.ToLower()
            $isText = @('.txt','.ini','.cfg','.xml','.json','.csv','.log','.md') -contains $ext
            if ($isText) {
                $i = 0
                Get-Content -LiteralPath $f.FullName -ErrorAction SilentlyContinue | ForEach-Object {
                    $i++
                    if ($_ -match $pattern) {
                        [PSCustomObject]@{ path = $f.FullName; line_no = $i; line = $_.Trim() } |
                            Export-Csv -Delimiter "`t" -NoTypeInformation -Path $textOut -Append
                    }
                }
            }
            else {
                try {
                    $raw = & strings -a -n 4 $f.FullName 2>$null
                    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 1) {
                        $j = 0
                        foreach ($s in $raw) {
                            $j++
                            if ($s -match $pattern) {
                                [PSCustomObject]@{ path = $f.FullName; string_line_no = $j; string = $s.Trim() } |
                                    Export-Csv -Delimiter "`t" -NoTypeInformation -Path $binOut -Append
                            }
                        }
                    }
                }
                catch {
                }
            }
        }
    }

    'sqlite-probe' {
        if (-not $Area) { throw 'Mode sqlite-probe requires -Area' }
        $areaPath = Resolve-NormPath $Area
        $py = @'
import json, sqlite3, re
from pathlib import Path
area=Path(r"__AREA__")
out=Path(r"__OUT__")
terms=__TERMS__
pat=re.compile("|".join(re.escape(t) for t in terms), re.I)
obj={"databases": []}
for p in sorted(area.rglob("*")):
    if not p.is_file():
        continue
    if p.suffix.lower() not in {".db",".sqlite",".sqlite3"}:
        continue
    di={"path": str(p), "tables": []}
    try:
        con=sqlite3.connect(f"file:{p}?mode=ro", uri=True)
        cur=con.cursor()
        tables=[r[0] for r in cur.execute("select name from sqlite_master where type='table' order by name")]
        for t in tables:
            cols=[r[1] for r in cur.execute(f"pragma table_info('{t}')")]
            hit=bool(pat.search(t) or any(pat.search(c) for c in cols))
            sample=[]
            if hit:
                for row in cur.execute(f"select * from '{t}' limit 5"):
                    sample.append([str(x)[:500] for x in row])
            di["tables"].append({"name": t, "columns": cols, "keyword_hit": hit, "sample_rows": sample})
        con.close()
    except Exception as e:
        di["error"]=str(e)
    obj["databases"].append(di)
out.write_text(json.dumps(obj, indent=2), encoding="utf-8")
'@
        $py = $py.Replace('__AREA__', $areaPath.Replace('\','\\'))
        $py = $py.Replace('__OUT__', (Join-Path $OutDir 'sqlite_probe.json').Replace('\','\\'))
        $termsJson = ($Terms | ConvertTo-Json -Compress)
        $py = $py.Replace('__TERMS__', $termsJson)
        $tmpPy = Join-Path $OutDir '_sqlite_probe_tmp.py'
        Set-Content -LiteralPath $tmpPy -Value $py -Encoding UTF8
        python $tmpPy
        Remove-Item -LiteralPath $tmpPy -Force -ErrorAction SilentlyContinue
    }

    'diff' {
        if (-not $AreaA -or -not $AreaB) { throw 'Mode diff requires -AreaA and -AreaB' }
        $a = Resolve-NormPath $AreaA
        $b = Resolve-NormPath $AreaB

        $mapA = @{}
        foreach ($f in (Get-AllFiles $a)) {
            $rel = $f.FullName.Substring($a.Length).TrimStart('\\','/') -replace '\\','/'
            $mapA[$rel] = $f
        }
        $mapB = @{}
        foreach ($f in (Get-AllFiles $b)) {
            $rel = $f.FullName.Substring($b.Length).TrimStart('\\','/') -replace '\\','/'
            $mapB[$rel] = $f
        }

        $all = ($mapA.Keys + $mapB.Keys | Sort-Object -Unique)
        $out = Join-Path $OutDir 'area_diff.tsv'
        @() | Export-Csv -Delimiter "`t" -NoTypeInformation -Path $out

        foreach ($rel in $all) {
            $fa = $mapA[$rel]
            $fb = $mapB[$rel]
            if (-not $fa) {
                $hb = (Get-FileHash -LiteralPath $fb.FullName -Algorithm SHA256).Hash.ToLower()
                [PSCustomObject]@{status='only_in_b'; relpath=$rel; size_a=''; size_b=$fb.Length; sha256_a=''; sha256_b=$hb} |
                    Export-Csv -Delimiter "`t" -NoTypeInformation -Path $out -Append
                continue
            }
            if (-not $fb) {
                $ha = (Get-FileHash -LiteralPath $fa.FullName -Algorithm SHA256).Hash.ToLower()
                [PSCustomObject]@{status='only_in_a'; relpath=$rel; size_a=$fa.Length; size_b=''; sha256_a=$ha; sha256_b=''} |
                    Export-Csv -Delimiter "`t" -NoTypeInformation -Path $out -Append
                continue
            }
            if ($fa.Length -ne $fb.Length) {
                [PSCustomObject]@{status='size_diff'; relpath=$rel; size_a=$fa.Length; size_b=$fb.Length; sha256_a=''; sha256_b=''} |
                    Export-Csv -Delimiter "`t" -NoTypeInformation -Path $out -Append
                continue
            }
            $ha = (Get-FileHash -LiteralPath $fa.FullName -Algorithm SHA256).Hash.ToLower()
            $hb = (Get-FileHash -LiteralPath $fb.FullName -Algorithm SHA256).Hash.ToLower()
            if ($ha -ne $hb) {
                [PSCustomObject]@{status='hash_diff'; relpath=$rel; size_a=$fa.Length; size_b=$fb.Length; sha256_a=$ha; sha256_b=$hb} |
                    Export-Csv -Delimiter "`t" -NoTypeInformation -Path $out -Append
            }
        }
    }
}

Write-Host "Completed $Mode. Output folder: $OutDir"
