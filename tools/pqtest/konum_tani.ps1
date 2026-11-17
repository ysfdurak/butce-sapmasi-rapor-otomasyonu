# Geçici tanı: otomatik konumun M motorunda gördüğü değerler.
param([Parameter(Mandatory = $true)][string]$Ayar)
. (Join-Path $PSScriptRoot "ortak.ps1")
PQTest-Hazirla
$kok = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$klasor = Split-Path -Parent $Ayar
PQTest-Kimlik (@("File.Contents(""$Ayar"")", "Folder.Files(""$klasor\maliyet"")", "Folder.Contents(""$klasor"")") +
    @(Get-ChildItem $klasor -Filter "*.xlsx" | Where-Object { $_.Name -ne (Split-Path -Leaf $Ayar) } |
        ForEach-Object { "File.Contents(""$($_.FullName)"")" }))
$q = Get-Content -Raw -Encoding UTF8 (Join-Path $kok "powerquery/Butce_Veri.pq")
$q = PQTest-Degistir $q "AyarDosyasi" $Ayar
$tani = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot "konum_tani.m")
$q = $q -replace '(?s)\r?\nin\r?\n\s+Sonuc\s*$', ("`r`nin`r`n" + $tani)
PQTest-Calistir $q "KonumTani" "$env:RUNNER_TEMP\tani.json"
Get-Content "$env:RUNNER_TEMP\tani.json" | Select-Object -First 60
