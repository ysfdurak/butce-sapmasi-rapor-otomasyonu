# powerquery/Butce_Veri.pq sorgusunu gerçek M motoruyla çalıştırır; sonucu JSON olarak yazar.
# tools/pqtest/karsilastir.py bu çıktıyı Python benzetimiyle (tools/pq_veri.py) karşılaştırır.
# Kullanım (Windows, CI): pwsh tools/pqtest/calistir.ps1 -Ayar <Ayarlar.xlsx tam yolu> -Cikti <sonuc.json> [-Kapanis 2026-09] [-Mudurluk Depolama]
param(
    [Parameter(Mandatory = $true)][string]$Ayar,
    [Parameter(Mandatory = $true)][string]$Cikti,
    [string]$Kapanis = "",
    [string]$Mudurluk = ""
)
. (Join-Path $PSScriptRoot "ortak.ps1")
PQTest-Hazirla
$kok = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$klasor = Split-Path -Parent $Ayar
PQTest-Kimlik (@("File.Contents(""$Ayar"")", "Folder.Files(""$klasor\maliyet"")", "Folder.Contents(""$klasor"")") +
    @(Get-ChildItem $klasor -Filter "*.xlsx" | Where-Object { $_.Name -ne (Split-Path -Leaf $Ayar) } |
        ForEach-Object { "File.Contents(""$($_.FullName)"")" }))

$q = Get-Content -Raw -Encoding UTF8 (Join-Path $kok "powerquery/Butce_Veri.pq")
$q = PQTest-Degistir $q "AyarDosyasi" $Ayar
$q = PQTest-Degistir $q "KapanisDonemi" $Kapanis
$q = PQTest-Degistir $q "MudurlukDosyasi" $Mudurluk
PQTest-Calistir $q ("Butce_" + $(if ($Kapanis) { $Kapanis } else { "son" }) + "_" + $Mudurluk) $Cikti
