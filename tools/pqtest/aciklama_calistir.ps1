# powerquery/Butce_Aciklama.pq sorgusunu gerçek M motoruyla çalıştırır. Excel.CurrentWorkbook() (önceki onay tablosu ve
# adlandırılmış hücreler) Excel dışında olmadığından -Kitap dosyasındaki M tablosuyla değiştirilir; saat sabitlenir.
# Kullanım: pwsh tools/pqtest/aciklama_calistir.ps1 -Ayar <Ayarlar.xlsx> -Klasor <dönen dosyalar> -Kitap <stub.pq> -Cikti <json> -Ad <ad>
param(
    [Parameter(Mandatory = $true)][string]$Ayar,
    [Parameter(Mandatory = $true)][string]$Klasor,
    [Parameter(Mandatory = $true)][string]$Kitap,
    [Parameter(Mandatory = $true)][string]$Cikti,
    [string]$Ad = "Aciklama"
)
. (Join-Path $PSScriptRoot "ortak.ps1")
PQTest-Hazirla
$kok = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
PQTest-Kimlik @("File.Contents(""$Ayar"")", "Folder.Files(""$Klasor"")")

$q = Get-Content -Raw -Encoding UTF8 (Join-Path $kok "powerquery/Butce_Aciklama.pq")
$q = PQTest-Degistir $q "AyarDosyasi" $Ayar
$q = PQTest-Degistir $q "AciklamaKlasoru" $Klasor
$stub = Get-Content -Raw -Encoding UTF8 $Kitap
if ($q -notmatch 'CalismaKitabi = Excel\.CurrentWorkbook\(\),') { throw "CalismaKitabi satırı bulunamadı" }
$q = $q.Replace("CalismaKitabi = Excel.CurrentWorkbook(),", "CalismaKitabi = " + $stub.Trim() + ",")
$q = $q.Replace("DateTime.LocalNow()", "#datetime(2026, 11, 13, 10, 0, 0)")
PQTest-Calistir $q $Ad $Cikti
