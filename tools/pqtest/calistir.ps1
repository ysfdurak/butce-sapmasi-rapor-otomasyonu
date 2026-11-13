# Power Query sorgusunu gerçek M motoruyla (Power Query SDK › PQTest.exe) çalıştırır ve sonucu JSON'a yazar.
# Kullanım (Windows): pwsh tools/pqtest/calistir.ps1 -Ayar <Ayarlar.xlsx> -Cikti <sonuc.json> [-Kapanis 2026-10]
param([string]$Ayar, [string]$Cikti, [string]$Kapanis = "")
$ErrorActionPreference = "Stop"
$kok = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$is = Join-Path $env:RUNNER_TEMP "pqtest"
New-Item -ItemType Directory -Force $is | Out-Null

$sdk = Join-Path $is "sdk"
if (-not (Test-Path "$sdk/tools/PQTest.exe")) {
    Invoke-WebRequest "https://www.nuget.org/api/v2/package/Microsoft.PowerQuery.SdkTools" -OutFile "$is/sdk.zip"
    Expand-Archive "$is/sdk.zip" $sdk -Force
}
$pqtest = "$sdk/tools/PQTest.exe"
& $pqtest version

# Uzantı: .mez = Deneme.pq içeren zip
$mez = Join-Path $is "Deneme.mez"
Remove-Item $mez -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $PSScriptRoot "Deneme.pq") -DestinationPath "$is/Deneme.zip" -Force
Move-Item "$is/Deneme.zip" $mez -Force

# Sorgu: ayar yolunu yerleştir, sonucu JSON metni olarak döndür
$q = Get-Content -Raw -Encoding UTF8 (Join-Path $kok "powerquery/Butce_Veri.pq")
$q = $q -replace 'AyarDosyasi = "[^"]*"', ('AyarDosyasi = "' + $Ayar + '"')
$q = $q -replace 'KapanisDonemi = "[^"]*"', ('KapanisDonemi = "' + $Kapanis + '"')
$sarili = "let`r`n    Sorgu = " + $q + "`r`nin`r`n    Text.FromBinary(Json.FromValue({Table.ColumnNames(Sorgu)} & Table.ToRows(Sorgu)))"
$qdosya = Join-Path $is "Butce.query.pq"
Set-Content -Path $qdosya -Value $sarili -Encoding UTF8

# Kimlik bilgisi: PQTest her veri kaynağı için ayrı ister; her kaynağa tek kaynaklı küçük bir sorgu yazılır.
$ayarKlasoru = Split-Path -Parent $Ayar
$kaynaklar = @(
    "File.Contents(""$Ayar"")",
    "File.Contents(""$ayarKlasoru\Kontrol.xlsx"")",
    "File.Contents(""$ayarKlasoru\Endeks_2026_10.xlsx"")",
    "Folder.Files(""$ayarKlasoru\maliyet"")"
)
$i = 0
foreach ($ifade in $kaynaklar) {
    $i++
    $kq = Join-Path $is "Kaynak$i.query.pq"
    Set-Content -Path $kq -Value $ifade -Encoding UTF8
    foreach ($ak in @("Windows", "Anonymous")) {
        $sablon = (& $pqtest credential-template -e $mez -q $kq -p --authenticationKind $ak 2>&1) -join "`n"
        Write-Host "== $ifade / $ak"; Write-Host $sablon
        if ($LASTEXITCODE -eq 0) {
            $sonuc = ($sablon | & $pqtest set-credential -e $mez -q $kq -p 2>&1) -join "`n"
            Write-Host $sonuc
            if ($LASTEXITCODE -eq 0) { break }
        }
    }
}
& $pqtest run-test -e $mez -q $qdosya -p 2>&1 | Out-File $Cikti -Encoding utf8
Get-Content $Cikti -TotalCount 120
