# powerquery/Butce_Veri.pq sorgusunu Excel'deki Power Query ile aynı M motoruyla (Power Query SDK › PQTest.exe)
# çalıştırır; sonucu JSON olarak yazar. tools/pqtest/karsilastir.py bu çıktıyı Python benzetimiyle karşılaştırır.
# Kullanım (Windows, CI): pwsh tools/pqtest/calistir.ps1 -Ayar <Ayarlar.xlsx tam yolu> -Cikti <sonuc.json> [-Kapanis 2026-09]
param(
    [Parameter(Mandatory = $true)][string]$Ayar,
    [Parameter(Mandatory = $true)][string]$Cikti,
    [string]$Kapanis = ""
)
$ErrorActionPreference = "Stop"
$kok = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$is = Join-Path ([IO.Path]::GetTempPath()) "pqtest"
New-Item -ItemType Directory -Force $is | Out-Null

# Power Query SDK araçları (NuGet). PQTest bir veri kaynağı uzantısı ister: boş Deneme.pq → Deneme.mez (zip).
$sdk = Join-Path $is "sdk"
if (-not (Test-Path "$sdk/tools/PQTest.exe")) {
    Invoke-WebRequest "https://www.nuget.org/api/v2/package/Microsoft.PowerQuery.SdkTools" -OutFile "$is/sdk.zip"
    Expand-Archive "$is/sdk.zip" $sdk -Force
}
$pqtest = "$sdk/tools/PQTest.exe"
$mez = Join-Path $is "Deneme.mez"
Compress-Archive -Path (Join-Path $PSScriptRoot "Deneme.pq") -DestinationPath "$is/Deneme.zip" -Force
Move-Item "$is/Deneme.zip" $mez -Force

# Sorgu: iki ayar satırı doldurulur, tablo JSON metni olarak döndürülür.
$q = Get-Content -Raw -Encoding UTF8 (Join-Path $kok "powerquery/Butce_Veri.pq")
# Yalnız satır başındaki tanımlar: koşuldaki `KapanisDonemi = ""` karşılaştırmasına dokunulmaz.
$q = $q -replace '(?m)^(\s*)AyarDosyasi = "[^"]*"', ('${1}AyarDosyasi = "' + $Ayar + '"')
$q = $q -replace '(?m)^(\s*)KapanisDonemi = "[^"]*"', ('${1}KapanisDonemi = "' + $Kapanis + '"')
$sarili = "let`r`n    Sorgu = " + $q + "`r`nin`r`n    Text.FromBinary(Json.FromValue({Table.ColumnNames(Sorgu)} & Table.ToRows(Sorgu)))"
$qdosya = Join-Path $is ("Butce_" + $(if ($Kapanis) { $Kapanis } else { "son" }) + ".query.pq")
Set-Content -Path $qdosya -Value $sarili -Encoding UTF8

# Kimlik bilgisi: PQTest her veri kaynağını ayrı ister; her kaynak için tek kaynaklı küçük bir sorgu yazılır.
# Yerel dosya için Windows kimliği, kullanıcı adı/parola boş (geçerli kullanıcı).
$klasor = Split-Path -Parent $Ayar
$kaynaklar = @("File.Contents(""$Ayar"")", "Folder.Files(""$klasor\maliyet"")") +
    @(Get-ChildItem $klasor -Filter "*.xlsx" | Where-Object { $_.Name -ne (Split-Path -Leaf $Ayar) } |
        ForEach-Object { "File.Contents(""$($_.FullName)"")" })
$i = 0
foreach ($ifade in $kaynaklar) {
    $i++
    $kq = Join-Path $is "Kaynak$i.query.pq"
    Set-Content -Path $kq -Value $ifade -Encoding UTF8
    $sablon = (& $pqtest credential-template -e $mez -q $kq -p --authenticationKind Windows) -join "`n"
    $sablon = $sablon -replace '(?s)"AuthenticationProperties":\s*\{.*?\}', '"AuthenticationProperties": {}'
    $sonuc = ($sablon | & $pqtest set-credential -e $mez -q $kq -p) -join "`n"
    if ($sonuc -notmatch '"Status": "Success"') { throw "Kimlik bilgisi ayarlanamadı: $ifade`n$sonuc" }
}

$sure = Measure-Command { & $pqtest run-test -e $mez -q $qdosya -p | Out-File $Cikti -Encoding utf8 }
Write-Host ("PQTest süresi: {0:N1} sn" -f $sure.TotalSeconds)
