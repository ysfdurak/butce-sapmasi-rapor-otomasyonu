# PQTest ortak işlevleri: Power Query SDK araçlarını indirir, kimlik bilgilerini ayarlar, sorgu çalıştırır.
# Excel'deki Power Query ile aynı M motoru (Microsoft.PowerQuery.SdkTools › PQTest.exe). Yalnız Windows.
$ErrorActionPreference = "Stop"
$script:Is = Join-Path ([IO.Path]::GetTempPath()) "pqtest"
New-Item -ItemType Directory -Force $script:Is | Out-Null

function PQTest-Hazirla {
    # PQTest bir veri kaynağı uzantısı ister: boş Deneme.pq → Deneme.mez (zip).
    $sdk = Join-Path $script:Is "sdk"
    if (-not (Test-Path "$sdk/tools/PQTest.exe")) {
        Invoke-WebRequest "https://www.nuget.org/api/v2/package/Microsoft.PowerQuery.SdkTools" -OutFile "$script:Is/sdk.zip"
        Expand-Archive "$script:Is/sdk.zip" $sdk -Force
    }
    $script:PQTest = "$sdk/tools/PQTest.exe"
    $script:Mez = Join-Path $script:Is "Deneme.mez"
    if (-not (Test-Path $script:Mez)) {
        Compress-Archive -Path (Join-Path $PSScriptRoot "Deneme.pq") -DestinationPath "$script:Is/Deneme.zip" -Force
        Move-Item "$script:Is/Deneme.zip" $script:Mez -Force
    }
}

function PQTest-Kimlik([string[]]$Kaynaklar) {
    # Her veri kaynağı için tek kaynaklı küçük bir sorgu; yerel dosyada Windows kimliği (geçerli kullanıcı).
    $i = 0
    foreach ($ifade in $Kaynaklar) {
        $i++
        $kq = Join-Path $script:Is ("Kaynak_" + [Math]::Abs($ifade.GetHashCode()) + ".query.pq")
        Set-Content -Path $kq -Value $ifade -Encoding UTF8
        $sablon = (& $script:PQTest credential-template -e $script:Mez -q $kq -p --authenticationKind Windows) -join "`n"
        $sablon = $sablon -replace '(?s)"AuthenticationProperties":\s*\{.*?\}', '"AuthenticationProperties": {}'
        $sonuc = ($sablon | & $script:PQTest set-credential -e $script:Mez -q $kq -p) -join "`n"
        if ($sonuc -notmatch '"Status": "Success"') { throw "Kimlik bilgisi ayarlanamadı: $ifade`n$sonuc" }
    }
}

function PQTest-Degistir([string]$Metin, [string]$Ad, [string]$Deger) {
    # Yalnız satır başındaki `Ad = "..."` tanımını değiştirir (koşullardaki karşılaştırmalara dokunmaz).
    return $Metin -replace ('(?m)^(\s*)' + [regex]::Escape($Ad) + ' = "[^"]*"'), ('${1}' + $Ad + ' = "' + $Deger + '"')
}

function PQTest-DegistirDeger([string]$Metin, [string]$Ad, [string]$Deger) {
    # Tırnaksız değer (true/false): satır başındaki `Ad = deger` tanımı.
    return $Metin -replace ('(?m)^(\s*)' + [regex]::Escape($Ad) + ' = (true|false)'), ('${1}' + $Ad + ' = ' + $Deger)
}

function PQTest-Calistir([string]$Sorgu, [string]$Ad, [string]$Cikti) {
    # Sorgunun tablosunu JSON metni olarak döndürür ve PQTest çıktısını dosyaya yazar.
    $sarili = "let`r`n    Sorgu = " + $Sorgu + "`r`nin`r`n    Text.FromBinary(Json.FromValue({Table.ColumnNames(Sorgu)} & Table.ToRows(Sorgu)))"
    $qdosya = Join-Path $script:Is ($Ad + ".query.pq")
    Set-Content -Path $qdosya -Value $sarili -Encoding UTF8
    $sure = Measure-Command { & $script:PQTest run-test -e $script:Mez -q $qdosya -p | Out-File $Cikti -Encoding utf8 }
    Write-Host ("PQTest süresi: {0:N1} sn" -f $sure.TotalSeconds)
}
