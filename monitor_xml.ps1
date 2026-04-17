$downloads = "$env:USERPROFILE\Downloads"
$baseCooper = "G:\.shortcut-targets-by-id"
$cacheFile = "C:\XML_MDFE\cache_paths.json"

New-Item -ItemType Directory -Force -Path "C:\XML_MDFE" | Out-Null

while (-not (Test-Path $baseCooper)) {
    Start-Sleep 5
}

$pastaPublico = $null
$pastaQuimico = $null
$pastaTI      = $null

if (Test-Path $cacheFile) {
    try {
        $cache = Get-Content $cacheFile -Raw | ConvertFrom-Json

        if (
            (Test-Path $cache.publico) -and
            (Test-Path $cache.quimico) -and
            (Test-Path $cache.ti)
        ) {
            $pastaPublico = $cache.publico
            $pastaQuimico = $cache.quimico
            $pastaTI      = $cache.ti
        }
    }
    catch {

    }
}

if (-not $pastaPublico -or -not $pastaQuimico -or -not $pastaTI) {
    $nivel1 = Get-ChildItem $baseCooper -Directory -ErrorAction SilentlyContinue

    $pastaPublicoObj = $nivel1 | ForEach-Object {
        Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
    } | Where-Object { $_.Name -like "*blico" } | Select-Object -First 1

    if (-not $pastaPublicoObj) {
        exit
    }

    $pastaPublico = $pastaPublicoObj.FullName

    $pastaQuimicoObj = Get-ChildItem $pastaPublico -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*10.*" } |
        Select-Object -First 1

    if (-not $pastaQuimicoObj) {
        exit
    }

    $pastaQuimico = $pastaQuimicoObj.FullName

    $pastaTIObj = Get-ChildItem $pastaPublico -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*16.*" } |
        Select-Object -First 1

    if (-not $pastaTIObj) {
        exit
    }

    $pastaTI = $pastaTIObj.FullName

    @{
        publico = $pastaPublico
        quimico = $pastaQuimico
        ti      = $pastaTI
    } | ConvertTo-Json | Set-Content $cacheFile -Encoding UTF8

}

$pastaXML            = Join-Path $pastaQuimico "XML MDFe SEGURO RCV"
$pastaControle       = Join-Path $pastaTI "LOGS AVERBACAO"
$arquivoControleFinal = Join-Path $pastaControle "controle_mdfe.json"

New-Item -ItemType Directory -Force -Path $pastaXML      | Out-Null
New-Item -ItemType Directory -Force -Path $pastaControle | Out-Null

while ($true) {

    $controleTemp = Join-Path $downloads "controle_mdfe.json"

    if (Test-Path $controleTemp) {

        try {
            Start-Sleep -Milliseconds 300

            $jsonString = Get-Content $controleTemp -Raw

            if (-not $jsonString -or $jsonString.Trim() -eq "") {
                Remove-Item $controleTemp -Force
                continue
            }

            $jsonTemp = $jsonString | ConvertFrom-Json

            if ($jsonTemp.tipo -ne "MDFe") {
                Remove-Item $controleTemp -Force
                continue
            }

            if (-not $jsonTemp.numeroCte) {
                Remove-Item $controleTemp -Force
                continue
            }

            Start-Sleep -Seconds 1

            $xml = Get-ChildItem $downloads -Filter *.xml |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($xml) {

                $destinoXML = Join-Path $pastaXML $xml.Name

                if (Test-Path $destinoXML) {
                    $destinoXML = Join-Path $pastaXML ("dup_" + $xml.Name)
                }

                Move-Item $xml.FullName $destinoXML -ErrorAction Stop

            }

            $novoRegistro = [PSCustomObject]@{
                tipo           = $jsonTemp.tipo
                numeroCte      = $jsonTemp.numeroCte
                id             = $jsonTemp.id
                usuarioSistema = $jsonTemp.usuarioSistema
                data           = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }

            if (Test-Path $arquivoControleFinal) {
                $conteudo = Get-Content $arquivoControleFinal -Raw | ConvertFrom-Json

                if ($conteudo -isnot [System.Collections.IEnumerable]) {
                    $conteudo = @($conteudo)
                }
            } else {
                $conteudo = @()
            }

            $conteudo += $novoRegistro

            $conteudo | ConvertTo-Json -Depth 5 | Set-Content $arquivoControleFinal -Encoding UTF8

            Remove-Item $controleTemp -Force

        } catch {
            if (Test-Path $controleTemp) {
                Remove-Item $controleTemp -Force
            }
        }
    }

    Start-Sleep 2
}