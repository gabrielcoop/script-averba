$downloads = "$env:USERPROFILE\Downloads"
$baseCooper = "G:\.shortcut-targets-by-id"

while (-not (Test-Path $baseCooper)) {
    Write-Host "Aguardando Google Drive montar..."
    Start-Sleep 10
}

Write-Host "Google Drive encontrado!"

$pastaQuimico = Get-ChildItem $baseCooper -Recurse -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "*mico*" } |
    Select-Object -First 1

if (-not $pastaQuimico) {
    Write-Host "Pasta Químico não encontrada!"
    exit
}

$pastaQuimico = $pastaQuimico.FullName
Write-Host "Pasta Químico:" $pastaQuimico

$pastaTI = Get-ChildItem $baseCooper -Recurse -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "*T.I*" } |
    Select-Object -First 1

if (-not $pastaTI) {
    Write-Host "Pasta TI não encontrada!"
    exit
}

$pastaTI = $pastaTI.FullName
Write-Host "Pasta TI:" $pastaTI

$pastaXML = Join-Path $pastaQuimico "XML MDFe SEGURO RCV"
$pastaControle = Join-Path $pastaTI "LOGS AVERBACAO"

$arquivoControleFinal = Join-Path $pastaControle "controle_mdfe.json"

New-Item -ItemType Directory -Force -Path $pastaXML | Out-Null
New-Item -ItemType Directory -Force -Path $pastaControle | Out-Null

Write-Host "Destino XML:" $pastaXML
Write-Host "Destino Controle:" $pastaControle

while ($true) {

    $controleTemp = Join-Path $downloads "controle_mdfe.json"

    if (Test-Path $controleTemp) {

        try {
            Start-Sleep -Milliseconds 300

            $jsonString = Get-Content $controleTemp -Raw

            if (-not $jsonString -or $jsonString.Trim() -eq "") {
                Write-Host "JSON vazio, ignorando..."
                Remove-Item $controleTemp -Force
                continue
            }

            $jsonTemp = $jsonString | ConvertFrom-Json

            if (-not $jsonTemp.tipo -or $jsonTemp.tipo -ne "MDFe") {
                Write-Host "Tipo inválido, ignorando..."
                Remove-Item $controleTemp -Force
                continue
            }

            if (-not $jsonTemp.numero) {
                Write-Host "JSON sem número, ignorando..."
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
                Write-Host "XML movido:" $xml.Name

            } else {
                Write-Host "Nenhum XML encontrado"
            }

            $novoRegistro = [PSCustomObject]@{
                tipo   = $jsonTemp.tipo
                numeroCte = $jsonTemp.numeroCte
                id     = $jsonTemp.id
                usuarioSistema = $jsonTemp.usuarioSistema
                data   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }

            if (Test-Path $arquivoControleFinal) {
                $conteudo = Get-Content $arquivoControleFinal -Raw | ConvertFrom-Json

                if ($conteudo -isnot [System.Collections.IEnumerable]) {
                    $conteudo = @($conteudo)
                }
            } else {
                $conteudo = @()
            }

            $conteudo = @($conteudo)
            $conteudo += $novoRegistro

            $conteudo | ConvertTo-Json -Depth 5 | Set-Content $arquivoControleFinal -Encoding UTF8

            Write-Host "Registro adicionado ao histórico"

            Remove-Item $controleTemp -Force

        } catch {
            Write-Host "Erro:" $_

            if (Test-Path $controleTemp) {
                Remove-Item $controleTemp -Force
            }
        }
    }

    Start-Sleep 2
}