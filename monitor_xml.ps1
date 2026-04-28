$downloads = "$env:USERPROFILE\Downloads"
$baseCooper = "G:\.shortcut-targets-by-id"
$cacheFile = "C:\XML_MDFE\cache_paths.json"
$logFile = "C:\XML_MDFE\log.txt"
$scriptName = "monitor_xml.ps1"
$maquina = $env:COMPUTERNAME

function Log($msg) {
    $usuario = $env:USERNAME
    $linha = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $maquina | $usuario | $msg"
    Add-Content $logFile $linha
}

$jaRodando = Get-CimInstance Win32_Process |
Where-Object { $_.CommandLine -like "*$scriptName*" }

if ($jaRodando.Count -gt 1) {
    Log "Já existe outra instância rodando"
    exit
}

New-Item -ItemType Directory -Force -Path "C:\XML_MDFE" | Out-Null

$driveProcess = "GoogleDriveFS"

function Get-GoogleDrivePath {

    Log "Buscando Google Drive"

    $basePath = "$env:ProgramFiles\Google\Drive File Stream"

    if (-not (Test-Path $basePath)) {
        return $null
    }

    $versoes = Get-ChildItem $basePath -Directory | Sort-Object Name -Descending

    foreach ($v in $versoes) {
        $exe = Join-Path $v.FullName "GoogleDriveFS.exe"
        if (Test-Path $exe) {
            return $exe
        }
    }

    return $null
}

function Start-GoogleDrive {

    Log "Tentando iniciar Google Drive"

    $path = Get-GoogleDrivePath

    if ($path -and (Test-Path $path)) {
        Start-Process $path
        Log "Iniciado: $path"
    } else {
        Log "Google Drive não encontrado"
    }
}

function Stop-GoogleDrive {
    Log "Finalizando Google Drive"
    Get-Process $driveProcess -ErrorAction SilentlyContinue | Stop-Process -Force
}

function Garantir-GoogleDrive {

    $tentativas = 0
    $inicio = Get-Date

    while ($true) {

        $processo = Get-Process $driveProcess -ErrorAction SilentlyContinue

        $driveOk = (Test-Path "G:\") -and (Test-Path $baseCooper)

        if ($processo -and $driveOk) {
            Log "Google Drive OK (processo + unidade OK)"
            return
        }

        if (-not $processo) {
            Log "Processo não encontrado"
            Start-GoogleDrive
        } else {
            Log "Processo existe, mas unidade não está pronta"
        }

        Start-Sleep -Seconds 15
        $tentativas++

        $driveOk = (Test-Path "G:\") -and (Test-Path $baseCooper)

        if ($driveOk) {
            Log "Drive montado com sucesso"
            return
        }

        if ($tentativas -ge 8) {
            Log "Drive não respondeu após várias tentativas, reiniciando..."
            Stop-GoogleDrive
            Start-Sleep 5
            Start-GoogleDrive
            $tentativas = 0
        }

        if ((Get-Date) -gt $inicio.AddMinutes(10)) {
            Log "Timeout geral atingido (10 minutos)"
            break
        }
    }
}

Log "==== INICIO DO SCRIPT ===="

try {
    Garantir-GoogleDrive
    Log "Inicialização concluída"
} catch {
    Log "Erro geral: $_"
}

Log "==== FIM DO SCRIPT ===="

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
    } catch {
        Log "Erro ao ler cache"
    }
}

if (-not $pastaPublico -or -not $pastaQuimico -or -not $pastaTI) {

    Log "Recriando cache de pastas"

    $nivel1 = Get-ChildItem $baseCooper -Directory -ErrorAction SilentlyContinue

    $pastaPublicoObj = $nivel1 | ForEach-Object {
        Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
    } | Where-Object { $_.Name -like "*blico" } | Select-Object -First 1

    if (-not $pastaPublicoObj) {
        Log "Erro ao localizar pasta Público"
        exit
    }

    $pastaPublico = $pastaPublicoObj.FullName

    $pastaQuimicoObj = Get-ChildItem $pastaPublico -Directory |
        Where-Object { $_.Name -like "*10.*" } |
        Select-Object -First 1

    if (-not $pastaQuimicoObj) {
        Log "Erro ao localizar pasta Químico"
        exit
    }

    $pastaQuimico = $pastaQuimicoObj.FullName

    $pastaTIObj = Get-ChildItem $pastaPublico -Directory |
        Where-Object { $_.Name -like "*16.*" } |
        Select-Object -First 1

    if (-not $pastaTIObj) {
        Log "Erro ao localizar pasta TI"
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

$logNuvem = $pastaControle

if (-not (Test-Path $logNuvem)) {
    New-Item -ItemType Directory -Force -Path $logNuvem | Out-Null
}

$logFile = Join-Path $logNuvem "monitor_xml_$maquina.log"

Log "Log agora sendo salvo na nuvem"

New-Item -ItemType Directory -Force -Path $pastaXML      | Out-Null
New-Item -ItemType Directory -Force -Path $pastaControle | Out-Null

while ($true) {
    if (-not (Test-Path $baseCooper)) {
        Log "Drive caiu, tentando recuperar"
        Garantir-GoogleDrive
    }

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

            if ($jsonTemp.tipo -ne "MDFe" -or -not $jsonTemp.numeroCte) {
                Remove-Item $controleTemp -Force
                continue
            }

            Start-Sleep -Seconds 1

            $agora = Get-Date

            $xml = Get-ChildItem $downloads -Filter *.xml |
                Where-Object { $_.LastWriteTime -gt $agora.AddMinutes(-2) } |
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
                dataEvento     = $jsonTemp.data
                dataProcessado = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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

            $tempFile = "$arquivoControleFinal.tmp"

            $conteudo | ConvertTo-Json -Depth 5 | Set-Content $tempFile -Encoding UTF8
            Move-Item $tempFile $arquivoControleFinal -Force

            Remove-Item $controleTemp -Force

            Log "Processado CTe $($jsonTemp.numeroCte) | XML: $($xml.name)"

        } catch {
            Log "Erro ao processar JSON: $_"
            if (Test-Path $controleTemp) {
                Remove-Item $controleTemp -Force
            }
        }
    }

    Start-Sleep 2
}