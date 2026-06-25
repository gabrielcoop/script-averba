$downloads   = "$env:USERPROFILE\Downloads"
$baseCooper  = "G:\.shortcut-targets-by-id"
$cacheFile   = "C:\XML_MDFE\cache_paths.json"
$logLocal    = "C:\XML_MDFE\log_local.txt"
$errorLocal  = "C:\XML_MDFE\error_local.txt"
$scriptName  = "monitor_xml.ps1"
$maquina     = $env:COMPUTERNAME

$modelos = @{
    "55" = "NF-e"
    "57" = "CT-e"
    "58" = "MDF-e"
    "65" = "NFC-e"
    "67" = "CT-e OS"
}

function Log($msg, $arquivo = $null) {
    $usuario = $env:USERNAME
    $linha   = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $maquina | $usuario | $msg"

    if ($arquivo) {
        $alvo = $arquivo
    } else {
        $alvo = $script:logFile
    }

    if (-not $alvo) {
        $alvo = $logLocal
    }

    try {
        Add-Content $alvo $linha
    } catch {
        Add-Content $logLocal $linha
    }
}

function LogErro($msg) {
    if ($script:errorLogFile) {
        $arquivo = $script:errorLogFile
    } else {
        $arquivo = $errorLocal
    }

    Log "[ERRO] $msg" $arquivo
}

function LogInicio {
    $usuario = $env:USERNAME
    $data    = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $sep     = "=================================================="

    if ($script:logFile) {
        $alvo = $script:logFile
    } else {
        $alvo = $logLocal
    }

    Add-Content $alvo ""
    Add-Content $alvo $sep
    Add-Content $alvo "INICIO DO SCRIPT | $data | $maquina | $usuario"
    Add-Content $alvo $sep
}

New-Item -ItemType Directory -Force -Path "C:\XML_MDFE" | Out-Null

$jaRodando = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*$scriptName*" }

if ($jaRodando.Count -gt 1) {
    Add-Content $logLocal "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $maquina | Instancia duplicada detectada - encerrando"
    exit
}

$driveProcess = "GoogleDriveFS"

function Get-GoogleDrivePath {
    $base = "$env:ProgramFiles\Google\Drive File Stream"

    if (-not (Test-Path $base)) {
        return $null
    }

    $versoes = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending

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

    if ($path) {
        Start-Process $path
        Log "Google Drive iniciado: $path"
    } else {
        Log "Executavel do Google Drive nao encontrado"
    }
}

function Stop-GoogleDrive {
    Log "Finalizando processo Google Drive"

    Get-Process $driveProcess -ErrorAction SilentlyContinue |
        Stop-Process -Force
}

function Garantir-GoogleDrive-Inicializacao {
    $tentativas = 0
    $inicio     = Get-Date

    Log "Aguardando Google Drive ficar disponivel (timeout 10min)"

    while ($true) {

        $processo = Get-Process $driveProcess -ErrorAction SilentlyContinue
        $driveOk  = (Test-Path "G:\") -and (Test-Path $baseCooper)

        if ($processo -and $driveOk) {
            Log "Google Drive OK"
            return $true
        }

        if (-not $processo) {
            Log "Processo nao encontrado, iniciando Drive"
            Start-GoogleDrive
        } else {
            Log "Processo existe mas unidade nao esta pronta"
        }

        Start-Sleep -Seconds 15
        $tentativas++

        if ((Test-Path "G:\") -and (Test-Path $baseCooper)) {
            Log "Drive montado apos aguardar"
            return $true
        }

        if ($tentativas -ge 8) {
            Log "Reiniciando Drive apos $tentativas tentativas sem resposta"

            Stop-GoogleDrive

            Start-Sleep -Seconds 5

            Start-GoogleDrive

            $tentativas = 0
        }

        if ((Get-Date) -gt $inicio.AddMinutes(10)) {
            Log "Timeout de inicializacao atingido (10 min) - continuando sem Drive"
            return $false
        }
    }
}

function Drive-Disponivel {
    return ((Test-Path "G:\") -and (Test-Path $baseCooper))
}

Log "==== MONITOR INICIADO ===="

try {
    Garantir-GoogleDrive-Inicializacao | Out-Null
} catch {
    Log "Erro durante inicializacao do Drive: $_"
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

            Log "Pastas carregadas do cache"
        }

    } catch {
        Log "Erro ao ler cache de pastas: $_"
    }
}

if (
    (-not $pastaPublico) -or
    (-not $pastaQuimico) -or
    (-not $pastaTI)
) {

    Log "Recriando cache de pastas"

    $nivel1 = Get-ChildItem $baseCooper -Directory -ErrorAction SilentlyContinue

    $pastaPublicoObj = $nivel1 | ForEach-Object {
        Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
    } | Where-Object {
        $_.Name -like "*blico"
    } | Select-Object -First 1

    if (-not $pastaPublicoObj) {
        Log "ERRO CRITICO: Pasta Publico nao localizada. Encerrando."
        exit 1
    }

    $pastaPublico = $pastaPublicoObj.FullName

    $pastaQuimicoObj = Get-ChildItem $pastaPublico -Directory |
        Where-Object { $_.Name -like "*10.*" } |
        Select-Object -First 1

    if (-not $pastaQuimicoObj) {
        Log "ERRO CRITICO: Pasta Quimico nao localizada. Encerrando."
        exit 1
    }

    $pastaQuimico = $pastaQuimicoObj.FullName

    $pastaTIObj = Get-ChildItem $pastaPublico -Directory |
        Where-Object { $_.Name -like "*16.*" } |
        Select-Object -First 1

    if (-not $pastaTIObj) {
        Log "ERRO CRITICO: Pasta TI nao localizada. Encerrando."
        exit 1
    }

    $pastaTI = $pastaTIObj.FullName

    @{
        publico = $pastaPublico
        quimico = $pastaQuimico
        ti      = $pastaTI
    } | ConvertTo-Json | Set-Content $cacheFile -Encoding UTF8
}

$pastaXML             = Join-Path $pastaQuimico "XML MDFe SEGURO RCV"
$pastaControle        = Join-Path $pastaTI "LOGS AVERBACAO"
$arquivoControleFinal = Join-Path $pastaControle "controle_mdfe.json"

New-Item -ItemType Directory -Force -Path $pastaXML | Out-Null
New-Item -ItemType Directory -Force -Path $pastaControle | Out-Null

$script:logFile      = Join-Path $pastaControle "monitor_xml_$maquina.log"
$script:errorLogFile = Join-Path $pastaControle "monitor_xml_error_$maquina.log"

LogInicio
Log "Logs redirecionados para o Drive"
Log "Pasta XML: $pastaXML"
Log "Pasta Controle: $pastaControle"

function Registrar-Averbacao($JsonTemp, $XmlNome, $Movido, $MotivoFalha) {

    $novoRegistro = [PSCustomObject]@{
        tipo           = $JsonTemp.tipo
        numero         = $JsonTemp.numero
        usuarioSistema = $JsonTemp.usuarioSistema
        dataEvento     = $JsonTemp.data
        xmlNome        = $XmlNome
    }

    if (-not (Drive-Disponivel)) {

        LogErro "Drive indisponivel ao registrar averbacao"

        $novoRegistro |
            ConvertTo-Json |
            Add-Content (Join-Path "C:\XML_MDFE" "averbacoes_pendentes.json") -Encoding UTF8

        return
    }

    try {

        if (Test-Path $arquivoControleFinal) {

            $conteudoRaw = Get-Content $arquivoControleFinal -Raw | ConvertFrom-Json
            $conteudo    = @($conteudoRaw)

        } else {

            $conteudo = @()
        }

        $conteudo += $novoRegistro

        $tempFile = "$arquivoControleFinal.tmp"

        $conteudo |
            ConvertTo-Json -Depth 5 |
            Set-Content $tempFile -Encoding UTF8

        Move-Item $tempFile $arquivoControleFinal -Force

        if ($Movido) {
            $status = "OK"
        } else {
            $status = "FALHA"
        }

        if ($XmlNome) {
            $xmlLog = $XmlNome
        } else {
            $xmlLog = "NAO ENCONTRADO"
        }

        Log "Averbacao registrada | $($JsonTemp.tipo) $($JsonTemp.numero) | XML: $xmlLog | Status: $status"

    } catch {
        LogErro "Falha ao gravar controle_mdfe.json: $_"
    }
}

function Processar-Controle($controleTemp) {

    $jsonString = Get-Content $controleTemp -Raw -ErrorAction SilentlyContinue

    if (-not $jsonString -or $jsonString.Trim() -eq "") {

        Log "controle_mdfe.json vazio - descartando"

        Remove-Item $controleTemp -Force -ErrorAction SilentlyContinue

        return
    }

    $jsonTemp = $null

    try {
        $jsonTemp = $jsonString | ConvertFrom-Json
    } catch {
    }

    if (
        (-not $jsonTemp) -or
        (
            ($jsonTemp.tipo -ne "MDF-e") -and
            ($jsonTemp.tipo -ne "CT-e")
        ) -or
        (-not $jsonTemp.numero)
    ){

        Log "controle_mdfe.json invalido ou de outro tipo - descartando"

        Remove-Item $controleTemp -Force -ErrorAction SilentlyContinue

        return
    }

    try {
        $dataEvento = [datetime]::Parse($jsonTemp.data)
    } catch {

        LogErro "Timestamp invalido no JSON ('$($jsonTemp.data)'): $_"

        Remove-Item $controleTemp -Force -ErrorAction SilentlyContinue

        return
    }

    Log "Aguardando XML para $($jsonTemp.tipo) $($jsonTemp.numero) (timeout 30s)"

    $xml          = $null
    $timeout      = (Get-Date).AddSeconds(30)
    $janelaInicio = $dataEvento.AddSeconds(-15)

    while ((Get-Date) -lt $timeout) {

        $candidatos = Get-ChildItem $downloads -Filter "*.xml" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.LastWriteTime -gt $janelaInicio
            } |
            Sort-Object LastWriteTime -Descending

        foreach ($candidato in $candidatos) {

            $chave = $candidato.BaseName

            if ($chave.Length -ne 44) {
                continue
            }

            $modelo = $chave.Substring(20, 2)

            if ($modelo -ne "58") {
                continue
            }

            $tamanho1 = $candidato.Length

            Start-Sleep -Milliseconds 300

            $itemAtual = Get-Item $candidato.FullName -ErrorAction SilentlyContinue

            if ($itemAtual) {
                $tamanho2 = $itemAtual.Length
            } else {
                $tamanho2 = $null
            }

            if (
                ($tamanho1 -eq $tamanho2) -and
                ($tamanho2 -gt 0)
            ) {

                $xml = $candidato
                break
            }
        }

        if ($xml) {
            break
        }

        Start-Sleep -Milliseconds 500
    }

    if (-not $xml) {

        LogErro "XML nao encontrado em 30s para $($jsonTemp.tipo) $($jsonTemp.numero)"

        Registrar-Averbacao `
            -JsonTemp $jsonTemp `
            -XmlNome $null `
            -Movido $false `
            -MotivoFalha "XML nao chegou em 30s"

        Remove-Item $controleTemp -Force -ErrorAction SilentlyContinue

        return
    }

    Log "XML encontrado: $($xml.Name)"

    $destino = Join-Path $pastaXML $xml.Name
    $movido  = $false
    $motivo  = $null

    try {

        if (Test-Path $destino) {

            LogErro "XML ja existe no destino: $($xml.Name)"

            $motivo = "XML ja existia no destino"

        } else {

            Move-Item $xml.FullName $destino -Force

            Log "XML movido para: $destino"

            $movido = $true
        }

    } catch {

        LogErro "Falha ao mover XML $($xml.Name): $_"

        $motivo = "Erro ao mover: $_"
    }

    Registrar-Averbacao `
        -JsonTemp $jsonTemp `
        -XmlNome $xml.Name `
        -Movido $movido `
        -MotivoFalha $motivo

    Remove-Item $controleTemp -Force -ErrorAction SilentlyContinue
}

Log "Loop principal iniciado"

while ($true) {

    if (-not (Drive-Disponivel)) {

        Log "Drive indisponivel - aguardando 30s"

        Start-Sleep -Seconds 30

        continue
    }

    $controleTemp = Join-Path $downloads "controle_mdfe.json"

    if (Test-Path $controleTemp) {

        Start-Sleep -Milliseconds 300

        $controleProcessando = Join-Path `
            $downloads `
            "controle_mdfe_proc_$(Get-Date -Format 'HHmmss_fff').json"

        try {

            Rename-Item `
                $controleTemp `
                $controleProcessando `
                -Force `
                -ErrorAction Stop

        } catch {

            Start-Sleep -Seconds 1

            continue
        }

        try {

            Processar-Controle $controleProcessando

        } catch {

            LogErro "Excecao nao tratada ao processar controle: $_"

            Remove-Item `
                $controleProcessando `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }

    Start-Sleep -Seconds 2
}