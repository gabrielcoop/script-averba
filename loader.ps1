$jaRodando = Get-CimInstance Win32_Process |
Where-Object { $_.CommandLine -like "*monitor_xml.ps1*" }

if ($jaRodando) {
    Write-Host "Monitor já está rodando"
}

$scriptUrl = "https://raw.githubusercontent.com/gabrielcoop/script-averba/master/monitor_xml.ps1"

$pastaLocal = "C:\XML_MDFE"
$scriptLocal = Join-Path $pastaLocal "monitor_xml.ps1"

New-Item -ItemType Directory -Force -Path $pastaLocal | Out-Null

try {
    Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptLocal -UseBasicParsing
    Write-Host "Script atualizado com sucesso!"
}
catch {
    Write-Host "Erro ao baixar script remoto: $($_.Exception.Message)"
}

Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptLocal`"" -WindowStyle Hidden