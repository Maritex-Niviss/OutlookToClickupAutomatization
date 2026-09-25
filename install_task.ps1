# Rejestruje (lub usuwa z -Uninstall) zadanie Harmonogramu zadań uruchamiające
# skrypt przy logowaniu bieżącego użytkownika, bez okna konsoli.
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$taskName = 'ClickUp Outlook Bridge'

if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Usunięto zadanie '$taskName'."
    return
}

$script = Join-Path $PSScriptRoot 'outlook_to_clickup.pyw'
$python = (& python -c "import sys; print(sys.executable)").Trim()
$pythonw = Join-Path (Split-Path $python) 'pythonw.exe'
if (-not (Test-Path $pythonw)) { throw "Nie znaleziono pythonw.exe obok $python" }

$user = "$env:USERDOMAIN\$env:USERNAME"
$action = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`"" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$trigger.Delay = 'PT1M'  # daje czas na podłączenie dysków sieciowych i start Outlooka
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
# Limited (bez podniesionych uprawnień) - musi pasować do Outlooka, inaczej COM się nie podłączy.
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Principal $principal -Force | Out-Null
Write-Host "Zarejestrowano zadanie '$taskName' ($pythonw `"$script`")."
Write-Host "Uruchom teraz: Start-ScheduledTask -TaskName '$taskName'"
