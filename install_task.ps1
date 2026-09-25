# Rejestruje (lub usuwa z -Uninstall) zadanie Harmonogramu zadań uruchamiające skrypt
# przy starcie komputera, niezależnie od tego, czy ktoś jest zalogowany.
# -User: konto z dostępem do dysku sieciowego, na którym zapisano sekret (--set-secret).
# Uruchom w PowerShell jako administrator.
param(
    [string]$User = "$env:USERDOMAIN\$env:USERNAME",
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$taskName = 'ClickUp Exchange Bridge'

if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Usunięto zadanie '$taskName'."
    return
}

$script = Join-Path $PSScriptRoot 'exchange_to_clickup.py'
$python = (& python -c "import sys; print(sys.executable)").Trim()
$pythonw = Join-Path (Split-Path $python) 'pythonw.exe'
if (-not (Test-Path $pythonw)) { throw "Nie znaleziono pythonw.exe obok $python" }

# Hasło konta Windows jest potrzebne, żeby zadanie miało dostęp do dysku sieciowego i Menedżera
# poświadczeń tego konta (tryb "Uruchom niezależnie od tego, czy użytkownik jest zalogowany").
$cred = Get-Credential -UserName $User -Message "Hasło konta, jako które ma działać skrypt"

$action = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`"" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -AtStartup
$trigger.Delay = 'PT1M'  # daje czas na start sieci po uruchomieniu komputera
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Limited -Force | Out-Null
Write-Host "Zarejestrowano zadanie '$taskName' ($pythonw `"$script`") jako $($cred.UserName)."
Write-Host "Uruchom teraz: Start-ScheduledTask -TaskName '$taskName'"
