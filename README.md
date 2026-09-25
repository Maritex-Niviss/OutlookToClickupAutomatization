# ClickUp Exchange Bridge

Skrypt łączy się z Exchange Online przez Microsoft Graph jako aplikacja zarejestrowana w Entra ID, więc Outlook nie jest potrzebny. Dla każdego nieprzeczytanego maila w skrzynce pośredniej:

1. zakłada kolejny folder `<root>\2026\DIQ26001 - Kontrahent` na dysku sieciowym jako kopię całego folderu wzorcowego (`template`),
2. przekazuje maila na adres Email-to-Task ClickUp, dopisując ścieżkę do nowego folderu,
3. oznacza maila jako przeczytanego.

## Co musi przygotować admin Microsoft 365 / Entra ID

1. **Rejestracja aplikacji**: [Entra admin center](https://entra.microsoft.com) → *App registrations* → *New registration*, np. „ClickUp Exchange Bridge” (tylko to konto organizacji, bez adresu przekierowania). Z zakładki *Overview* potrzebny jest **Application (client) ID**. W zakładce *Certificates & secrets* trzeba utworzyć **client secret**. Jego wartość jest widoczna tylko raz, a sekret wygasa najpóźniej po 24 miesiącach.
2. **Dostęp tylko do skrzynki pośredniej** przez *RBAC for Applications* w Exchange Online PowerShell. W tym wariancie **nie** dodaje się uprawnień Graph w Entra ID, bo dawałyby dostęp do wszystkich skrzynek:
   ```powershell
   Connect-ExchangeOnline
   # ObjectId = Object ID z Entra → Enterprise applications → ClickUp Exchange Bridge
   New-ServicePrincipal -AppId <client_id> -ObjectId <object_id> -DisplayName "ClickUp Exchange Bridge"
   New-ManagementScope -Name "ClickUp Bridge - skrzynka" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'clickupautomatization@maritex.eu'"
   New-ManagementRoleAssignment -App <client_id> -Role "Application Mail.ReadWrite" -CustomResourceScope "ClickUp Bridge - skrzynka"
   New-ManagementRoleAssignment -App <client_id> -Role "Application Mail.Send" -CustomResourceScope "ClickUp Bridge - skrzynka"
   # sprawdzenie: InScope = True dla skrzynki pośredniej, False dla każdej innej
   Test-ServicePrincipalAuthorization -Identity <client_id> -Resource clickupautomatization@maritex.eu
   ```
   Uprawnienia zaczynają działać po maksymalnie 30–60 minutach.

## Instalacja

Skrypt działa na dowolnym komputerze lub serwerze z Windows, który ma dostęp do internetu i do dysku sieciowego z projektami.

```powershell
pip install -r requirements.txt
# uzupełnij config.ini: client_id (od admina), mailbox, ścieżki, adres ClickUp
python exchange_to_clickup.py --set-secret   # wklej sekret aplikacji, trafi do Menedżera poświadczeń
python exchange_to_clickup.py --check        # test konfiguracji, niczego nie zmienia
python exchange_to_clickup.py --once         # jeden prawdziwy cykl, logi w konsoli

# jako administrator: autostart przy starcie komputera, jako bieżący użytkownik
.\install_task.ps1
Start-ScheduledTask -TaskName 'ClickUp Exchange Bridge'
```

`--set-secret` trzeba uruchomić jako ten sam użytkownik Windows, pod którym potem działa zadanie. Ten użytkownik musi też mieć dostęp do dysku sieciowego. Inne konto wskażesz przez `.\install_task.ps1 -User DOMENA\login`.

`--check` pokazuje, czy działa szablon i dysk, logowanie do Exchange Online i skrzynka, a także wypisuje nieprzeczytane maile i nadawców.

Odinstalowanie: `.\install_task.ps1 -Uninstall`.

Gdy sekret aplikacji wygaśnie, admin tworzy nowy, a Ty uruchamiasz ponownie `--set-secret`. Do tego czasu w `bridge.log` pojawiają się błędy logowania (`AADSTS7000222`). Po zmianie hasła konta Windows trzeba ponownie uruchomić `install_task.ps1`.

## Logi i stan

Pliki leżą w `%LOCALAPPDATA%\ClickUpExchangeBridge\` użytkownika, pod którym działa skrypt:

- `bridge.log` – przebieg pracy i błędy,
- `state.json` – maile przerwane w trakcie przetwarzania. Przy ponownej próbie skrypt używa tego samego folderu i nie wysyła maila do ClickUp drugi raz.

## Uwagi

- Maile przeczytane ręcznie w skrzynce pośredniej zostaną pominięte. Taki mail można przetworzyć, oznaczając go ponownie jako nieprzeczytany.
- Przetwarzane są tylko maile od nadawców z domen wymienionych w `allowed_domains` (`maritex.eu`, `maritex.com.pl`). Pozostałe skrypt oznacza jako przeczytane i pomija, a informację o tym zapisuje w `bridge.log`. Warto to uzupełnić regułą w Exchange.
- Jeśli wysłanie się nie powiedzie, mail zostaje nieprzeczytany i skrypt ponawia próbę w kolejnym cyklu. Gdy Exchange jest niedostępny, skrypt łączy się ponownie w następnym cyklu.
- Foldery trafiają do podfolderu roku w `root`, np. `\\serwer\display\Projekty\DIQ_Internal_quote\2026`. Numer to najwyższy istniejący `DIQ<rok>nnn` w tym podfolderze + 1. W nowym roku skrypt sam zakłada folder roku (np. `2027`) i zaczyna numerację od `001`, czyli od `DIQ27001`. Rok jest brany z daty otrzymania maila w czasie lokalnym.
- Nazwą kontrahenta jest temat maila bez tagów ClickUp (`<assign me>`, `<tag nazwa>`, `<due tomorrow>` itp.), bez prefiksów `RE:`/`PD:`/`FW:` i bez znaków niedozwolonych w nazwach folderów (`\ / : * ? " < > |`). Do ClickUp trafia temat bez `PD:`/`RE:`, ale razem z tagami.
- Zadanie jest przypisywane osobie, która przesłała maila na skrzynkę pośrednią: skrzynka dopisuje do tematu `<assign jej@adres>`, a `<assign me>` zamienia na ten adres. Adres tej osoby musi być adresem jej konta w ClickUp. Tę funkcję wyłącza ustawienie `assign_sender = no`.
- W `root` używaj ścieżki UNC (`\\serwer\udzial\...`), a nie litery dysku. Zadanie uruchomione przy starcie komputera nie ma zmapowanych dysków.
- Certyfikaty HTTPS są sprawdzane w magazynie certyfikatów Windows, więc skrypt działa też za firmowym proxy z inspekcją SSL.
- Kopie maili przekazanych do ClickUp trafiają do Elementów wysłanych skrzynki pośredniej.
