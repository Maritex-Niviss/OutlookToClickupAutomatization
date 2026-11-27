# ClickUp Exchange Bridge

Skrypt łączy się z Exchange Online przez Microsoft Graph jako aplikacja zarejestrowana w Entra ID, więc Outlook nie jest potrzebny. Dla każdego nieprzeczytanego maila w skrzynce pośredniej:

1. odczytuje kategorię z pierwszego słowa tematu (`DIQ`, `RFQ`, `C` albo `R`) i zakłada folder na dysku sieciowym jako kopię folderu wzorcowego tej kategorii (`template`):
   - `DIQ Firma XYZ` / `RFQ Firma XYZ` → kolejny folder w roku, np. `<root>\2026\RFQ26001 - Firma XYZ`,
   - `C Lattice` / `R Lattice` → nowy projekt w folderze firmy, np. `<root>\Lattice\CTM2601`,
2. przekazuje maila na adres Email-to-Task ClickUp. Zadanie nazywa się jak folder (`DIQ26001 - Kontrahent`), a jego opisem jest treść maila bez stopek i bez nagłówków `Od:/Wysłane:/Do:`,
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
- Przetwarzane są tylko maile od nadawców z domen wymienionych w `allowed_domains` (sekcja `[mail]`) (`maritex.eu`, `maritex.com.pl`). Pozostałe skrypt oznacza jako przeczytane i pomija, a informację o tym zapisuje w `bridge.log`. Warto to uzupełnić regułą w Exchange.
- Jeśli wysłanie się nie powiedzie, mail zostaje nieprzeczytany i skrypt ponawia próbę w kolejnym cyklu. Gdy Exchange jest niedostępny, skrypt łączy się ponownie w następnym cyklu.
- Kategorie są zdefiniowane w sekcjach `[type DIQ]`, `[type RFQ]` i `[type C, R]` w `config.ini`. Każda ma własny układ (`layout`), `root`, `template` i listę ClickUp (`clickup_address`). Nową kategorię dodaje się kolejną sekcją `[type XYZ]`.
  - Słowo typu może mieć po sobie `:`, `-` lub `.` (`RFQ: Firma`, `RFQ - Firma`), a przed nim może stać `PD:`/`RE:`. Nie trafia ono do nazwy kontrahenta.
  - Wielkość liter, spacje (także twarde i niewidoczne) i separator po słowie typu nie mają znaczenia: `rfq:Firma`, `  Rfq - Firma` i `RFQTestfirma` to RFQ. Cyfra zaraz po słowie typu (`RFQ26022 - …`) oznacza numer folderu, a nie typ.
  - Maile bez DIQ/RFQ/C/R na początku tematu są **ignorowane**: nie powstaje folder ani zadanie. Nadawca dostaje ze skrzynki pośredniej krótką wiadomość „Nie przetworzono: …” z informacją, że brakuje kategorii. Na autoodpowiedzi („Poza biurem”) skrypt nie odpowiada. Zachowanie zmienia ustawienie `default_type` w sekcji `[mail]` (np. `default_type = DIQ`).
  - Każda kategoria trafia na listę ClickUp z własnego `clickup_address`.
- **Trasy według nadawcy** (`[route NAZWA]` w `config.ini`): zadania od wymienionych nadawców (`senders`, pełne adresy) trafiają na listę z `clickup_address` trasy zamiast na listę kategorii, i to dla wszystkich kategorii. Folder na dysku powstaje jak zwykle. Przy kilku trasach wygrywa pierwsza pasująca. Z `create_folders = no` trasa tworzy **tylko zadanie w ClickUp**: nic nie powstaje na dysku, kategoria w temacie nie jest wymagana, a zadanie nazywa się jak temat maila (bez `PD:`/`RE:`). `--check` pokazuje trasy i przy każdym nieprzeczytanym mailu informuje, czy trafi na trasę.
- **Projekty C/R** (`layout = company`):
  - Po `C`/`R` nie trzeba spacji: `RRadwag`, `R Radwag` i `r:Radwag` dają R z firmą Radwag. Uwaga: każdy temat zaczynający się od C lub R (np. `Cześć`, `Radwag` bez kategorii) zostanie potraktowany jako projekt (`C` + „ześć”, `R` + „adwag”). RFQ ma pierwszeństwo przed R.
  - Folder firmy to nazwa z tematu bez spacji (`R Base System` → `BaseSystem`). Jeśli folder tej firmy już istnieje, skrypt wchodzi do niego i zakłada tylko nowy projekt. Przy porównywaniu nazw wielkość liter, spacje, `-` i `_` nie mają znaczenia: `Base System`, `BaseSystem` i `basesystem` to ta sama firma.
  - Kod projektu: `C`/`R`, potem pierwsza i ostatnia litera nazwy firmy przesunięte szyfrem Cezara o 8 (`Lattice`: L→T, e→M), dwie ostatnie cyfry roku i numer w roku: `CTM2601`, `RTM2601`, `CTM2602`… Numeracja jest liczona osobno dla C i R, w obrębie folderu firmy.
  - Zadanie w ClickUp nazywa się `CTM2601 - Lattice`. Mail z samym `C` albo `R`, bez nazwy firmy, jest odrzucany i nadawca dostaje informację.
- Foldery trafiają do podfolderu roku w `root` danego typu, np. `\\serwer\display\Projekty\RFQ\2026`. Numer to najwyższy istniejący `RFQ<rok>nnn` (albo `DIQ<rok>nnn`) w tym podfolderze + 1. W nowym roku skrypt sam zakłada folder roku (np. `2027`) i zaczyna numerację od `001`, czyli od `DIQ27001` i `RFQ27001`. Rok jest brany z daty otrzymania maila w czasie lokalnym.
- Nazwą kontrahenta jest temat maila bez tagów ClickUp (`<assign me>`, `<tag nazwa>`, `<due tomorrow>` itp.), bez prefiksów `RE:`/`PD:`/`FW:` i bez znaków niedozwolonych w nazwach folderów (`\ / : * ? " < > |`). Nazwą zadania w ClickUp jest nazwa folderu, a tagi z tematu są dopisywane na końcu, żeby ClickUp mógł je odczytać (ustawienie `subject` w `config.ini`).
- Treść zadania to treść maila w kilku krokach:
  - Stopka jest wycinana od linii zaczynającej się od „Pozdrawiam”, „Z poważaniem”, „Best regards” itp. (lista w `signature_markers`) aż do następnej cytowanej wiadomości.
  - Nagłówki cytowanych wiadomości (`Od:`/`Wysłane:`/`Do:`/`Temat:`) są usuwane, a ich treść zostaje.
  - Obrazki ze stopek (logo) są usuwane. Zwykłe załączniki zostają.
  - Pełny, nieobcięty mail jest zawsze dostępny w skrzynce pośredniej.
- Zadanie jest przypisywane osobie, która przesłała maila na skrzynkę pośrednią: skrzynka dopisuje do tematu `<assign jej@adres>`, a `<assign me>` zamienia na ten adres. Adres tej osoby musi być adresem jej konta w ClickUp. Tę funkcję wyłącza ustawienie `assign_sender = no`.
- W `root` i `template` używaj ścieżki UNC (`\\serwer\udzial\...`), a nie litery dysku. Zadanie uruchomione przy starcie komputera nie ma zmapowanych dysków.
- Certyfikaty HTTPS są sprawdzane w magazynie certyfikatów Windows, więc skrypt działa też za firmowym proxy z inspekcją SSL.
- Kopie maili przekazanych do ClickUp trafiają do Elementów wysłanych skrzynki pośredniej.
