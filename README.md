# ClickUp Outlook Bridge

Skrypt w tle podpina się pod działający klasyczny Outlook. Dla każdego nieprzeczytanego maila w skrzynce pośredniej:

1. zakłada kolejny folder `DIQ26001 - Kontrahent` na dysku sieciowym jako kopię całego folderu wzorcowego (`template`),
2. przekazuje maila na adres Email-to-Task ClickUp, dopisując ścieżkę do nowego folderu,
3. oznacza maila jako przeczytanego.

## Instalacja

```powershell
pip install -r requirements.txt
# uzupełnij config.ini (skrzynka, ścieżki, adres ClickUp)
python outlook_to_clickup.pyw --check   # test konfiguracji, niczego nie zmienia
python outlook_to_clickup.pyw --once    # jeden prawdziwy cykl, logi w konsoli
.\install_task.ps1                      # autostart przy logowaniu (bez okna)
Start-ScheduledTask -TaskName 'ClickUp Outlook Bridge'
```

Odinstalowanie: `.\install_task.ps1 -Uninstall`.

## Logi i stan

`%LOCALAPPDATA%\ClickUpOutlookBridge\`:

- `bridge.log` – przebieg pracy i błędy,
- `state.json` – maile przerwane w trakcie przetwarzania. Przy ponownej próbie skrzynka używa tego samego folderu i nie wysyła maila do ClickUp drugi raz.

## Uwagi

- Maile przeczytane ręcznie w skrzynce pośredniej zostaną pominięte. Taki mail można przetworzyć, oznaczając go ponownie jako nieprzeczytany.
- Przetwarzane są tylko maile od nadawców z domen wymienionych w `allowed_domains` (`maritex.eu`, `maritex.com.pl`). Pozostałe skrypt oznacza jako przeczytane i pomija, a informację o tym zapisuje w `bridge.log`. Warto to uzupełnić regułą w Exchange.
- Jeśli wysłanie się nie powiedzie, mail zostaje nieprzeczytany i skrypt ponawia próbę w kolejnym cyklu.
- Numer to najwyższy istniejący `DIQ<rok>nnn` z danego roku + 1. W nowym roku numeracja zaczyna się od `001`, np. `DIQ27001`. Rok jest brany z daty otrzymania maila.
- Nazwą kontrahenta jest temat maila bez tagów ClickUp (`<assign me>`, `<tag nazwa>`, `<due tomorrow>` itp.), bez prefiksów `RE:`/`PD:`/`FW:` i bez znaków niedozwolonych w nazwach folderów (`\ / : * ? " < > |`). Do ClickUp trafia temat bez `PD:`/`RE:`, ale razem z tagami.
- Zadanie jest przypisywane osobie, która przesłała maila na skrzynkę pośrednią: skrzynka dopisuje do tematu `<assign jej@adres>`, a `<assign me>` zamienia na ten adres. Adres tej osoby musi być adresem jej konta w ClickUp. Tę funkcję wyłącza ustawienie `assign_sender = no`.
- W `root` używaj ścieżki UNC (`\\serwer\udzial\...`), a nie litery dysku. Zmapowane litery bywają niedostępne zaraz po zalogowaniu.
- Outlook i skrypt muszą działać na tym samym poziomie uprawnień, czyli żaden z nich nie może być uruchomiony „jako administrator”.
- Bez aktualnego antywirusa Outlook może przy każdej wysyłce pytać, czy zezwolić programowi na wysłanie wiadomości. Ustawienie znajduje się w Centrum zaufania → Dostęp programowy.
