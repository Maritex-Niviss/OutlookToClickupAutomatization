"""Most Exchange Online -> dysk sieciowy -> ClickUp.

Łączy się z Exchange Online przez Microsoft Graph (bez Outlooka), jako aplikacja
zarejestrowana w Entra ID, i co `poll_interval` sekund sprawdza nieprzeczytane
maile w skrzynce pośredniej. Dla każdego maila:
  1. zakłada kolejny folder "<root>\\2026\\DIQ26001 - Kontrahent" na dysku sieciowym
     jako kopię folderu szablonu (numeracja od 001 w każdym roku, folder roku
     zakładany automatycznie),
  2. przekazuje maila (Forward) na adres Email-to-Task listy ClickUp,
     dopisując ścieżkę do folderu,
  3. oznacza maila jako przeczytanego.

Uruchomienie ręczne (z konsolą, do testów):
    python exchange_to_clickup.py --set-secret   # zapisuje sekret aplikacji w Menedżerze poświadczeń
    python exchange_to_clickup.py --check        # sprawdza konfigurację, nic nie zmienia
    python exchange_to_clickup.py --once         # jeden cykl przetwarzania i koniec
"""

import configparser
import getpass
import html
import json
import logging
import logging.handlers
import msvcrt
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import keyring
import msal
import requests
import truststore

truststore.inject_into_ssl()  # certyfikaty z magazynu Windows (np. firmowy proxy z inspekcją SSL)

APP_NAME = "ClickUpExchangeBridge"
GRAPH_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.ini"
DATA_DIR = Path(os.environ.get("LOCALAPPDATA", BASE_DIR)) / APP_NAME
STATE_PATH = DATA_DIR / "state.json"
LOG_PATH = DATA_DIR / "bridge.log"
LOCK_PATH = DATA_DIR / "bridge.lock"

CLICKUP_TAG = re.compile(r"<[^<>]*>")  # <assign me>, <tag pilne> itp. w temacie
REPLY_PREFIX = re.compile(r"^(\s*(re|fw|fwd|pd|odp|wg)\s*:)+", re.IGNORECASE)
INVALID_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')  # niedozwolone w nazwach folderów Windows
ASSIGN_ME = re.compile(r"<\s*assign[\s.]+me\s*>", re.IGNORECASE)

log = logging.getLogger(APP_NAME)


class Config:
    def __init__(self, path):
        cp = configparser.ConfigParser(interpolation=None)
        if not cp.read(path, encoding="utf-8"):
            raise FileNotFoundError(f"Brak pliku konfiguracyjnego: {path}")
        self.tenant_id = cp.get("exchange", "tenant_id").strip()
        self.client_id = cp.get("exchange", "client_id").strip()
        self.mailbox = cp.get("exchange", "mailbox").strip()
        self.subfolder = cp.get("exchange", "subfolder", fallback="").strip()
        self.poll_interval = cp.getint("exchange", "poll_interval", fallback=60)
        self.allowed_domains = {
            d.strip().lstrip("@").casefold()
            for d in cp.get("exchange", "allowed_domains", fallback="").split(",")
            if d.strip()
        }
        self.docs_root = Path(cp.get("documents", "root").strip())
        self.prefix = cp.get("documents", "prefix", fallback="DIQ").strip()
        self.number_width = cp.getint("documents", "number_width", fallback=3)
        self.template = Path(cp.get("documents", "template").strip())
        self.clickup_address = cp.get("clickup", "address").strip()
        self.subject_format = cp.get("clickup", "subject", fallback="{subject}").strip()
        self.assign_sender = cp.getboolean("clickup", "assign_sender", fallback=True)


# --- Dysk sieciowy -----------------------------------------------------------

def client_name(subject):
    """Nazwa kontrahenta z tematu: bez tagów ClickUp, prefiksów RE:/PD: i znaków niedozwolonych."""
    name = CLICKUP_TAG.sub(" ", subject or "")
    name = REPLY_PREFIX.sub("", name)
    name = INVALID_CHARS.sub(" ", name)
    return " ".join(name.split())[:100].rstrip(" .")


def year_folder(root, year):
    """Podfolder roku, np. ...\\DIQ_Internal_quote\\2026."""
    return root / str(year)


def next_folder_id(root, prefix, width, year):
    """DIQ26001, DIQ26002, ... - numeracja od początku w każdym roku (w podfolderze roku)."""
    yy = f"{year % 100:02d}"
    pattern = re.compile(rf"^{re.escape(prefix)}{yy}(\d+)(?!\d)", re.IGNORECASE)
    year_dir = year_folder(root, year)
    numbers = [
        int(m.group(1))
        for entry in (os.scandir(year_dir) if year_dir.is_dir() else ())
        if entry.is_dir() and (m := pattern.match(entry.name))
    ]
    return f"{prefix}{yy}{max(numbers, default=0) + 1:0{width}d}"


def create_next_folder(root, prefix, width, year, name):
    year_dir = year_folder(root, year)
    year_dir.mkdir(exist_ok=True)  # w nowym roku zakłada np. 2027
    for _ in range(20):
        folder_id = next_folder_id(root, prefix, width, year)
        path = year_dir / (f"{folder_id} - {name}" if name else folder_id)
        try:
            path.mkdir()
            return path
        except FileExistsError:
            continue  # ktoś równolegle założył ten numer - liczymy od nowa
    raise RuntimeError(f"Nie udało się założyć nowego folderu {prefix} w {root}")


def _copy_if_missing(src, dst):
    # Przy wznowieniu nie nadpisujemy plików już skopiowanych (lub zmienionych).
    if not os.path.exists(dst):
        shutil.copy2(src, dst)


def fill_from_template(template, pr_dir):
    shutil.copytree(template, pr_dir, dirs_exist_ok=True, copy_function=_copy_if_missing)


# --- Stan (ochrona przed duplikatami po awarii w połowie cyklu) --------------

def load_state():
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, ValueError):
        log.exception("Nie można odczytać %s - zaczynam z pustym stanem", STATE_PATH)
        return {}


def save_state(state):
    tmp = STATE_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, STATE_PATH)


# --- Exchange Online (Microsoft Graph) ---------------------------------------

class GraphError(RuntimeError):
    pass


class Graph:
    """Skrzynka pośrednia w Exchange Online, dostęp jako aplikacja z Entra ID."""

    def __init__(self, cfg):
        secret = keyring.get_password(APP_NAME, cfg.client_id)
        if secret is None:
            raise RuntimeError(
                "Brak sekretu aplikacji - uruchom jako użytkownik, pod którym działa skrypt: "
                "python exchange_to_clickup.py --set-secret"
            )
        self.app = msal.ConfidentialClientApplication(
            cfg.client_id,
            authority=f"https://login.microsoftonline.com/{cfg.tenant_id}",
            client_credential=secret,
        )
        self.session = requests.Session()
        self.base = f"{GRAPH_URL}/users/{quote(cfg.mailbox)}"

    def request(self, method, path, **kwargs):
        token = self.app.acquire_token_for_client(scopes=GRAPH_SCOPE)  # MSAL trzyma token do wygaśnięcia
        if "access_token" not in token:
            raise GraphError(f"Logowanie aplikacji do Entra ID nieudane: {token.get('error_description')}")
        url = path if path.startswith("https://") else self.base + path
        response = self.session.request(
            method, url, headers={"Authorization": f"Bearer {token['access_token']}"}, timeout=60, **kwargs
        )
        if not response.ok:
            raise GraphError(f"{method} {url} -> {response.status_code}: {response.text[:500]}")
        return response.json() if response.content else None


def find_inbox(graph, subfolder):
    folder = graph.request("GET", "/mailFolders/inbox")
    for part in filter(None, re.split(r"[\\/]", subfolder)):
        children = graph.request("GET", f"/mailFolders/{folder['id']}/childFolders", params={"$top": 200})
        folder = next(
            (f for f in children["value"] if f["displayName"].casefold() == part.casefold()), None
        )
        if folder is None:
            raise LookupError(f"Brak podfolderu '{part}' w skrzynce")
    return folder


def fetch_unread(graph, folder):
    # Graph wymaga, żeby pole z $orderby było też pierwsze w $filter.
    params = {
        "$filter": "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false",
        "$orderby": "receivedDateTime",
        "$select": "id,subject,sender,from,receivedDateTime",
        "$top": 50,
    }
    url, mails = f"/mailFolders/{folder['id']}/messages", []
    while url:
        page = graph.request("GET", url, params=params)
        params = None  # nextLink zawiera już wszystkie parametry
        # Tylko zwykłe maile - bez zaproszeń na spotkania (eventMessage).
        mails += [m for m in page["value"] if m.get("@odata.type", "#microsoft.graph.message") == "#microsoft.graph.message"]
        url = page.get("@odata.nextLink")
    return mails


def sender_name(mail):
    return ((mail.get("sender") or mail.get("from") or {}).get("emailAddress") or {}).get("name") or "?"


def sender_smtp(mail):
    """Adres e-mail osoby, która przesłała maila na skrzynkę pośrednią."""
    address = ((mail.get("sender") or mail.get("from") or {}).get("emailAddress") or {}).get("address") or ""
    if "@" in address:
        return address
    log.warning("Nie udało się ustalić adresu e-mail nadawcy '%s'", sender_name(mail))
    return ""


def mark_read(graph, mail):
    graph.request("PATCH", f"/messages/{mail['id']}", json={"isRead": True})


def received_year(mail):
    # Graph podaje czas w UTC - rok liczymy lokalnie (mail z Sylwestra po 23:00 to wciąż stary rok).
    received = datetime.fromisoformat(mail["receivedDateTime"].replace("Z", "+00:00"))
    return received.astimezone().year


def clickup_subject(mail_subject, sender, assign_sender):
    """Temat dla ClickUp: bez "PD:"/"FW:", z tagami ClickUp i przypisaniem nadawcy."""
    subject = REPLY_PREFIX.sub("", mail_subject or "").strip() or "(bez tematu)"
    if sender:
        tag = f"<assign {sender}>"
        # "me" oznaczałoby w ClickUp skrzynkę pośrednią, a nie handlowca.
        subject = ASSIGN_ME.sub(lambda _: tag, subject)
        if assign_sender and tag.casefold() not in subject.casefold():
            subject = f"{subject} {tag}"
    return subject


def forward_to_clickup(cfg, graph, mail, pr_dir):
    subject = clickup_subject(mail["subject"], sender_smtp(mail), cfg.assign_sender)
    subject = cfg.subject_format.format(
        subject=subject, folder=pr_dir.name, name=client_name(mail["subject"])
    )
    # Szkic przekazania (z oryginalną treścią) -> dopisanie ścieżki na górze -> wysyłka.
    draft = graph.request(
        "POST",
        f"/messages/{mail['id']}/createForward",
        json={"message": {"toRecipients": [{"emailAddress": {"address": cfg.clickup_address}}]}},
    )
    body = draft["body"]
    if body["contentType"].lower() == "html":
        link = f'<a href="{html.escape(pr_dir.as_uri())}">{html.escape(str(pr_dir))}</a>'
        note = f"<p><b>Folder dokumentacji:</b> {link}</p><hr>"
        m = re.search(r"<body[^>]*>", body["content"], re.IGNORECASE)
        pos = m.end() if m else 0
        content = body["content"][:pos] + note + body["content"][pos:]
    else:
        content = f"Folder dokumentacji: {pr_dir}\r\n\r\n" + body["content"]
    try:
        graph.request(
            "PATCH",
            f"/messages/{draft['id']}",
            json={"subject": subject, "body": {"contentType": body["contentType"], "content": content}},
        )
        graph.request("POST", f"/messages/{draft['id']}/send")  # kopia trafia do Elementów wysłanych
    except Exception:
        # Nieudana wysyłka nie zostawia szkicu - w kolejnym cyklu powstanie nowy.
        try:
            graph.request("DELETE", f"/messages/{draft['id']}")
        except Exception:
            log.warning("Nie udało się usunąć szkicu przekazania '%s'", mail["subject"])
        raise


def process_mail(cfg, graph, mail, state):
    entry = state.get(mail["id"])

    if entry and (cfg.docs_root / entry["pr"]).is_dir():
        pr_dir = cfg.docs_root / entry["pr"]
        log.info("Wznawiam przerwane przetwarzanie '%s' -> %s", mail["subject"], pr_dir.name)
    else:
        name = client_name(mail["subject"])
        if not name:
            log.warning("Brak nazwy kontrahenta w temacie '%s'", mail["subject"])
        pr_dir = create_next_folder(
            cfg.docs_root, cfg.prefix, cfg.number_width, received_year(mail), name
        )
        entry = state[mail["id"]] = {"pr": str(pr_dir.relative_to(cfg.docs_root)), "sent": False}
        save_state(state)

    fill_from_template(cfg.template, pr_dir)

    if not entry["sent"]:
        forward_to_clickup(cfg, graph, mail, pr_dir)
        entry["sent"] = True
        save_state(state)

    mark_read(graph, mail)
    del state[mail["id"]]
    save_state(state)
    log.info("Przetworzono '%s' -> %s", mail["subject"], pr_dir)


def is_allowed_sender(cfg, graph, mail):
    """Mail spoza dozwolonych domen jest oznaczany jako przeczytany i pomijany."""
    if not cfg.allowed_domains:
        return True
    sender = sender_smtp(mail)
    if sender.rpartition("@")[2].casefold() in cfg.allowed_domains:
        return True
    log.warning("Pomijam mail od '%s' (domena spoza listy): '%s'", sender or sender_name(mail), mail["subject"])
    mark_read(graph, mail)
    return False


def poll_once(cfg, graph, state):
    inbox = find_inbox(graph, cfg.subfolder)
    for mail in fetch_unread(graph, inbox):
        try:
            if not is_allowed_sender(cfg, graph, mail):
                continue
            process_mail(cfg, graph, mail, state)
        except Exception:
            # Mail zostaje nieprzeczytany i wróci w kolejnym cyklu.
            log.exception("Błąd przetwarzania maila '%s'", mail.get("subject"))


# --- Uruchomienie ------------------------------------------------------------

def setup_logging():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    log.addHandler(file_handler)
    if sys.stderr is not None:  # pod pythonw.exe nie ma konsoli
        console = logging.StreamHandler()
        console.setFormatter(fmt)
        log.addHandler(console)


def single_instance_lock():
    """Blokada pliku - zwalniana przez system, gdy proces się zakończy."""
    lock = open(LOCK_PATH, "a")
    try:
        msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        lock.close()
        return None
    return lock


def set_secret(cfg):
    print(f"Sekret zostanie zapisany w Menedżerze poświadczeń Windows użytkownika '{getpass.getuser()}'.")
    print("Skrypt musi potem działać jako ten sam użytkownik.")
    secret = getpass.getpass(f"Sekret aplikacji {cfg.client_id} (niewidoczny przy wpisywaniu): ")
    keyring.set_password(APP_NAME, cfg.client_id, secret.strip())
    print("Zapisano.")


def check(cfg):
    """Sprawdza konfigurację bez żadnych zmian w poczcie ani na dysku."""
    print(f"Szablon:        {cfg.template} -> {'OK' if cfg.template.is_dir() else 'BRAK FOLDERU'}")
    if cfg.docs_root.is_dir():
        folder_id = next_folder_id(cfg.docs_root, cfg.prefix, cfg.number_width, time.localtime().tm_year)
        year_dir = year_folder(cfg.docs_root, time.localtime().tm_year)
        print(f"Dysk:           {cfg.docs_root} -> OK, następny folder: {year_dir}\\{folder_id} - <kontrahent>")
    else:
        print(f"Dysk:           {cfg.docs_root} -> NIEDOSTĘPNY")
    graph = Graph(cfg)
    inbox = find_inbox(graph, cfg.subfolder)
    print(f"Exchange:       Microsoft Graph, aplikacja {cfg.client_id} -> OK")
    print(f"Skrzynka:       {cfg.mailbox} / {inbox['displayName']} -> OK")
    for mail in fetch_unread(graph, inbox):
        print(
            f"  nieprzeczytany: {mail['receivedDateTime']}  {mail['subject']}"
            f"  od {sender_smtp(mail)}  -> '{client_name(mail['subject'])}'"
        )
    print(f"ClickUp:        {cfg.clickup_address}")


def main():
    args = set(sys.argv[1:])
    setup_logging()

    try:
        cfg = Config(CONFIG_PATH)
    except Exception:
        log.exception("Błędna konfiguracja %s", CONFIG_PATH)
        return 1

    if "--set-secret" in args:
        set_secret(cfg)
        return 0
    if "--check" in args:
        check(cfg)
        return 0

    lock = single_instance_lock()  # noqa: F841 - trzymamy uchwyt do końca procesu
    if lock is None:
        log.info("Inna instancja już działa - kończę")
        return 0

    log.info("Start: skrzynka '%s' przez Microsoft Graph, co %s s", cfg.mailbox, cfg.poll_interval)
    state = load_state()
    graph = None

    while True:
        try:
            if graph is None:
                graph = Graph(cfg)
            poll_once(cfg, graph, state)
        except Exception:
            log.exception("Błąd połączenia z Exchange Online - ponowię w kolejnym cyklu")
            graph = None

        if "--once" in args:
            return 0
        time.sleep(cfg.poll_interval)


if __name__ == "__main__":
    sys.exit(main())
