"""Most Outlook -> dysk sieciowy -> ClickUp.

Działa w tle (pythonw.exe), co `poll_interval` sekund sprawdza nieprzeczytane
maile w skrzynce pośredniej. Dla każdego maila:
  1. zakłada kolejny folder "DIQ26001 - Kontrahent" na dysku sieciowym
     jako kopię folderu szablonu (numeracja od 001 w każdym roku),
  2. przekazuje maila (Forward) na adres Email-to-Task listy ClickUp,
     dopisując ścieżkę do folderu,
  3. oznacza maila jako przeczytanego.

Uruchomienie ręczne (z konsolą, do testów):
    python outlook_to_clickup.pyw --check   # sprawdza konfigurację, nic nie zmienia
    python outlook_to_clickup.pyw --once    # jeden cykl przetwarzania i koniec
"""

import configparser
import html
import json
import logging
import logging.handlers
import os
import re
import shutil
import sys
import time
from pathlib import Path

import pywintypes
import win32api
import win32com.client
import win32event
import winerror

APP_NAME = "ClickUpOutlookBridge"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.ini"
DATA_DIR = Path(os.environ.get("LOCALAPPDATA", BASE_DIR)) / APP_NAME
STATE_PATH = DATA_DIR / "state.json"
LOG_PATH = DATA_DIR / "bridge.log"

OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43
OL_FORMAT_HTML = 2

INBOX_NAMES = ("odebrane", "skrzynka odbiorcza", "inbox")

CLICKUP_TAG = re.compile(r"<[^<>]*>")  # <assign me>, <tag pilne> itp. w temacie
REPLY_PREFIX = re.compile(r"^(\s*(re|fw|fwd|pd|odp|wg)\s*:)+", re.IGNORECASE)
INVALID_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')  # niedozwolone w nazwach folderów Windows
ASSIGN_ME = re.compile(r"<\s*assign[\s.]+me\s*>", re.IGNORECASE)
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

log = logging.getLogger(APP_NAME)


class Config:
    def __init__(self, path):
        cp = configparser.ConfigParser(interpolation=None)
        if not cp.read(path, encoding="utf-8"):
            raise FileNotFoundError(f"Brak pliku konfiguracyjnego: {path}")
        self.mailbox = cp.get("outlook", "mailbox").strip()
        self.subfolder = cp.get("outlook", "subfolder", fallback="").strip()
        self.poll_interval = cp.getint("outlook", "poll_interval", fallback=60)
        self.allowed_domains = {
            d.strip().lstrip("@").casefold()
            for d in cp.get("outlook", "allowed_domains", fallback="").split(",")
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


def next_folder_id(root, prefix, width, year):
    """DIQ26001, DIQ26002, ... - numeracja od początku w każdym roku."""
    yy = f"{year % 100:02d}"
    pattern = re.compile(rf"^{re.escape(prefix)}{yy}(\d+)(?!\d)", re.IGNORECASE)
    numbers = [
        int(m.group(1))
        for entry in os.scandir(root)
        if entry.is_dir() and (m := pattern.match(entry.name))
    ]
    return f"{prefix}{yy}{max(numbers, default=0) + 1:0{width}d}"


def create_next_folder(root, prefix, width, year, name):
    for _ in range(20):
        folder_id = next_folder_id(root, prefix, width, year)
        path = root / (f"{folder_id} - {name}" if name else folder_id)
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


# --- Outlook -----------------------------------------------------------------

def get_outlook():
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except pywintypes.com_error:
        return None


def find_inbox(ns, mailbox, subfolder):
    store = next((s for s in ns.Stores if s.DisplayName.casefold() == mailbox.casefold()), None)
    if store is None:
        available = ", ".join(f"'{s.DisplayName}'" for s in ns.Stores)
        raise LookupError(f"Nie znaleziono skrzynki '{mailbox}'. Dostępne: {available}")
    try:
        folder = store.GetDefaultFolder(OL_FOLDER_INBOX)
    except pywintypes.com_error:
        # Część skrzynek współdzielonych nie udostępnia folderów domyślnych.
        folder = next(
            (f for f in store.GetRootFolder().Folders if f.Name.casefold() in INBOX_NAMES), None
        )
        if folder is None:
            raise LookupError(f"Skrzynka '{mailbox}' nie ma folderu Odebrane/Inbox")
    for part in filter(None, re.split(r"[\\/]", subfolder)):
        folder = folder.Folders[part]
    return folder


def fetch_unread(inbox):
    unread = inbox.Items.Restrict("[UnRead] = True")
    unread.Sort("[ReceivedTime]")
    # Kopia listy - oznaczanie jako przeczytane zmienia wynik Restrict w trakcie iteracji.
    return [item for item in unread if item.Class == OL_MAIL_ITEM]


def sender_smtp(mail):
    """Adres e-mail osoby, która przesłała maila na skrzynkę pośrednią."""
    if mail.SenderEmailType != "EX":
        return mail.SenderEmailAddress or ""
    # Nadawca z tej samej organizacji Exchange ma adres w formacie /O=.../CN=..., a nie SMTP.
    try:
        user = mail.Sender.GetExchangeUser()
        if user is not None and user.PrimarySmtpAddress:
            return user.PrimarySmtpAddress
        return mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
    except pywintypes.com_error:
        log.warning("Nie udało się ustalić adresu e-mail nadawcy '%s'", mail.SenderName)
        return ""


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


def forward_to_clickup(cfg, mail, pr_dir):
    fwd = mail.Forward()
    fwd.To = cfg.clickup_address
    subject = clickup_subject(mail.Subject, sender_smtp(mail), cfg.assign_sender)
    fwd.Subject = cfg.subject_format.format(
        subject=subject, folder=pr_dir.name, name=client_name(mail.Subject)
    )

    if fwd.BodyFormat == OL_FORMAT_HTML:
        link = f'<a href="{html.escape(pr_dir.as_uri())}">{html.escape(str(pr_dir))}</a>'
        note = f"<p><b>Folder dokumentacji:</b> {link}</p><hr>"
        body = fwd.HTMLBody
        m = re.search(r"<body[^>]*>", body, re.IGNORECASE)
        pos = m.end() if m else 0
        fwd.HTMLBody = body[:pos] + note + body[pos:]
    else:
        fwd.Body = f"Folder dokumentacji: {pr_dir}\r\n\r\n" + fwd.Body

    fwd.Send()


def process_mail(cfg, mail, state):
    entry_id = mail.EntryID
    entry = state.get(entry_id)

    if entry and (cfg.docs_root / entry["pr"]).is_dir():
        pr_dir = cfg.docs_root / entry["pr"]
        log.info("Wznawiam przerwane przetwarzanie '%s' -> %s", mail.Subject, pr_dir.name)
    else:
        name = client_name(mail.Subject)
        if not name:
            log.warning("Brak nazwy kontrahenta w temacie '%s'", mail.Subject)
        pr_dir = create_next_folder(
            cfg.docs_root, cfg.prefix, cfg.number_width, mail.ReceivedTime.year, name
        )
        entry = state[entry_id] = {"pr": pr_dir.name, "sent": False}
        save_state(state)

    fill_from_template(cfg.template, pr_dir)

    if not entry["sent"]:
        forward_to_clickup(cfg, mail, pr_dir)
        entry["sent"] = True
        save_state(state)

    mail.UnRead = False
    mail.Save()
    del state[entry_id]
    save_state(state)
    log.info("Przetworzono '%s' -> %s", mail.Subject, pr_dir)


def is_allowed_sender(cfg, mail):
    """Mail spoza dozwolonych domen jest oznaczany jako przeczytany i pomijany."""
    if not cfg.allowed_domains:
        return True
    sender = sender_smtp(mail)
    if sender.rpartition("@")[2].casefold() in cfg.allowed_domains:
        return True
    log.warning("Pomijam mail od '%s' (domena spoza listy): '%s'", sender or mail.SenderName, mail.Subject)
    mail.UnRead = False
    mail.Save()
    return False


def poll_once(cfg, outlook, state):
    inbox = find_inbox(outlook.GetNamespace("MAPI"), cfg.mailbox, cfg.subfolder)
    for mail in fetch_unread(inbox):
        try:
            if not is_allowed_sender(cfg, mail):
                continue
            process_mail(cfg, mail, state)
        except Exception:
            # Mail zostaje nieprzeczytany i wróci w kolejnym cyklu.
            log.exception("Błąd przetwarzania maila '%s'", getattr(mail, "Subject", "?"))


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


def check(cfg):
    """Sprawdza konfigurację bez żadnych zmian w poczcie ani na dysku."""
    print(f"Szablon:        {cfg.template} -> {'OK' if cfg.template.is_dir() else 'BRAK FOLDERU'}")
    if cfg.docs_root.is_dir():
        folder_id = next_folder_id(cfg.docs_root, cfg.prefix, cfg.number_width, time.localtime().tm_year)
        print(f"Dysk:           {cfg.docs_root} -> OK, następny folder: {folder_id} - <kontrahent>")
    else:
        print(f"Dysk:           {cfg.docs_root} -> NIEDOSTĘPNY")
    outlook = get_outlook()
    if outlook is None:
        print("Outlook:        nie jest uruchomiony")
        return
    inbox = find_inbox(outlook.GetNamespace("MAPI"), cfg.mailbox, cfg.subfolder)
    print(f"Skrzynka:       {cfg.mailbox} / {inbox.Name} -> OK")
    for mail in fetch_unread(inbox):
        print(f"  nieprzeczytany: {mail.ReceivedTime}  {mail.Subject}  -> '{client_name(mail.Subject)}'")
    print(f"ClickUp:        {cfg.clickup_address}")


def main():
    args = set(sys.argv[1:])
    setup_logging()

    try:
        cfg = Config(CONFIG_PATH)
    except Exception:
        log.exception("Błędna konfiguracja %s", CONFIG_PATH)
        return 1

    if "--check" in args:
        check(cfg)
        return 0

    # Tylko jedna instancja na sesję użytkownika.
    mutex = win32event.CreateMutex(None, False, f"Local\\{APP_NAME}")  # noqa: F841 - trzymamy uchwyt
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("Inna instancja już działa - kończę")
        return 0

    log.info("Start: skrzynka '%s', co %s s", cfg.mailbox, cfg.poll_interval)
    state = load_state()
    outlook = None
    waiting_logged = False

    while True:
        try:
            if outlook is None:
                outlook = get_outlook()
                if outlook is None and not waiting_logged:
                    log.info("Czekam na uruchomienie Outlooka...")
                waiting_logged = outlook is None
            if outlook is not None:
                poll_once(cfg, outlook, state)
        except pywintypes.com_error:
            log.exception("Błąd połączenia z Outlookiem - ponowię podłączenie")
            outlook = None
        except Exception:
            log.exception("Błąd cyklu")

        if "--once" in args:
            return 0
        time.sleep(cfg.poll_interval)


if __name__ == "__main__":
    sys.exit(main())
