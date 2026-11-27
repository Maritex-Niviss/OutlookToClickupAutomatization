"""Most Exchange Online -> dysk sieciowy -> ClickUp.

Łączy się z Exchange Online przez Microsoft Graph (bez Outlooka), jako aplikacja
zarejestrowana w Entra ID, i co `poll_interval` sekund sprawdza nieprzeczytane
maile w skrzynce pośredniej. Dla każdego maila:
  1. z pierwszego słowa tematu ustala kategorię (sekcje [type ...] w config.ini)
     i zakłada folder jako kopię folderu szablonu:
       "RFQ Firma XYZ" -> <root>\\2026\\RFQ26001 - Firma XYZ  (kolejny numer w roku),
       "C Lattice"     -> <root>\\Lattice\\CTM2601          (projekt w folderze firmy),
  2. przekazuje maila (Forward) na adres Email-to-Task listy ClickUp - zadanie
     nazywa się jak folder, a jego opis to treść maila bez stopek i nagłówków
     "Od:/Wysłane:/Do:",
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
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from html.parser import HTMLParser
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
# Początek tematu = typ dokumentu ("RFQ Firma", " rfq:Firma", "<tag x> RFQ - Firma").
LEADING_TAGS = re.compile(r"\s*((?:<[^<>]*>\s*)*)")
TYPE_SEPARATOR = re.compile(r"[\s:;.,\-–—_/|]*")
# Niewidoczne znaki wklejane z Worda/Teams (spacje zerowej szerokości, BOM, miękki dywiz).
INVISIBLE = re.compile(r"[\u00ad\u180e\u200b-\u200f\u202a-\u202e\u2060-\u2064\ufeff]")

# Początek stopki: linia zaczynająca się od jednego z tych zwrotów (nadpisywane w config.ini).
DEFAULT_SIGNATURE_MARKERS = (
    "Pozdrawiam, Pozdrowienia, Serdecznie pozdrawiam, Serdeczne pozdrowienia, Z poważaniem, "
    "Z wyrazami szacunku, Best regards, Kind regards, Best wishes, Regards, "
    "Mit freundlichen Grüßen, Viele Grüße"
)
# Nagłówek przekazanej/cytowanej wiadomości ("Od: ... Wysłane: ... Do: ... Temat: ...").
HEADER_START = re.compile(
    r"^((from|od|von|de)\s*:|-{2,}.*(original|oryginaln|forwarded|przekazan|ursprüngliche).*-{2,})",
    re.IGNORECASE,
)
HEADER_FIELD = re.compile(
    r"^(sent|wysłane|wysłano|date|data|to|do|cc|dw|subject|temat|gesendet|an|betreff"
    r"|importance|ważność|attachments|załączniki)\s*:",
    re.IGNORECASE,
)

log = logging.getLogger(APP_NAME)


LAYOUTS = ("year", "company")


@dataclass
class DocType:
    """Typ dokumentu = słowo na początku tematu (DIQ, RFQ, C, R...).

    layout = "year":    <root>\\<rok>\\DIQ26001 - Firma        (numer w roku, number_width cyfr)
    layout = "company": <root>\\<Firma>\\CTM2601              (kod firmy szyfrem Cezara, numer w roku
                                                              w folderze firmy)
    """
    name: str
    layout: str
    root: Path
    template: Path
    clickup_address: str
    number_width: int
    cipher_shift: int


@dataclass
class Route:
    """[route NAZWA]: zadania od tych nadawców idą na inną listę ClickUp (dla wszystkich kategorii)."""
    name: str
    senders: frozenset
    clickup_address: str
    create_folders: bool = True  # no = tylko zadanie w ClickUp, nic na dysku


def _get(cp, section, option, legacy_section=None, fallback=""):
    """Odczyt z nowej sekcji, a gdy jej brak - z miejsca, w którym była w starym config.ini."""
    value = cp.get(section, option, fallback=None)
    if value is None and legacy_section:
        value = cp.get(legacy_section, option, fallback=None)
    return (value if value is not None else fallback).strip()


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
            for d in _get(cp, "mail", "allowed_domains", legacy_section="exchange").split(",")
            if d.strip()
        }
        default_address = cp.get("clickup", "address", fallback="").strip()
        self.doc_types = {}
        for section in cp.sections():
            kind, _, names = section.partition(" ")
            if kind.casefold() != "type" or not names.strip():
                continue
            s = cp[section]
            layout = s.get("layout", "year").strip().casefold()
            if layout not in LAYOUTS:
                raise ValueError(f"[{section}] layout = {layout} - dozwolone: {', '.join(LAYOUTS)}")
            address = s.get("clickup_address", "").strip() or default_address
            if not address:
                raise ValueError(f"[{section}] brak clickup_address")
            for name in re.split(r"[\s,;/]+", names.strip().upper()):  # [type C, R] = dwa typy
                if name:
                    self.doc_types[name] = DocType(
                        name=name,
                        layout=layout,
                        root=Path(s["root"].strip()),
                        template=Path(s["template"].strip()),
                        clickup_address=address,
                        number_width=s.getint("number_width", fallback=3 if layout == "year" else 2),
                        cipher_shift=s.getint("cipher_shift", fallback=8),
                    )
        if not self.doc_types:
            raise ValueError("Brak typów dokumentów - dodaj sekcje [type DIQ], [type RFQ] itd.")
        default = _get(cp, "mail", "default_type", legacy_section="documents").upper()
        if default and default not in self.doc_types:
            raise ValueError(f"default_type = {default}, a nie ma sekcji [type {default}]")
        self.default_type = self.doc_types.get(default)
        self.routes = []
        for section in cp.sections():
            kind, _, name = section.partition(" ")
            if kind.casefold() != "route":
                continue
            senders = frozenset(
                s.strip().casefold() for s in re.split(r"[\s,;]+", cp.get(section, "senders", fallback="")) if s.strip()
            )
            address = cp.get(section, "clickup_address", fallback="").strip()
            if not senders or not address:
                raise ValueError(f"[{section}] wymaga senders i clickup_address")
            self.routes.append(Route(
                name.strip() or section, senders, address,
                create_folders=cp.getboolean(section, "create_folders", fallback=True),
            ))
        self.subject_format = cp.get("clickup", "subject", fallback="{subject}").strip()
        self.assign_sender = cp.getboolean("clickup", "assign_sender", fallback=True)
        markers = cp.get("clickup", "signature_markers", fallback=DEFAULT_SIGNATURE_MARKERS)
        self.signature_markers = [m.strip().casefold() for m in markers.split(",") if m.strip()]


# --- Dysk sieciowy -----------------------------------------------------------

def split_type(cfg, subject):
    """Typ dokumentu z początku tematu i temat bez słowa typu.

    "PD: RFQ Firma XYZ <tag x>" -> (RFQ, "Firma XYZ <tag x>"). Wielkość liter, spacje
    (też twarde i niewidoczne), tabulatory i separator po słowie (":", "-", ".") są bez
    znaczenia: "rfq:Firma", "  Rfq  -  Firma", "RFQTestfirma" -> RFQ. Cyfra zaraz po
    słowie typu oznacza numer istniejącego folderu ("RFQ26022 - ..."), a nie typ.
    Dotyczy to też jednoliterowych typów: "C Lattice", "r:Lattice", "RRadwag" -> R + "Radwag".
    Bez rozpoznanego typu: typ domyślny (default_type) albo None - mail jest pomijany.
    """
    subject = REPLY_PREFIX.sub("", INVISIBLE.sub("", subject or "")).strip()
    tags = LEADING_TAGS.match(subject)
    rest = subject[tags.end():]
    for name in sorted(cfg.doc_types, key=len, reverse=True):  # dłuższa nazwa wygrywa (RFQ przed R)
        after = rest[len(name):]
        if rest[:len(name)].upper() != name or after[:1].isdigit():
            continue
        after = after[TYPE_SEPARATOR.match(after).end():]
        return cfg.doc_types[name], tags.group(1) + after
    return cfg.default_type, subject


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


def company_folder_name(subject):
    """Nazwa folderu firmy dla projektów C/R: kontrahent z tematu bez spacji ("Base System" -> "BaseSystem")."""
    return "".join(client_name(subject).split())


def project_code(letter, company, shift):
    """C/R + pierwsza i ostatnia litera nazwy firmy przesunięte szyfrem Cezara: ("C", "Lattice", 8) -> "CTM".

    Liczą się tylko litery (polskie znaki bez ogonków); None, gdy nazwa nie ma żadnej litery.
    """
    plain = unicodedata.normalize("NFKD", company.replace("ł", "l").replace("Ł", "L")).upper()
    letters = [c for c in plain if "A" <= c <= "Z"]
    if not letters:
        return None
    shifted = [chr((ord(c) - ord("A") + shift) % 26 + ord("A")) for c in (letters[0], letters[-1])]
    return letter + "".join(shifted)


def _match_key(name):
    return re.sub(r"[\W_]", "", name).casefold()


def find_company_dir(root, company):
    """Istniejący folder firmy (bez względu na wielkość liter, spacje, "-" i "_") albo nowy."""
    key = _match_key(company)
    for entry in os.scandir(root):
        if entry.is_dir() and _match_key(entry.name) == key:
            return Path(entry.path)
    return root / company


def next_project_id(company_dir, code, width, year):
    """CTM2601, CTM2602... - numeracja w roku, osobno dla C i R, w obrębie folderu firmy."""
    yy = f"{year % 100:02d}"
    pattern = re.compile(rf"^{re.escape(code[0])}[A-Za-z]{{2}}{yy}(\d+)(?!\d)", re.IGNORECASE)
    numbers = [
        int(m.group(1))
        for entry in (os.scandir(company_dir) if company_dir.is_dir() else ())
        if entry.is_dir() and (m := pattern.match(entry.name))
    ]
    return f"{code}{yy}{max(numbers, default=0) + 1:0{width}d}"


def create_project_folder(doc_type, year, company):
    code = project_code(doc_type.name, company, doc_type.cipher_shift)
    company_dir = find_company_dir(doc_type.root, company)
    company_dir.mkdir(exist_ok=True)  # nowa firma - nowy folder; istniejąca - wchodzimy do niej
    for _ in range(20):
        path = company_dir / next_project_id(company_dir, code, doc_type.number_width, year)
        try:
            path.mkdir()
            return path
        except FileExistsError:
            continue
    raise RuntimeError(f"Nie udało się założyć nowego projektu {code} w {company_dir}")


def create_folder(doc_type, year, subject):
    """Nowy folder dla maila - wg układu typu (rok albo firma)."""
    if doc_type.layout == "company":
        return create_project_folder(doc_type, year, company_folder_name(subject))
    return create_next_folder(doc_type.root, doc_type.name, doc_type.number_width, year, client_name(subject))


def folder_label(doc_type, pr_dir):
    """Nazwa zadania w ClickUp ({folder}): "DIQ26001 - Firma" albo "CTM2601 - Lattice"."""
    return f"{pr_dir.name} - {pr_dir.parent.name}" if doc_type.layout == "company" else pr_dir.name


def _copy_if_missing(src, dst):
    # Przy wznowieniu nie nadpisujemy plików już skopiowanych (lub zmienionych).
    if not os.path.exists(dst):
        shutil.copy2(src, dst)


def fill_from_template(template, pr_dir):
    shutil.copytree(template, pr_dir, dirs_exist_ok=True, copy_function=_copy_if_missing)


# --- Treść zadania: bez stopek i nagłówków przekazania ----------------------

class _LineFilter:
    """Decyduje linia po linii, co zostaje w treści.

    Stopka (od "Pozdrawiam", "Best regards"...) jest wycinana do następnej cytowanej
    wiadomości albo do końca; nagłówki cytowanych wiadomości (Od:/Wysłane:/Do:/Temat:)
    są usuwane, a sama treść cytowanych wiadomości zostaje.
    """

    def __init__(self, markers):
        self.markers = markers
        self.mode = "body"  # body | header | signature

    def keep(self, line):
        """`line` - niepusty, przycięty początek linii. Zwraca True, jeśli linia zostaje."""
        if HEADER_START.match(line):
            self.mode = "header"
            return False
        if self.mode == "header":
            if HEADER_FIELD.match(line):
                return False
            self.mode = "body"
        if self.mode == "body":
            low = line.casefold()
            # Krótka linia, np. "Pozdrawiam," albo "Best regards, Jan" - a nie zdanie w treści.
            if any(low.startswith(m) and len(line) <= len(m) + 40 for m in self.markers):
                self.mode = "signature"
        return self.mode == "body"


class _HtmlCleaner(HTMLParser):
    """Przepisuje HTML maila bez stopek i nagłówków, zachowując formatowanie reszty."""

    VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "wbr"}
    LINE_BREAK = {"br", "p", "div", "tr", "td", "th", "li", "ul", "ol", "table", "blockquote", "hr",
                  "h1", "h2", "h3", "h4", "h5", "h6"}
    HIDDEN = {"head", "style", "script", "title"}

    def __init__(self, markers):
        super().__init__(convert_charrefs=True)
        self.filter = _LineFilter(markers)
        self.out = []
        self.stack = []  # [tag, czy_wypisany] - żeby domknąć tylko to, co wypisaliśmy
        self.hidden = 0  # wewnątrz <head>/<style> - tekst niewidoczny, nie liczy się jako linia
        self.line_seen = False
        self.line_kept = True

    @property
    def cutting(self):
        return self.filter.mode == "signature"

    def _new_line(self):
        self.line_seen, self.line_kept = False, True

    def handle_starttag(self, tag, attrs):
        # <br> kończący usuniętą linię nagłówka też znika, żeby nie zostawiać pustych linii.
        emit = not self.cutting and not (tag == "br" and not self.line_kept)
        if tag in self.LINE_BREAK:
            self._new_line()
        if emit:
            self.out.append(self.get_starttag_text())
        if tag not in self.VOID:
            self.stack.append([tag, emit])
            if tag in self.HIDDEN:
                self.hidden += 1

    def handle_endtag(self, tag):
        if tag in self.VOID:
            return
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                break
        else:
            return  # zamknięcie bez otwarcia - pomijamy
        for name, emitted in reversed(self.stack[i:]):
            if name in self.HIDDEN:
                self.hidden -= 1
            if emitted:
                self.out.append(f"</{name}>")
        del self.stack[i:]
        if tag in self.LINE_BREAK:
            self._new_line()

    def handle_data(self, data):
        if self.hidden:
            if not self.cutting:
                self.out.append(data)  # CSS w <style> bez zmian
            return
        if not self.line_seen and data.strip():
            self.line_seen = True
            self.line_kept = self.filter.keep(data.strip())
        if self.line_kept and not self.cutting:
            self.out.append(html.escape(data, quote=False))

    def handle_comment(self, data):
        if not self.cutting:
            self.out.append(f"<!--{data}-->")

    def handle_decl(self, decl):
        self.out.append(f"<!{decl}>")

    def handle_pi(self, data):
        self.out.append(f"<?{data}>")

    def unknown_decl(self, data):
        if not self.cutting:
            self.out.append(f"<![{data}]>")

    def result(self):
        self.close()
        self.out += [f"</{name}>" for name, emitted in reversed(self.stack) if emitted]
        return "".join(self.out)


def clean_body(body, markers):
    """Treść maila dla ClickUp: bez stopek i nagłówków Od:/Wysłane:/Do: (HTML albo tekst)."""
    if body["contentType"].lower() == "html":
        cleaner = _HtmlCleaner(markers)
        cleaner.feed(body["content"])
        return cleaner.result()
    line_filter, lines = _LineFilter(markers), []
    for line in body["content"].splitlines():
        if line.strip():
            if line_filter.keep(line.strip()):
                lines.append(line)
        elif line_filter.mode == "body":
            lines.append(line)
    return "\r\n".join(lines).strip()


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


def drop_unused_inline_images(graph, message_id, content):
    """Usuwa ze szkicu obrazki osadzone (logo ze stopki itp.), do których treść już się nie odwołuje.

    Inaczej ClickUp dodałby je do zadania jako załączniki. Zwykłe załączniki zostają.
    """
    used = {cid.casefold() for cid in re.findall(r"cid:([^\"'\s>)]+)", content, re.IGNORECASE)}
    attachments = graph.request(
        "GET", f"/messages/{message_id}/attachments", params={"$select": "id,isInline,contentType"}
    )
    for att in attachments["value"]:
        if not att.get("isInline") or not (att.get("contentType") or "").lower().startswith("image/"):
            continue
        full = graph.request("GET", f"/messages/{message_id}/attachments/{att['id']}")
        if (full.get("contentId") or "").strip("<>").casefold() not in used:
            graph.request("DELETE", f"/messages/{message_id}/attachments/{att['id']}")


def find_route(cfg, sender):
    """Pierwsza trasa [route ...], do której należy nadawca, albo None."""
    return next((r for r in cfg.routes if sender.casefold() in r.senders), None)


def clickup_target(cfg, doc_type, sender):
    """Adres listy ClickUp: z trasy nadawcy, inaczej z kategorii. Zwraca (adres, nazwa trasy albo None)."""
    route = find_route(cfg, sender)
    return (route.clickup_address, route.name) if route else (doc_type.clickup_address, None)


def forward_to_clickup(cfg, graph, mail, address, label, mail_subject):
    """Wysyła maila jako zadanie na listę `address`.

    `label` - nazwa zadania ({folder}), np. nazwa folderu; `mail_subject` - temat bez słowa typu.
    """
    subject = clickup_subject(mail_subject, sender_smtp(mail), cfg.assign_sender)
    subject = cfg.subject_format.format(
        subject=subject,
        folder=label,
        name=client_name(mail_subject),
        tags=" ".join(CLICKUP_TAG.findall(subject)),  # <assign ...>, <tag ...> - ClickUp je usuwa z nazwy
    ).strip()
    # Treść = oryginalna treść maila bez stopek i nagłówków (bez nagłówka "Od: ... Do: skrzynka").
    original = graph.request("GET", f"/messages/{mail['id']}", params={"$select": "body"})["body"]
    content = clean_body(original, cfg.signature_markers)
    # Szkic przekazania (z załącznikami) -> podmiana tematu i treści -> wysyłka.
    draft = graph.request(
        "POST",
        f"/messages/{mail['id']}/createForward",
        json={"message": {"toRecipients": [{"emailAddress": {"address": address}}]}},
    )
    try:
        graph.request(
            "PATCH",
            f"/messages/{draft['id']}",
            json={"subject": subject, "body": {"contentType": original["contentType"], "content": content}},
        )
        drop_unused_inline_images(graph, draft["id"], content)
        graph.request("POST", f"/messages/{draft['id']}/send")  # kopia trafia do Elementów wysłanych
    except Exception:
        # Nieudana wysyłka nie zostawia szkicu - w kolejnym cyklu powstanie nowy.
        try:
            graph.request("DELETE", f"/messages/{draft['id']}")
        except Exception:
            log.warning("Nie udało się usunąć szkicu przekazania '%s'", mail["subject"])
        raise


def is_auto_generated(graph, mail):
    """Autoodpowiedź (np. "Poza biurem") albo mail masowy - na takie nie odpowiadamy (pętle)."""
    message = graph.request("GET", f"/messages/{mail['id']}", params={"$select": "internetMessageHeaders"})
    for header in message.get("internetMessageHeaders") or ():
        name = (header.get("name") or "").casefold()
        value = (header.get("value") or "").strip().casefold()
        if name == "auto-submitted" and value != "no":
            return True
        if name == "precedence" and value in ("bulk", "junk", "list", "auto_reply"):
            return True
    return False


def or_list(names):
    """["DIQ", "RFQ", "C", "R"] -> "DIQ, RFQ, C albo R"."""
    names = list(names)
    return names[0] if len(names) == 1 else ", ".join(names[:-1]) + " albo " + names[-1]


def reject_mail(cfg, graph, mail, problem, hint):
    """Pomija maila: informuje nadawcę (o ile to nie autoodpowiedź), loguje i oznacza jako przeczytany."""
    replied = reply_to_sender(cfg, graph, mail, problem, hint)
    log.warning(
        "Pomijam mail '%s' od %s - %s%s",
        mail["subject"], sender_smtp(mail), problem, " (wysłano informację do nadawcy)" if replied else "",
    )
    mark_read(graph, mail)


def reply_to_sender(cfg, graph, mail, problem, hint):
    """Krótka informacja do nadawcy, dlaczego mail nie został przetworzony. Zwraca True, jeśli wysłano."""
    sender = sender_smtp(mail)
    if not sender or sender.casefold() == cfg.mailbox.casefold() or is_auto_generated(graph, mail):
        return False
    text = (
        f"Nie udało się przetworzyć maila \"{mail['subject']}\" - {problem}.\r\n\r\n"
        f"{hint} Wyślij maila ponownie z poprawionym tematem.\r\n\r\n"
        "(wiadomość wysłana automatycznie)"
    )
    graph.request(
        "POST",
        "/sendMail",
        json={
            "message": {
                "subject": f"Nie przetworzono: {mail['subject']}",
                "body": {"contentType": "text", "content": text},
                "toRecipients": [{"emailAddress": {"address": sender}}],
                # Bez autoodpowiedzi "Poza biurem" zwrotnie do skrzynki pośredniej.
                "internetMessageHeaders": [{"name": "X-Auto-Response-Suppress", "value": "All"}],
            },
            "saveToSentItems": True,
        },
    )
    return True


def process_task_only(cfg, graph, mail, state, route):
    """Trasa z create_folders = no: tylko zadanie w ClickUp - nic na dysku, kategoria niewymagana.

    Nazwa zadania = temat maila bez "PD:"/"RE:" (razem z DIQ/RFQ/C/R, jeśli są) + tagi ClickUp.
    """
    subject = REPLY_PREFIX.sub("", INVISIBLE.sub("", mail["subject"] or "")).strip()
    entry = state.get(mail["id"])
    if not (entry and entry.get("sent")):
        label = " ".join(CLICKUP_TAG.sub(" ", subject).split()) or "(bez tematu)"
        forward_to_clickup(cfg, graph, mail, route.clickup_address, label, subject)
        state[mail["id"]] = {"pr": None, "sent": True}
        save_state(state)
    mark_read(graph, mail)
    state.pop(mail["id"], None)
    save_state(state)
    log.info("Przetworzono '%s' -> tylko zadanie ClickUp (trasa '%s', bez folderu)", mail["subject"], route.name)


def process_mail(cfg, graph, mail, state):
    route = find_route(cfg, sender_smtp(mail))
    if route and not route.create_folders:
        process_task_only(cfg, graph, mail, state, route)
        return

    doc_type, subject = split_type(cfg, mail["subject"])
    if doc_type is None:
        types = or_list(cfg.doc_types)
        reject_mail(
            cfg, graph, mail,
            f"brakuje {types} na początku tematu",
            f"Temat musi zaczynać się od {types}, np. \"{next(iter(cfg.doc_types))} Firma XYZ\".",
        )
        return
    if doc_type.layout == "company":
        company = company_folder_name(subject)
        if not project_code(doc_type.name, company, doc_type.cipher_shift):
            reject_mail(
                cfg, graph, mail,
                f"brakuje nazwy firmy po {doc_type.name}",
                f"Temat musi zawierać nazwę firmy, np. \"{doc_type.name} Lattice\".",
            )
            return

    entry = state.get(mail["id"])
    if entry and entry.get("pr") and (doc_type.root / entry["pr"]).is_dir():
        pr_dir = doc_type.root / entry["pr"]
        log.info("Wznawiam przerwane przetwarzanie '%s' -> %s", mail["subject"], pr_dir.name)
    else:
        if not client_name(subject):
            log.warning("Brak nazwy kontrahenta w temacie '%s'", mail["subject"])
        pr_dir = create_folder(doc_type, received_year(mail), subject)
        entry = state[mail["id"]] = {"pr": str(pr_dir.relative_to(doc_type.root)), "sent": False}
        save_state(state)

    fill_from_template(doc_type.template, pr_dir)

    if not entry["sent"]:
        address, route_name = clickup_target(cfg, doc_type, sender_smtp(mail))
        if route_name:
            log.info("Nadawca %s -> trasa '%s' (%s)", sender_smtp(mail), route_name, address)
        forward_to_clickup(cfg, graph, mail, address, folder_label(doc_type, pr_dir), subject)
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
    year = time.localtime().tm_year
    for t in cfg.doc_types.values():
        default = " (domyślny, gdy temat nie zaczyna się od typu)" if t is cfg.default_type else ""
        print(f"Typ {t.name}{default}:")
        print(f"  Szablon:      {t.template} -> {'OK' if t.template.is_dir() else 'BRAK FOLDERU'}")
        if not t.root.is_dir():
            print(f"  Dysk:         {t.root} -> NIEDOSTĘPNY")
        elif t.layout == "company":
            print(f"  Dysk:         {t.root} -> OK, nowy projekt: {t.root}\\<Firma>\\{t.name}xx{year % 100:02d}nn")
        else:
            folder_id = next_folder_id(t.root, t.name, t.number_width, year)
            print(f"  Dysk:         {t.root} -> OK, następny folder: {year_folder(t.root, year)}\\{folder_id} - <kontrahent>")
        print(f"  ClickUp:      {t.clickup_address}")
    if cfg.default_type is None:
        print(f"Maile bez {or_list(cfg.doc_types)} na początku tematu: pomijane (default_type puste)")
    for route in cfg.routes:
        print(f"Trasa '{route.name}': {', '.join(sorted(route.senders))}")
        print(f"  ClickUp:      {route.clickup_address} (wszystkie kategorie)")
        print(f"  Foldery:      {'tak' if route.create_folders else 'NIE - tylko zadanie w ClickUp, nic na dysku'}")
    graph = Graph(cfg)
    inbox = find_inbox(graph, cfg.subfolder)
    print(f"Exchange:       Microsoft Graph, aplikacja {cfg.client_id} -> OK")
    print(f"Skrzynka:       {cfg.mailbox} / {inbox['displayName']} -> OK")
    for mail in fetch_unread(graph, inbox):
        doc_type, subject = split_type(cfg, mail["subject"])
        route = find_route(cfg, sender_smtp(mail))
        if route and not route.create_folders:
            print(f"  nieprzeczytany: {mail['receivedDateTime']}  {mail['subject']}  od {sender_smtp(mail)}"
                  f"  -> tylko zadanie ClickUp (trasa '{route.name}'), bez folderu")
            continue
        if doc_type is None:
            target = "POMINIĘTY (brak typu)"
        elif doc_type.layout == "company":
            company = company_folder_name(subject)
            code = project_code(doc_type.name, company, doc_type.cipher_shift)
            if code:
                company_dir = find_company_dir(doc_type.root, company)
                state = "istniejący" if company_dir.is_dir() else "nowy"
                project = next_project_id(company_dir, code, doc_type.number_width, year)
                target = f"{doc_type.name} {company_dir} ({state}) \\ {project}"
            else:
                target = "POMINIĘTY (brak nazwy firmy)"
        else:
            target = f"{doc_type.name} '{client_name(subject)}'"
        if doc_type is not None and (route := clickup_target(cfg, doc_type, sender_smtp(mail))[1]):
            target += f"  [ClickUp: trasa '{route}']"
        print(f"  nieprzeczytany: {mail['receivedDateTime']}  {mail['subject']}  od {sender_smtp(mail)}  -> {target}")


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
