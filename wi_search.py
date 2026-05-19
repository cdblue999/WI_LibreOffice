"""
WI Search - Wyszukiwarka Korespondencji
Python/LibreOffice port of the Outlook VBA project.

Usage:
    python wi_search.py              # Run with default config
    python wi_search.py --config wi.ini  # Custom config file
    python wi_dialog.py              # Launch GUI
"""

import os
import sys
import re
import shutil
import json
import imaplib
import email
from email.header import decode_header
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple
from dataclasses import dataclass, field

# ============================================================
# CONFIGURATION
# ============================================================
@dataclass
class Config:
    base_path: str = r"D:\DANE"
    min_attachment_size: int = 51200  # 50 KB
    keywords: List[str] = field(default_factory=lambda: [
        "umowa", "podwykonaw", "aneks", "zlecenie", "kontrakt",
        "draft", "wniosek", "wdr", "faktura", "protokol",
        "zamowienie", "sprawozdanie", "pozwolenie", "decyzja"
    ])
    image_extensions: set = field(default_factory=lambda: {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff",
        ".tif", ".webp", ".svg", ".ico", ".emz", ".wmz"
    })
    save_metadata: bool = True
    recursive_search: bool = True
    # IMAP settings for email fetching
    imap_host: str = ""
    imap_port: int = 993
    imap_user: str = ""
    imap_password: str = ""
    imap_folder: str = "INBOX"
    date_from: Optional[str] = None  # "2026-02-01"
    date_to: Optional[str] = None

    @classmethod
    def from_ini(cls, path: str) -> "Config":
        """Load config from .ini file (configparser format)."""
        import configparser
        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")
        c = cls()
        if cfg.has_section("search"):
            c.base_path = cfg.get("search", "base_path", fallback=c.base_path)
            c.min_attachment_size = cfg.getint("search", "min_attachment_size", fallback=c.min_attachment_size)
            c.recursive_search = cfg.getboolean("search", "recursive_search", fallback=c.recursive_search)
            c.save_metadata = cfg.getboolean("search", "save_metadata", fallback=c.save_metadata)
        if cfg.has_section("keywords"):
            c.keywords = [v.strip() for v in cfg.get("keywords", "list", fallback=",".join(c.keywords)).split(",")]
        if cfg.has_section("imap"):
            c.imap_host = cfg.get("imap", "host", fallback="")
            c.imap_port = cfg.getint("imap", "port", fallback=993)
            c.imap_user = cfg.get("imap", "user", fallback="")
            c.imap_password = cfg.get("imap", "password", fallback="")
            c.imap_folder = cfg.get("imap", "folder", fallback="INBOX")
            c.date_from = cfg.get("imap", "date_from", fallback=None)
            c.date_to = cfg.get("imap", "date_to", fallback=None)
        return c

    def save(self, path: str):
        """Save config to .ini file."""
        import configparser
        cfg = configparser.ConfigParser()
        cfg["search"] = {
            "base_path": self.base_path,
            "min_attachment_size": str(self.min_attachment_size),
            "recursive_search": str(self.recursive_search),
            "save_metadata": str(self.save_metadata),
        }
        cfg["keywords"] = {"list": ",".join(self.keywords)}
        if self.imap_host:
            cfg["imap"] = {
                "host": self.imap_host,
                "port": str(self.imap_port),
                "user": self.imap_user,
                "password": self.imap_password,
                "folder": self.imap_folder,
            }
            if self.date_from:
                cfg["imap"]["date_from"] = self.date_from
            if self.date_to:
                cfg["imap"]["date_to"] = self.date_to
        with open(path, "w", encoding="utf-8") as f:
            cfg.write(f)


# ============================================================
# UTILITY FUNCTIONS
# ============================================================
def clean_filename(name: str) -> str:
    """Remove illegal filename characters."""
    for c in r'\/:*?"<>|':
        name = name.replace(c, "_")
    return name


def czy_to_obraz(nazwa: str) -> bool:
    """Check if file is an image (skip these for AI processing)."""
    return Path(nazwa).suffix.lower() in Config().image_extensions


def czy_to_dokument_umowny(nazwa: str, keywords: List[str]) -> bool:
    """Check if filename matches contract-related keywords."""
    ext = Path(nazwa).suffix.lower()
    if ext not in {".pdf", ".doc", ".docx", ".xls", ".xlsx"}:
        return False
    n = nazwa.lower()
    return any(k in n for k in keywords)


def save_metadata(dest_file: str, mail_data: dict, fso=None):
    """Save email metadata alongside the attachment."""
    meta_path = Path(dest_file).with_suffix(".meta")
    lines = [
        f"From: {mail_data.get('from', 'N/A')}",
        f"To: {mail_data.get('to', 'N/A')}",
        f"CC: {mail_data.get('cc', '')}",
        f"Subject: {mail_data.get('subject', 'N/A')}",
        f"Received: {mail_data.get('date', 'N/A')}",
        f"OriginalFile: {Path(dest_file).name}",
    ]
    meta_path.write_text("\n".join(lines), encoding="utf-8")


# ============================================================
# CONFLICT RESOLUTION
# ============================================================
class ConflictResolver:
    """Handles file overwrite/skip decisions."""
    def __init__(self):
        self.overwrite_all = False
        self.skip_all = False

    def resolve(self, dest: Path) -> bool:
        if not dest.exists():
            return True
        if self.overwrite_all:
            return True
        if self.skip_all:
            return False
        # Interactive mode: ask user
        print(f"  Konflikt: {dest.name}")
        print("  [T] Nadpisz  [N] Pomiń  [A] Wszystkie nadpisz  [P] Wszystkie pomiń")
        choice = input("  Wybór: ").strip().upper()
        if choice == "T":
            return True
        if choice == "N":
            return False
        if choice == "A":
            self.overwrite_all = True
            return True
        if choice == "P":
            self.skip_all = True
            return False
        return False


# ============================================================
# SERVER FILE SEARCH
# ============================================================
def szukaj_rekurencyjnie(
    sciezka: str,
    cel: str,
    klucze: List[str],
    resolver: ConflictResolver,
    config: Config,
) -> int:
    """Recursively search folders for matching files."""
    ile = 0
    base = Path(sciezka)
    if not base.is_dir():
        return 0

    for root, _, files in os.walk(base):
        for plik in files:
            # All keywords must match
            if not all(k.lower() in plik.lower() for k in klucze):
                continue
            if czy_to_obraz(plik):
                continue

            dest = Path(cel) / f"Serwer_{plik}"
            if resolver.resolve(dest):
                shutil.copy2(Path(root) / plik, dest)
                ile += 1
    return ile


def szukaj_plosko(
    sciezka: str,
    cel: str,
    klucze: List[str],
    resolver: ConflictResolver,
    config: Config,
) -> int:
    """Search only top-level folder (legacy behavior)."""
    ile = 0
    base = Path(sciezka)
    if not base.is_dir():
        return 0

    for plik in base.iterdir():
        if not plik.is_file():
            continue
        if not all(k.lower() in plik.name.lower() for k in klucze):
            continue
        if czy_to_obraz(plik.name):
            continue

        dest = Path(cel) / f"Serwer_{plik.name}"
        if resolver.resolve(dest):
            shutil.copy2(plik, dest)
            ile += 1
    return ile


# ============================================================
# IMAP EMAIL PROCESSING
# ============================================================
def decode_mime_header(header: str) -> str:
    """Decode MIME-encoded email headers."""
    if not header:
        return ""
    decoded = decode_header(header)
    parts = []
    for data, charset in decoded:
        if isinstance(data, bytes):
            parts.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(data)
    return " ".join(parts)


def fetch_emails(config: Config) -> List[dict]:
    """Fetch emails from IMAP server matching date range."""
    if not config.imap_host:
        print("  [INFO] Brak konfiguracji IMAP — pomijam pobieranie maili.")
        return []

    print(f"  Łączę z {config.imap_host}:{config.imap_port}...")
    mail = imaplib.IMAP4_SSL(config.imap_host, config.imap_port)
    mail.login(config.imap_user, config.imap_password)
    mail.select(config.imap_folder)

    # Build search criteria
    criteria = ["ALL"]
    if config.date_from:
        criteria.append(f'SINCE "{config.date_from}"')
    if config.date_to:
        criteria.append(f'BEFORE "{config.date_to}"')

    status, messages = mail.search(None, *criteria)
    if status != "OK":
        print("  [BŁĄD] Nie udało się pobrać listy maili.")
        return []

    email_ids = messages[0].split()
    print(f"  Znaleziono {len(email_ids)} maili.")

    results = []
    for eid in email_ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        mail_info = {
            "subject": decode_mime_header(msg.get("Subject", "")),
            "from": decode_mime_header(msg.get("From", "")),
            "to": decode_mime_header(msg.get("To", "")),
            "cc": decode_mime_header(msg.get("CC", "")),
            "date": msg.get("Date", ""),
            "attachments": [],
        }

        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue
            filename = part.get_filename()
            if not filename:
                continue
            filename = decode_mime_header(filename)
            payload = part.get_payload(decode=True)
            if payload:
                mail_info["attachments"].append({
                    "filename": filename,
                    "data": payload,
                    "size": len(payload),
                })

        results.append(mail_info)

    mail.logout()
    return results


def process_emails(
    emails: List[dict],
    folder_zapisu: str,
    config: Config,
    resolver: ConflictResolver,
) -> int:
    """Save matching attachments from fetched emails."""
    licznik = 0
    for mail in emails:
        for att in mail["attachments"]:
            if att["size"] < config.min_attachment_size:
                continue
            if czy_to_obraz(att["filename"]):
                continue
            if not czy_to_dokument_umowny(att["filename"], config.keywords):
                continue

            safe_name = clean_filename(att["filename"])
            dest = Path(folder_zapisu) / f"Outlook_{mail['date'][:10]}_{safe_name}"

            if resolver.resolve(dest):
                dest.write_bytes(att["data"])
                if config.save_metadata:
                    save_metadata(str(dest), mail)
                licznik += 1

    return licznik


# ============================================================
# MAIN ENTRY POINT
# ============================================================
def run(config: Config, output_dir: Optional[str] = None):
    """Execute the full search & archive workflow."""
    resolver = ConflictResolver()

    # Determine output directory
    if output_dir:
        base_out = Path(output_dir)
    else:
        base_out = Path(input("Folder na wyniki: ").strip())

    base_out.mkdir(parents=True, exist_ok=True)
    folder_zapisu = base_out / "wyniki"
    folder_zapisu.mkdir(exist_ok=True)

    print(f"\n📁 Wyniki: {folder_zapisu}")
    print(f"🔑 Klucze: {', '.join(config.keywords)}")
    print(f"📏 Min rozmiar: {config.min_attachment_size} B")
    print(f"🔄 Rekurencja: {'TAK' if config.recursive_search else 'NIE'}")
    print()

    # --- ETAP A: EMAIL ---
    licznik_mail = 0
    if config.imap_host:
        print("=== ETAP A: Pobieranie maili ===")
        emails = fetch_emails(config)
        licznik_mail = process_emails(emails, str(folder_zapisu), config, resolver)
        print(f"  Zapisano: {licznik_mail} załączników\n")

    # --- ETAP B: SERWER ---
    licznik_serwer = 0
    if config.base_path and Path(config.base_path).is_dir():
        print("=== ETAP B: Przeszukiwanie serwera ===")
        if config.recursive_search:
            licznik_serwer = szukaj_rekurencyjnie(
                config.base_path, str(folder_zapisu), config.keywords, resolver, config
            )
        else:
            licznik_serwer = szukaj_plosko(
                config.base_path, str(folder_zapisu), config.keywords, resolver, config
            )
        print(f"  Zapisano: {licznik_serwer} plików\n")

    # --- PODSUMOWANIE ---
    print("=" * 40)
    print(f"✅ Eksport zakończony!")
    print(f"   Mail: {licznik_mail} | Serwer: {licznik_serwer}")
    print(f"   Łącznie: {licznik_mail + licznik_serwer}")
    print(f"   Folder: {folder_zapisu}")

    # Open folder in file explorer
    os.startfile(str(folder_zapisu))


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="WI Search - Wyszukiwarka Korespondencji")
    parser.add_argument("--config", "-c", help="Path to .ini config file", default="wi.ini")
    parser.add_argument("--output", "-o", help="Output directory")
    parser.add_argument("--init-config", action="store_true", help="Generate default wi.ini and exit")
    args = parser.parse_args()

    if args.init_config:
        cfg = Config()
        cfg.save("wi.ini")
        print("✅ Utworzono wi.ini — edytuj i uruchom ponownie.")
        sys.exit(0)

    cfg = Config.from_ini(args.config) if Path(args.config).exists() else Config()
    run(cfg, args.output)
