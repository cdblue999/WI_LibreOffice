# WI Search — Wyszukiwarka Korespondencji

Pythonowy port projektu VBA/Outlook **WI_Outlook2013**. Działa niezależnie od Outlooka — wystarczy Python 3.8+.

## Funkcje

| Funkcja | VBA (Outlook) | Python (ten projekt) |
|---------|---------------|----------------------|
| Wyszukiwanie załączników po słowach kluczowych | ✅ | ✅ |
| Pobieranie maili z serwera | MAPI (tylko Outlook) | IMAP (dowolny serwer) |
| Rekurencyjne przeszukiwanie folderów | ✅ | ✅ |
| Pomijanie obrazów (PNG/JPG/GIF...) | ✅ | ✅ |
| Zapis metadanych (.meta) | ✅ | ✅ |
| Rozwiązywanie konfliktów plików | ✅ | ✅ |
| Konfiguracja | Rejestr Windows | Plik `wi.ini` |
| Interfejs | UserForm (VBA) | CLI + opcjonalnie tkinter |

## Instalacja

```bash
# Python 3.8+ jest wymagany (wbudowany w Windows 10/11)
python --version

# Brak dodatkowych zależności — wszystko w standard library
```

## Szybki start

```bash
# 1. Wygeneruj domyślną konfigurację
python wi_search.py --init-config

# 2. Edytuj wi.ini (base_path, keywords, opcjonalnie IMAP)
notepad wi.ini

# 3. Uruchom wyszukiwanie
python wi_search.py

# 4. Lub z własną konfiguracją i folderem wyjściowym
python wi_search.py -c moj_konfig.ini -o D:\WYNIKI
```

## Konfiguracja (wi.ini)

```ini
[search]
base_path = D:\DANE
min_attachment_size = 51200
recursive_search = True
save_metadata = True

[keywords]
list = umowa,podwykonaw,aneks,zlecenie,kontrakt,draft,wniosek,wdr,faktura,protokol,zamowienie,sprawozdanie,pozwolenie,decyzja

[imap]
; Opcjonalnie — pobieranie maili z dowolnego serwera
host = imap.gmail.com
port = 993
user = twoj@email.pl
password = twoje-haslo-aplikacji
folder = INBOX
date_from = 2026-02-01
date_to =
```

## Struktura projektu

```
WI_LibreOffice/
├── wi_search.py       # Główny silnik wyszukiwania
├── wi_dialog.py       # Prosty interfejs GUI (tkinter)
├── wi.ini             # Plik konfiguracyjny
├── requirements.txt   # Brak zewnętrznych zależności
── README.md          # Dokumentacja
```

## Różnice vs VBA

| Aspekt | VBA | Python |
|--------|-----|--------|
| Dostęp do maili | Tylko Outlook (MAPI) | Dowolny serwer IMAP |
| Konfiguracja | Rejestr Windows (`SaveSetting`) | Plik `wi.ini` (`configparser`) |
| Regex | `VBScript.RegExp` | `re` (standard library) |
| System plików | `Scripting.FileSystemObject` | `pathlib` + `shutil` |
| UI | `UserForm` (VBA) | `tkinter` (Python) |
| Uruchomienie | Z poziomu Outlooka | Dowolny terminal / skrypt |

## Integracja z LibreOffice

Skrypt działa **niezależnie** od LibreOffice. Aby zintegrować:

1. **Makro LibreOffice Basic** — wywołuje `wi_search.py` przez `Shell()`
2. **Python w LibreOffice** — używa `uno` do integracji z dokumentami
3. **Standalone** — uruchamia się z wiersza poleceń (rekomendowane)

## Licencja

Ten sam projekt co WI_Outlook2013 — prywatny, wewnętrzny.
