"""
WI Search — Prosty interfejs GUI (tkinter)
Uruchamia wi_search.py z parametrami z formularza.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import sys
from pathlib import Path

CONFIG_FILE = Path(__file__).parent / "wi.ini"


class WISearchDialog(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WI Search — Wyszukiwarka Korespondencji")
        self.geometry("520x420")
        self.resizable(False, False)

        # --- Konfiguracja ---
        self.config = self._load_config()

        # --- Widgety ---
        self._build_ui()

    def _load_config(self) -> dict:
        if CONFIG_FILE.exists():
            import configparser
            cfg = configparser.ConfigParser()
            cfg.read(CONFIG_FILE, encoding="utf-8")
            return {
                "base_path": cfg.get("search", "base_path", fallback=r"D:\DANE"),
                "min_size": cfg.getint("search", "min_attachment_size", fallback=51200),
                "recursive": cfg.getboolean("search", "recursive_search", fallback=True),
                "metadata": cfg.getboolean("search", "save_metadata", fallback=True),
                "keywords": cfg.get("keywords", "list", fallback=""),
                "imap_host": cfg.get("imap", "host", fallback=""),
                "imap_user": cfg.get("imap", "user", fallback=""),
            }
        return {
            "base_path": r"D:\DANE",
            "min_size": 51200,
            "recursive": True,
            "metadata": True,
            "keywords": "",
            "imap_host": "",
            "imap_user": "",
        }

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- Folder bazowy ---
        ttk.Label(self, text="Folder bazowy (serwer):").grid(row=0, column=0, sticky="w", **pad)
        self.var_path = tk.StringVar(value=self.config["base_path"])
        ttk.Entry(self, textvariable=self.var_path, width=40).grid(row=0, column=1, **pad)
        ttk.Button(self, text="📁", width=3, command=self._browse_path).grid(row=0, column=2, **pad)

        # --- Folder wyjściowy ---
        ttk.Label(self, text="Folder na wyniki:").grid(row=1, column=0, sticky="w", **pad)
        self.var_output = tk.StringVar(value=str(Path.home() / "Desktop" / "WI_wyniki"))
        ttk.Entry(self, textvariable=self.var_output, width=40).grid(row=1, column=1, **pad)
        ttk.Button(self, text="📁", width=3, command=self._browse_output).grid(row=1, column=2, **pad)

        # --- Słowa kluczowe ---
        ttk.Label(self, text="Słowa kluczowe (oddzielone przecinkiem):").grid(row=2, column=0, columnspan=2, sticky="w", **pad)
        self.var_keywords = tk.StringVar(value=self.config["keywords"])
        ttk.Entry(self, textvariable=self.var_keywords, width=55).grid(row=3, column=0, columnspan=3, **pad)

        # --- Min rozmiar ---
        ttk.Label(self, text="Min. rozmiar załącznika (B):").grid(row=4, column=0, sticky="w", **pad)
        self.var_min_size = tk.StringVar(value=str(self.config["min_size"]))
        ttk.Entry(self, textvariable=self.var_min_size, width=10).grid(row=4, column=1, sticky="w", **pad)

        # --- Checkboxy ---
        self.var_recursive = tk.BooleanVar(value=self.config["recursive"])
        ttk.Checkbutton(self, text="Przeszukuj podfoldery (rekurencyjnie)", variable=self.var_recursive).grid(row=5, column=0, columnspan=3, sticky="w", **pad)

        self.var_metadata = tk.BooleanVar(value=self.config["metadata"])
        ttk.Checkbutton(self, text="Zapisuj metadane (.meta)", variable=self.var_metadata).grid(row=6, column=0, columnspan=3, sticky="w", **pad)

        # --- Separator ---
        ttk.Separator(self, orient="horizontal").grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        # --- IMAP (opcjonalnie) ---
        ttk.Label(self, text="IMAP (opcjonalnie — puste = pomijam maile):").grid(row=8, column=0, columnspan=3, sticky="w", **pad)
        self.var_imap_host = tk.StringVar(value=self.config["imap_host"])
        ttk.Entry(self, textvariable=self.var_imap_host, width=20, placeholder="imap.gmail.com").grid(row=9, column=0, sticky="w", **pad)
        self.var_imap_user = tk.StringVar(value=self.config["imap_user"])
        ttk.Entry(self, textvariable=self.var_imap_user, width=20, placeholder="user@email.pl").grid(row=9, column=1, sticky="w", **pad)
        ttk.Label(self, text="(hasło podaj w wi.ini)").grid(row=9, column=2, sticky="w", **pad)

        # --- Przyciski ---
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=10, column=0, columnspan=3, pady=15)

        ttk.Button(btn_frame, text=" Szukaj", width=15, command=self._run_search).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="💾 Zapisz konfigurację", width=18, command=self._save_config).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ Zamknij", width=12, command=self.destroy).pack(side="left", padx=5)

        # --- Status ---
        self.var_status = tk.StringVar(value="Gotowe do wyszukiwania.")
        ttk.Label(self, textvariable=self.var_status, foreground="gray").grid(row=11, column=0, columnspan=3, **pad)

    def _browse_path(self):
        d = filedialog.askdirectory(initialdir=self.var_path.get())
        if d:
            self.var_path.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(initialdir=self.var_output.get())
        if d:
            self.var_output.set(d)

    def _save_config(self):
        import configparser
        cfg = configparser.ConfigParser()
        cfg["search"] = {
            "base_path": self.var_path.get(),
            "min_attachment_size": self.var_min_size.get(),
            "recursive_search": str(self.var_recursive.get()),
            "save_metadata": str(self.var_metadata.get()),
        }
        cfg["keywords"] = {"list": self.var_keywords.get()}
        if self.var_imap_host.get():
            cfg["imap"] = {
                "host": self.var_imap_host.get(),
                "port": "993",
                "user": self.var_imap_user.get(),
                "folder": "INBOX",
            }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)
        self.var_status.set("✅ Konfiguracja zapisana do wi.ini")

    def _run_search(self):
        # Zapisz konfigurację tymczasową
        import configparser
        cfg = configparser.ConfigParser()
        cfg["search"] = {
            "base_path": self.var_path.get(),
            "min_attachment_size": self.var_min_size.get(),
            "recursive_search": str(self.var_recursive.get()),
            "save_metadata": str(self.var_metadata.get()),
        }
        cfg["keywords"] = {"list": self.var_keywords.get()}
        if self.var_imap_host.get():
            cfg["imap"] = {
                "host": self.var_imap_host.get(),
                "port": "993",
                "user": self.var_imap_user.get(),
                "folder": "INBOX",
            }
        tmp_cfg = Path(__file__).parent / "_tmp_config.ini"
        with open(tmp_cfg, "w", encoding="utf-8") as f:
            cfg.write(f)

        self.var_status.set("⏳ Uruchamiam wyszukiwanie...")
        self.update()

        try:
            proc = subprocess.run(
                [sys.executable, str(Path(__file__).parent / "wi_search.py"),
                 "-c", str(tmp_cfg), "-o", self.var_output.get()],
                capture_output=True, text=True, timeout=600,
            )
            if proc.returncode == 0:
                self.var_status.set("✅ Zakończono! Sprawdź folder wyników.")
                messagebox.showinfo("WI Search", proc.stdout)
            else:
                self.var_status.set("❌ Błąd — sprawdź konsolę.")
                messagebox.showerror("Błąd", proc.stderr)
        except subprocess.TimeoutExpired:
            self.var_status.set("⏱️ Przekroczono limit czasu (10 min).")
        except Exception as e:
            self.var_status.set(f"❌ {e}")
        finally:
            if tmp_cfg.exists():
                tmp_cfg.unlink()


if __name__ == "__main__":
    app = WISearchDialog()
    app.mainloop()
