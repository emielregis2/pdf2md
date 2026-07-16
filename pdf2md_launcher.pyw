#!/usr/bin/env python3
"""
pdf2md_launcher.pyw
---------------------
Launcher GUI dla konwertera PDF -> Markdown, przeznaczony do wywoływania
z menu kontekstowego Eksploratora Windows (prawy klawisz na pliku .pdf).

Rozszerzenie .pyw sprawia, że skrypt uruchamia się przez pythonw.exe,
czyli BEZ okna konsoli w tle - użytkownik widzi tylko małe okienko postępu.

Sposób działania:
    pythonw.exe pdf2md_launcher.pyw "C:\\ścieżka\\do\\dokumentu.pdf"

Po zakończeniu:
    - w tym samym katalogu powstaje dokumentu.md,
    - okno pokazuje log konwersji i przycisk "Otwórz plik wynikowy".

Ten plik NIE zawiera logiki konwersji - importuje ją z pdf_to_markdown_universal.py,
który musi leżeć w tym samym katalogu.
"""

from __future__ import annotations

import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import scrolledtext, messagebox

# Katalog, w którym leży ten plik - stąd importujemy właściwy konwerter
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

try:
    from pdf_to_markdown_universal import convert_pdf, ConversionConfig
except Exception as exc:  # noqa: BLE001
    # Jeśli brakuje zależności / modułu, pokaż to w oknie zamiast wywalać się bez śladu
    import tkinter.messagebox as mb
    root = tk.Tk()
    root.withdraw()
    mb.showerror(
        "pdf2md - błąd startu",
        "Nie udało się załadować modułu konwertera "
        f"(pdf_to_markdown_universal.py):\n\n{exc}\n\n"
        "Sprawdź, czy wszystkie zależności są zainstalowane (patrz README_INSTALACJA_WINDOWS.md).",
    )
    sys.exit(1)


class QueueWriter:
    """Podstawia się pod sys.stdout/sys.stderr i przekazuje tekst do wątku GUI przez kolejkę."""

    def __init__(self, q: "queue.Queue[str]"):
        self.q = q

    def write(self, text: str) -> None:
        if text:
            self.q.put(text)

    def flush(self) -> None:
        pass


class Pdf2MdWindow(tk.Tk):
    def __init__(self, pdf_path: Path):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_path: Path | None = None
        self.finished = False
        self.failed = False

        self.title("pdf2md - konwersja do Markdown")
        self.geometry("640x420")
        self.minsize(480, 300)

        label = tk.Label(
            self, text=f"Konwertuję: {pdf_path.name}", font=("Segoe UI", 10, "bold"), anchor="w"
        )
        label.pack(fill="x", padx=10, pady=(10, 4))

        self.log_widget = scrolledtext.ScrolledText(
            self, wrap="word", font=("Consolas", 9), state="disabled", bg="#111", fg="#ddd"
        )
        self.log_widget.pack(fill="both", expand=True, padx=10, pady=4)

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=(4, 10))

        self.open_btn = tk.Button(
            btn_frame, text="Otwórz plik wynikowy", command=self._open_result, state="disabled"
        )
        self.open_btn.pack(side="left")

        self.close_btn = tk.Button(btn_frame, text="Zamknij", command=self.destroy)
        self.close_btn.pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self.destroy)

        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.after(80, self._poll_queue)

        self.worker = threading.Thread(target=self._run_conversion, daemon=True)
        self.worker.start()

    # ------------------------------------------------------------------

    def _append_log(self, text: str) -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", text)
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _poll_queue(self) -> None:
        try:
            while True:
                text = self.log_queue.get_nowait()
                self._append_log(text)
        except queue.Empty:
            pass

        if self.finished:
            if self.failed:
                self.title("pdf2md - BŁĄD konwersji")
            else:
                self.title("pdf2md - gotowe ✔")
                self.open_btn.configure(state="normal")
        else:
            self.after(80, self._poll_queue)

    def _run_conversion(self) -> None:
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = QueueWriter(self.log_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            config = ConversionConfig(lang="pol+eng", dpi=300)
            out_path = convert_pdf(self.pdf_path, config=config, show_stats=True)
            self.output_path = out_path
            self.log_queue.put("\n=== KONWERSJA ZAKOŃCZONA POMYŚLNIE ===\n")
        except Exception:  # noqa: BLE001
            self.failed = True
            self.log_queue.put("\n=== BŁĄD KONWERSJI ===\n")
            self.log_queue.put(traceback.format_exc())
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr
            self.finished = True

    def _open_result(self) -> None:
        if self.output_path and self.output_path.exists():
            import os
            os.startfile(str(self.output_path))  # noqa: S606 - celowe, to jest funkcja Windows


def main() -> None:
    if len(sys.argv) < 2:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "pdf2md",
            "Nie podano ścieżki do pliku PDF.\n\n"
            "Ten program jest przeznaczony do uruchamiania z menu kontekstowego "
            "Eksploratora Windows (prawy klawisz na pliku .pdf -> pdf2md).",
        )
        sys.exit(1)

    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("pdf2md", f"To nie jest prawidłowy plik PDF:\n{pdf_path}")
        sys.exit(1)

    app = Pdf2MdWindow(pdf_path)
    app.mainloop()


if __name__ == "__main__":
    main()
