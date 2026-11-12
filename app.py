"""Window app for people without Python: pick an Excel file, get the charts.

Run with `python app.py`, or build HyroxRadar.exe with build.ps1.
Started with arguments it behaves exactly like radar.py (used to test the .exe).
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from pathlib import Path

APP_TITLE = "Hyrox Radar"
AUTO = "Otomatik"
SCALES = {"En iyinin yüzdesi (önerilen)": "pct-of-best", "Min-maks (en kötü = 0)": "minmax"}
COMPARE = {"Kendi kategorisi içinde": "group", "Herkes tek grupta": "all"}
LANGS = {"Türkçe": "tr", "English": "en"}
AXIS = {"Otomatik (önerilen)": "auto", "0'dan başlat (yakınlaştırma yok)": "zero",
        "Elle gir": "custom"}

# Dark theme, matching the charts.
BG, PANEL, FIELD, LINE = "#0D0E11", "#16181D", "#1F2229", "#2A2E37"
ACCENT, ACCENT_HI = "#FFD60A", "#FFE45C"
FG, MUTED, DISABLED = "#F2F4F7", "#8B93A1", "#565D69"
FONT = "Segoe UI"


def apply_theme(root, style) -> None:
    """Colours and fonts only; no widget behaviour changes here."""
    root.configure(background=BG)
    root.option_add("*TCombobox*Listbox.background", FIELD)
    root.option_add("*TCombobox*Listbox.foreground", FG)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", BG)
    root.option_add("*TCombobox*Listbox.font", (FONT, 10))

    style.configure(".", background=BG, foreground=FG, font=(FONT, 10),
                    bordercolor=LINE, lightcolor=LINE, darkcolor=LINE, focuscolor=ACCENT)
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=PANEL)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Card.TLabel", background=PANEL, foreground=FG)
    style.configure("Caption.TLabel", background=PANEL, foreground=ACCENT,
                    font=(FONT, 8, "bold"))
    style.configure("Count.TLabel", background=BG, foreground=MUTED, font=(FONT, 10, "bold"))

    style.configure("TEntry", fieldbackground=FIELD, foreground=FG, insertcolor=FG,
                    padding=6, relief="flat")
    style.map("TEntry", fieldbackground=[("disabled", PANEL)],
              foreground=[("disabled", DISABLED)],
              bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])
    style.configure("TCombobox", fieldbackground=FIELD, background=FIELD, foreground=FG,
                    arrowcolor=ACCENT, padding=5, relief="flat")
    style.map("TCombobox", fieldbackground=[("readonly", FIELD)],
              foreground=[("readonly", FG)], background=[("active", LINE)],
              selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", FG)],
              bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])

    style.configure("TButton", background=FIELD, foreground=FG, padding=(14, 7),
                    relief="flat", borderwidth=1)
    style.map("TButton", background=[("disabled", PANEL), ("pressed", LINE), ("active", LINE)],
              foreground=[("disabled", DISABLED)])
    style.configure("Accent.TButton", background=ACCENT, foreground=BG,
                    font=(FONT, 11, "bold"), padding=(22, 10), bordercolor=ACCENT,
                    lightcolor=ACCENT, darkcolor=ACCENT)
    style.map("Accent.TButton",
              background=[("disabled", LINE), ("pressed", ACCENT_HI), ("active", ACCENT_HI)],
              foreground=[("disabled", DISABLED)],
              bordercolor=[("disabled", LINE)], lightcolor=[("disabled", LINE)],
              darkcolor=[("disabled", LINE)])
    style.configure("Link.TButton", background=BG, foreground=MUTED, padding=(2, 4),
                    borderwidth=0, font=(FONT, 9))
    style.map("Link.TButton", background=[("active", BG), ("pressed", BG)],
              foreground=[("active", ACCENT)])

    style.configure("TLabelframe", background=PANEL, relief="flat", borderwidth=1,
                    labelmargins=(16, 10, 0, 2))
    style.configure("TLabelframe.Label", background=PANEL, foreground=ACCENT,
                    font=(FONT, 8, "bold"))
    style.configure("Horizontal.TProgressbar", troughcolor=FIELD, background=ACCENT,
                    bordercolor=FIELD, lightcolor=ACCENT, darkcolor=ACCENT, thickness=10)
    style.configure("Vertical.TScrollbar", background=FIELD, troughcolor=BG,
                    arrowcolor=MUTED, bordercolor=BG, lightcolor=FIELD, darkcolor=FIELD)
    style.map("Vertical.TScrollbar", background=[("active", LINE)])


def read_sheet_info(path: str, sheet: str | None = None):
    """Sheet names and the headers after column A, for the drop-down lists."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        names = wb.sheetnames
        ws = wb[sheet] if sheet in names else wb[names[0]]
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers = [str(v).strip() for v in first[1:] if v is not None and str(v).strip()]
        return names, headers
    finally:
        wb.close()


def build_options(file: str, out: str, sheet: str, group: str, compare: str, scale: str,
                  axis_mode: str, axis_value: str, lang: str, dpi: str):
    """Turn the form values into runner Options. Raises ValueError with a Turkish message."""
    from hyrox_radar.runner import Options, parse_axis_min

    if not file.strip():
        raise ValueError("Önce bir Excel dosyası seçin.")
    if not Path(file).is_file():
        raise ValueError(f"Dosya bulunamadı:\n{file}")
    if not out.strip():
        raise ValueError("Grafiklerin kaydedileceği klasörü seçin.")
    mode = AXIS.get(axis_mode, "auto")
    if mode == "auto":
        axis = "auto"
    elif mode == "zero":
        axis = 0.0
    else:
        try:
            axis = parse_axis_min(axis_value)
        except ValueError:
            raise ValueError("Eksen başlangıcı 0 ile 95 arasında bir sayı olmalı.") from None
    try:
        dpi_value = int(dpi)
        if not 50 <= dpi_value <= 600:
            raise ValueError
    except ValueError:
        raise ValueError("Çözünürlük (DPI) 50 ile 600 arasında bir tam sayı olmalı.") from None
    return Options(
        xlsx=file.strip(), out=out.strip(),
        sheet=None if sheet in ("", AUTO) else sheet,
        group_col=None if group in ("", AUTO) else group,
        compare=COMPARE.get(compare, "group"), scale=SCALES.get(scale, "pct-of-best"),
        axis_min=axis, lang=LANGS.get(lang, "tr"), dpi=dpi_value)


def run_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:  # sharp text on high-DPI screens
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title(APP_TITLE)
    root.minsize(760, 620)
    style = ttk.Style()
    try:
        style.theme_use("clam")  # the built-in theme that accepts custom colours
    except tk.TclError:
        pass
    apply_theme(root, style)
    try:  # dark title bar on Windows 10/11
        import ctypes
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        flag = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(flag), 4)
    except Exception:
        pass

    events: queue.Queue = queue.Queue()
    state = {"running": False, "out_dir": None, "auto_out": ""}

    v_file, v_out = tk.StringVar(), tk.StringVar()
    v_sheet, v_group = tk.StringVar(value=AUTO), tk.StringVar(value=AUTO)
    v_compare = tk.StringVar(value=next(iter(COMPARE)))
    v_scale = tk.StringVar(value=next(iter(SCALES)))
    v_axis = tk.StringVar(value=next(iter(AXIS)))
    v_axis_value = tk.StringVar(value="50")
    v_lang = tk.StringVar(value=next(iter(LANGS)))
    v_dpi = tk.StringVar(value="200")

    header = tk.Frame(root, background=PANEL)
    header.pack(fill="x")
    brand = tk.Frame(header, background=PANEL)
    brand.pack(anchor="w", padx=22, pady=(16, 0))
    tk.Label(brand, text="HYROX", background=PANEL, foreground=FG,
             font=(FONT, 20, "bold")).pack(side="left")
    tk.Label(brand, text=" RADAR", background=PANEL, foreground=ACCENT,
             font=(FONT, 20, "bold")).pack(side="left")
    tk.Label(header, text="Excel dosyasını seçin, grafikler katılımcı başına bir PNG "
                          "olarak kaydedilir.", background=PANEL, foreground=MUTED,
             font=(FONT, 10)).pack(anchor="w", padx=24, pady=(0, 14))
    tk.Frame(root, background=ACCENT, height=3).pack(fill="x")

    main = ttk.Frame(root, padding=(22, 18, 22, 18))
    main.pack(fill="both", expand=True)
    main.columnconfigure(0, weight=1)

    card = ttk.Frame(main, style="Card.TFrame", padding=(16, 12, 16, 14))
    card.grid(row=1, column=0, columnspan=3, sticky="ew")
    card.columnconfigure(1, weight=1)
    ttk.Label(card, text="DOSYA", style="Caption.TLabel").grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

    # --- file and output folder -------------------------------------------------
    def refresh_lists(*_):
        path = v_file.get().strip()
        if not Path(path).is_file():
            return
        try:
            sheets, headers = read_sheet_info(path, v_sheet.get())
        except Exception as exc:
            log(f"Dosya okunamadı: {exc}", "error")
            return
        cb_sheet["values"] = [AUTO] + sheets
        if v_sheet.get() not in cb_sheet["values"]:
            v_sheet.set(AUTO)
        cb_group["values"] = [AUTO] + headers
        if v_group.get() not in cb_group["values"]:
            v_group.set(AUTO)

    def choose_file():
        path = filedialog.askopenfilename(
            title="Excel dosyasını seçin",
            filetypes=[("Excel dosyası", "*.xlsx *.xlsm"), ("Tüm dosyalar", "*.*")])
        if not path:
            return
        v_file.set(os.path.normpath(path))
        # Suggest an output folder next to the file unless the user picked one.
        if not v_out.get().strip() or v_out.get() == state["auto_out"]:
            state["auto_out"] = str(Path(path).parent / "grafikler")
            v_out.set(state["auto_out"])
        v_sheet.set(AUTO)
        v_group.set(AUTO)
        refresh_lists()

    def choose_out():
        path = filedialog.askdirectory(title="Grafiklerin kaydedileceği klasör")
        if path:
            v_out.set(os.path.normpath(path))

    ttk.Label(card, text="Excel dosyası", style="Card.TLabel").grid(row=1, column=0, sticky="w")
    ttk.Entry(card, textvariable=v_file).grid(row=1, column=1, sticky="ew", padx=10, pady=4)
    ttk.Button(card, text="Seç...", command=choose_file).grid(row=1, column=2)
    ttk.Label(card, text="Çıktı klasörü", style="Card.TLabel").grid(row=2, column=0, sticky="w")
    ttk.Entry(card, textvariable=v_out).grid(row=2, column=1, sticky="ew", padx=10, pady=4)
    ttk.Button(card, text="Seç...", command=choose_out).grid(row=2, column=2)

    # --- advanced settings (collapsed by default) -------------------------------
    adv = ttk.LabelFrame(main, text="GELİŞMİŞ AYARLAR", padding=(16, 8, 4, 12))
    adv.columnconfigure(1, weight=1)
    adv.columnconfigure(3, weight=1)

    def combo(parent, var, values, row, col, label, **kw):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row, column=col, sticky="w", pady=4)
        box = ttk.Combobox(parent, textvariable=var, values=values, state="readonly", **kw)
        box.grid(row=row, column=col + 1, sticky="ew", padx=(10, 16), pady=4)
        return box

    cb_sheet = combo(adv, v_sheet, [AUTO], 0, 0, "Sayfa")
    cb_group = combo(adv, v_group, [AUTO], 0, 2, "Kategori sütunu")
    combo(adv, v_compare, list(COMPARE), 1, 0, "Karşılaştırma")
    combo(adv, v_scale, list(SCALES), 1, 2, "Puan ölçeği")
    cb_axis = combo(adv, v_axis, list(AXIS), 2, 0, "Eksen başlangıcı")
    axis_entry = ttk.Entry(adv, textvariable=v_axis_value, width=8)
    ttk.Label(adv, text="Değer (0-95)", style="Card.TLabel").grid(row=2, column=2, sticky="w")
    axis_entry.grid(row=2, column=3, sticky="w", padx=(10, 16))
    combo(adv, v_lang, list(LANGS), 3, 0, "Grafik dili")
    ttk.Label(adv, text="Çözünürlük (DPI)", style="Card.TLabel").grid(
        row=3, column=2, sticky="w")
    ttk.Entry(adv, textvariable=v_dpi, width=8).grid(row=3, column=3, sticky="w", padx=(10, 16))

    def sync_axis(*_):
        axis_entry.configure(state="normal" if AXIS.get(v_axis.get()) == "custom"
                             else "disabled")

    cb_axis.bind("<<ComboboxSelected>>", sync_axis)
    cb_sheet.bind("<<ComboboxSelected>>", refresh_lists)
    sync_axis()

    shown = {"adv": False}

    def toggle_adv():
        shown["adv"] = not shown["adv"]
        if shown["adv"]:
            adv.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(6, 0))
            btn_adv.configure(text="▾  Gelişmiş ayarları gizle")
        else:
            adv.grid_remove()
            btn_adv.configure(text="▸  Gelişmiş ayarlar")

    btn_adv = ttk.Button(main, text="▸  Gelişmiş ayarlar", command=toggle_adv,
                         style="Link.TButton", cursor="hand2")
    btn_adv.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

    # --- run row -----------------------------------------------------------------
    bar = ttk.Frame(main)
    bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(14, 10))
    bar.columnconfigure(1, weight=1)
    btn_run = ttk.Button(bar, text="Grafikleri oluştur", style="Accent.TButton",
                         cursor="hand2")
    btn_run.grid(row=0, column=0)
    progress = ttk.Progressbar(bar, mode="determinate")
    progress.grid(row=0, column=1, sticky="ew", padx=(14, 10))
    v_count = tk.StringVar()
    ttk.Label(bar, textvariable=v_count, style="Count.TLabel", width=8, anchor="e").grid(
        row=0, column=2, padx=(0, 12))
    btn_open = ttk.Button(bar, text="Klasörü aç", state="disabled")
    btn_open.grid(row=0, column=3)

    # --- log -----------------------------------------------------------------------
    box = ttk.Frame(main)
    box.grid(row=6, column=0, columnspan=3, sticky="nsew")
    main.rowconfigure(6, weight=1)
    box.columnconfigure(0, weight=1)
    box.rowconfigure(0, weight=1)
    text = tk.Text(box, height=14, wrap="word", state="disabled", font=("Consolas", 10),
                   background=PANEL, foreground="#C9CFD8", relief="flat", borderwidth=0,
                   padx=12, pady=10, highlightthickness=1, highlightbackground=LINE,
                   highlightcolor=LINE, selectbackground=ACCENT, selectforeground=BG,
                   spacing1=2)
    scroll = ttk.Scrollbar(box, command=text.yview)
    text.configure(yscrollcommand=scroll.set)
    text.grid(row=0, column=0, sticky="nsew")
    scroll.grid(row=0, column=1, sticky="ns")
    text.tag_configure("warn", foreground="#FFB454")
    text.tag_configure("error", foreground="#FF6B7A")
    text.tag_configure("ok", foreground="#3DDC84")

    def log(message: str, tag: str = ""):
        text.configure(state="normal")
        text.insert("end", message + "\n", tag)
        text.see("end")
        text.configure(state="disabled")

    # --- work in a thread, report through the queue -------------------------------
    def worker(opts):
        from hyrox_radar.loader import InputError
        from hyrox_radar.runner import run
        try:
            result = run(opts,
                         info=lambda m: events.put(("info", m)),
                         warn=lambda m: events.put(("warn", m)),
                         progress=lambda done, total: events.put(("progress", (done, total))))
            events.put(("done", result))
        except InputError as exc:
            events.put(("error", str(exc)))
        except PermissionError as exc:
            events.put(("error", f"Dosyaya yazılamadı. Açık bir grafik ya da Excel dosyası "
                                 f"varsa kapatın.\n{exc}"))
        except Exception:
            events.put(("error", "Beklenmeyen hata:\n" + traceback.format_exc()))

    def start():
        if state["running"]:
            return
        try:
            opts = build_options(v_file.get(), v_out.get(), v_sheet.get(), v_group.get(),
                                 v_compare.get(), v_scale.get(), v_axis.get(),
                                 v_axis_value.get(), v_lang.get(), v_dpi.get())
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        text.configure(state="normal")
        text.delete("1.0", "end")
        text.configure(state="disabled")
        state.update(running=True, out_dir=None)
        btn_run.configure(state="disabled")
        btn_open.configure(state="disabled")
        progress.configure(value=0, maximum=1)
        v_count.set("")
        log("Dosya okunuyor...")
        threading.Thread(target=worker, args=(opts,), daemon=True).start()

    def finish():
        state["running"] = False
        btn_run.configure(state="normal")

    def poll():
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "info":
                    log(payload)
                elif kind == "warn":
                    log("UYARI: " + payload, "warn")
                elif kind == "progress":
                    done, total = payload
                    progress.configure(maximum=total, value=done)
                    v_count.set(f"{done} / {total}")
                elif kind == "done":
                    state["out_dir"] = str(payload.out_dir)
                    log(f"Tamamlandı: {payload.written} grafik kaydedildi.", "ok")
                    btn_open.configure(state="normal")
                    finish()
                elif kind == "error":
                    log("HATA: " + payload, "error")
                    finish()
        except queue.Empty:
            pass
        root.after(80, poll)

    def open_folder():
        if state["out_dir"] and Path(state["out_dir"]).is_dir():
            os.startfile(state["out_dir"])  # noqa: S606 - opens Explorer on Windows

    btn_run.configure(command=start)
    btn_open.configure(command=open_folder)

    # Hooks for the automated smoke test.
    root.app = {"file": v_file, "out": v_out, "start": start, "state": state,
                "refresh": refresh_lists, "text": text, "group_box": cb_group,
                "sheet_box": cb_sheet, "toggle_adv": toggle_adv, "poll": poll}
    if os.environ.get("HYROX_GUI_SMOKE"):
        return root  # the test drives the event loop itself
    poll()
    root.mainloop()
    return 0


def main() -> int:
    if len(sys.argv) > 1:
        from radar import main as cli_main
        return cli_main()
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())
