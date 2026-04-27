"""
WedstrijdCalculator GUI
=======================
Moderne grafische interface voor de Portsmouth Yardstick scoring tool.

Verbeteringen v2:
- Naam van de wedstrijd invulbaar, gebruikt in export en titels
- Detail-tab gegroepeerd per reeks met duidelijke scheiding
- Invoertab: wedstrijd aanmaken, deelnemers beheren, reeksen + tijden invoeren

Vereisten:
    pip install customtkinter pandas openpyxl

Starten:
    python wedstrijd_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from pathlib import Path
import threading
import pandas as pd

from wedstrijd_calculator import (
    load_boat_py_table,
    load_race_data,
    calculate_elapsed_seconds,
    calculate_corrected_time_py,
    rank_each_race,
    calculate_points,
    drop_worst_result,
    generate_summary_tables,
    export_to_excel,
)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

CLR_PRIMARY = "#1F4E79"
CLR_ACCENT  = "#2E75B6"

# Row highlight colors — background / foreground pairs for dark and light mode.
# Always pick combinations with high contrast (WCAG AA minimum 4.5:1).
ROW_COLORS = {
    # tag        dark_bg      dark_fg      light_bg     light_fg
    "gold":   ("#7A6000",  "#FFE680",  "#FFF3B0",  "#5A4000"),
    "silver": ("#4A5568",  "#E2E8F0",  "#E2E8F0",  "#2D3748"),
    "bronze": ("#7A4010",  "#FFD0A0",  "#FFE8D0",  "#6B3010"),
    "odd":    ("#323232",  "#E8E8E8",  "#EDF2FB",  "#1A1A2E"),
    "even":   ("#2B2B2B",  "#E8E8E8",  "#FFFFFF",  "#1A1A2E"),
}


# ==============================================================================
# Herbruikbare gestijlde Treeview-tabel
# ==============================================================================

class DataTable(tk.Frame):
    def __init__(self, parent, columns: list[str], row_height: int = 30, **kw):
        dark = ctk.get_appearance_mode() == "Dark"
        super().__init__(parent, bg="#2b2b2b" if dark else "#f4f6fa", **kw)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._apply_style(dark, row_height)
        self.tv = ttk.Treeview(self, columns=columns, show="headings",
                               style="NWV.Treeview")
        vsb = ttk.Scrollbar(self, orient="vertical",   command=self.tv.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self._dark = dark
        self._setup_tags()
        for col in columns:
            self.tv.heading(col, text=col)
            self.tv.column(col, width=100, anchor="center", minwidth=40)

    @staticmethod
    def _apply_style(dark: bool, row_height: int = 30):
        s = ttk.Style()
        # "clam" is required on Windows: the default vista/winnative theme ignores
        # tag foreground colors completely, making text unreadable.
        s.theme_use("clam")
        bg = "#2b2b2b" if dark else "#f4f6fa"
        fg = "#E8E8E8" if dark else "#1a1a1a"
        s.configure("NWV.Treeview", background=bg, foreground=fg,
                    rowheight=row_height, fieldbackground=bg,
                    borderwidth=0, font=("Segoe UI", 11))
        s.configure("NWV.Treeview.Heading", background=CLR_PRIMARY,
                    foreground="#ffffff", font=("Segoe UI", 11, "bold"),
                    relief="flat", padding=8)
        s.map("NWV.Treeview",
              background=[("selected", CLR_ACCENT)],
              foreground=[("selected", "#ffffff")])

    def _setup_tags(self):
        d = self._dark
        idx = 0 if d else 2   # dark_bg=0, dark_fg=1, light_bg=2, light_fg=3
        for tag, vals in ROW_COLORS.items():
            self.tv.tag_configure(tag, background=vals[idx], foreground=vals[idx + 1])
        self.tv.tag_configure("header", background=CLR_ACCENT,
                              foreground="#ffffff",
                              font=("Segoe UI", 11, "bold"))

    def set_columns(self, columns: list[str], widths: dict | None = None):
        widths = widths or {}
        self.tv["columns"] = columns
        for col in columns:
            w = widths.get(col, 90)
            a = "w" if col in ("Naam", "Boottype") else "center"
            self.tv.heading(col, text=col)
            self.tv.column(col, width=w, anchor=a, minwidth=40)

    def clear(self):
        self.tv.delete(*self.tv.get_children())

    def add_row(self, values: list, tag: str = "even"):
        self.tv.insert("", "end", values=values, tags=(tag,))


# ==============================================================================
# Invoer-tab: deelnemers + reeksen + tijden beheren
# ==============================================================================

class InvoerTab(ctk.CTkFrame):
    """
    Drie subpanels:
      Links  - deelnemerslijst (naam + boottype)
      Midden - reekslijst
      Rechts - tijden per deelnemer voor geselecteerde reeks
    """

    def __init__(self, parent, py_table: dict, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.py_table = py_table
        self._deelnemers: list[dict] = []
        self._reeksen: list[int]     = []
        self._tijden: dict           = {}
        self._selected_reeks: int | None = None

        self.grid_columnconfigure((0, 1, 2), weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_deelnemers_panel()
        self._build_reeksen_panel()
        self._build_tijden_panel()

    def _build_deelnemers_panel(self):
        fr = ctk.CTkFrame(self)
        fr.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=0)
        fr.grid_rowconfigure(4, weight=1)
        fr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(fr, text="Deelnemers",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        f1 = ctk.CTkFrame(fr, fg_color="transparent")
        f1.grid(row=1, column=0, sticky="ew", padx=10, pady=2)
        f1.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(f1, text="Naam", width=55).grid(row=0, column=0, sticky="w")
        self._inp_naam = ctk.CTkEntry(f1, placeholder_text="bijv. Stef")
        self._inp_naam.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        f2 = ctk.CTkFrame(fr, fg_color="transparent")
        f2.grid(row=2, column=0, sticky="ew", padx=10, pady=2)
        f2.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(f2, text="Boot", width=55).grid(row=0, column=0, sticky="w")
        boot_opties = sorted(self.py_table.keys())
        self._boot_var = tk.StringVar(value=boot_opties[0] if boot_opties else "")
        self._boot_dd = ctk.CTkOptionMenu(f2, variable=self._boot_var,
                                          values=boot_opties, width=160)
        self._boot_dd.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        btn_row = ctk.CTkFrame(fr, fg_color="transparent")
        btn_row.grid(row=3, column=0, sticky="ew", padx=10, pady=6)
        ctk.CTkButton(btn_row, text="+ Toevoegen", width=110,
                      command=self._add_deelnemer).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row, text="Verwijder", width=100,
                      fg_color="transparent", border_width=1,
                      command=self._del_deelnemer).pack(side="left")

        list_fr = tk.Frame(fr, bg="#2b2b2b")
        list_fr.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_fr.grid_rowconfigure(0, weight=1)
        list_fr.grid_columnconfigure(0, weight=1)

        self._dl_tv = ttk.Treeview(list_fr, columns=("naam", "boot", "py"),
                                   show="headings", style="NWV.Treeview",
                                   selectmode="browse")
        self._dl_tv.heading("naam", text="Naam")
        self._dl_tv.heading("boot", text="Boottype")
        self._dl_tv.heading("py",   text="PY")
        self._dl_tv.column("naam", width=100, anchor="w")
        self._dl_tv.column("boot", width=120, anchor="w")
        self._dl_tv.column("py",   width=55,  anchor="center")
        vsb = ttk.Scrollbar(list_fr, orient="vertical", command=self._dl_tv.yview)
        self._dl_tv.configure(yscrollcommand=vsb.set)
        self._dl_tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

    def _build_reeksen_panel(self):
        fr = ctk.CTkFrame(self)
        fr.grid(row=0, column=1, sticky="nsew", padx=6, pady=0)
        fr.grid_rowconfigure(2, weight=1)
        fr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(fr, text="Reeksen",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 8))

        btn_row = ctk.CTkFrame(fr, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))
        ctk.CTkButton(btn_row, text="+ Reeks toevoegen", width=140,
                      command=self._add_reeks).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row, text="Verwijder", width=100,
                      fg_color="transparent", border_width=1,
                      command=self._del_reeks).pack(side="left")

        list_fr = tk.Frame(fr, bg="#2b2b2b")
        list_fr.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_fr.grid_rowconfigure(0, weight=1)
        list_fr.grid_columnconfigure(0, weight=1)

        self._rk_lb = tk.Listbox(list_fr, bg="#2b2b2b", fg="#e0e0e0",
                                 font=("Segoe UI", 12), selectbackground=CLR_ACCENT,
                                 activestyle="none", relief="flat", bd=0,
                                 highlightthickness=0)
        vsb = ttk.Scrollbar(list_fr, orient="vertical", command=self._rk_lb.yview)
        self._rk_lb.configure(yscrollcommand=vsb.set)
        self._rk_lb.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._rk_lb.bind("<<ListboxSelect>>", self._on_reeks_select)

    def _build_tijden_panel(self):
        self._tijden_fr = ctk.CTkFrame(self)
        self._tijden_fr.grid(row=0, column=2, sticky="nsew", padx=(6, 0), pady=0)
        self._tijden_fr.grid_columnconfigure(0, weight=1)
        self._tijden_fr.grid_rowconfigure(1, weight=1)

        self._tijden_title = ctk.CTkLabel(
            self._tijden_fr, text="Selecteer een reeks",
            font=ctk.CTkFont(size=14, weight="bold"))
        self._tijden_title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))

        self._tijden_scroll = ctk.CTkScrollableFrame(self._tijden_fr)
        self._tijden_scroll.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self._tijden_scroll.grid_columnconfigure(1, weight=1)
        self._tijd_widgets: dict[str, tuple] = {}

    def _refresh_tijden_panel(self):
        for w in self._tijden_scroll.winfo_children():
            w.destroy()
        self._tijd_widgets = {}

        if self._selected_reeks is None:
            self._tijden_title.configure(text="Selecteer een reeks")
            return

        self._tijden_title.configure(text=f"Tijden - Reeks {self._selected_reeks}")

        r = 0
        for col, txt in enumerate(["Deelnemer", "Boot", "Min", "Sec"]):
            ctk.CTkLabel(self._tijden_scroll, text=txt,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="gray").grid(
                row=r, column=col, padx=8, pady=(0, 6), sticky="w")
        r += 1

        for dl in self._deelnemers:
            naam  = dl["naam"]
            boot  = dl["boottype"]
            store = self._tijden.setdefault(self._selected_reeks, {})
            if naam not in store:
                store[naam] = [tk.StringVar(value=""), tk.StringVar(value="")]
            min_var, sec_var = store[naam]

            ctk.CTkLabel(self._tijden_scroll, text=naam,
                         font=ctk.CTkFont(size=12), anchor="w").grid(
                row=r, column=0, padx=8, pady=4, sticky="w")
            ctk.CTkLabel(self._tijden_scroll, text=boot,
                         font=ctk.CTkFont(size=11), text_color="gray",
                         anchor="w").grid(row=r, column=1, padx=4, pady=4, sticky="w")
            ctk.CTkEntry(self._tijden_scroll, textvariable=min_var,
                         width=64, placeholder_text="min").grid(
                row=r, column=2, padx=4, pady=4)
            ctk.CTkEntry(self._tijden_scroll, textvariable=sec_var,
                         width=64, placeholder_text="sec").grid(
                row=r, column=3, padx=4, pady=4)
            self._tijd_widgets[naam] = (min_var, sec_var)
            r += 1

    def _add_deelnemer(self):
        naam = self._inp_naam.get().strip()
        boot = self._boot_var.get().strip()
        if not naam:
            messagebox.showwarning("Naam vereist", "Vul een naam in voor de deelnemer.")
            return
        if any(d["naam"].lower() == naam.lower() for d in self._deelnemers):
            messagebox.showwarning("Duplicaat", f"'{naam}' staat al in de lijst.")
            return
        self._deelnemers.append({"naam": naam, "boottype": boot})
        py = self.py_table.get(boot, "?")
        self._dl_tv.insert("", "end", iid=naam, values=(naam, boot, py))
        self._inp_naam.delete(0, "end")
        self._refresh_tijden_panel()

    def _del_deelnemer(self):
        sel = self._dl_tv.selection()
        if not sel:
            return
        naam = sel[0]
        self._deelnemers = [d for d in self._deelnemers if d["naam"] != naam]
        self._dl_tv.delete(naam)
        for r in self._tijden.values():
            r.pop(naam, None)
        self._refresh_tijden_panel()

    def _add_reeks(self):
        nr = len(self._reeksen) + 1
        self._reeksen.append(nr)
        self._rk_lb.insert("end", f"  Reeks {nr}")
        self._rk_lb.selection_clear(0, "end")
        self._rk_lb.selection_set("end")
        self._rk_lb.event_generate("<<ListboxSelect>>")

    def _del_reeks(self):
        sel = self._rk_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        nr  = self._reeksen[idx]
        self._reeksen.pop(idx)
        self._tijden.pop(nr, None)
        self._rk_lb.delete(idx)
        self._selected_reeks = None
        self._refresh_tijden_panel()

    def _on_reeks_select(self, _evt=None):
        sel = self._rk_lb.curselection()
        if not sel:
            return
        self._selected_reeks = self._reeksen[sel[0]]
        self._refresh_tijden_panel()

    def get_dataframe(self) -> pd.DataFrame:
        rows = []
        for reeks in self._reeksen:
            for dl in self._deelnemers:
                naam = dl["naam"]
                boot = dl["boottype"]
                store = self._tijden.get(reeks, {})
                if naam in store:
                    min_v, sec_v = store[naam]
                    min_s = min_v.get().strip()
                    sec_s = sec_v.get().strip()
                    if min_s or sec_s:
                        rows.append({
                            "naam": naam, "boottype": boot,
                            "reeks": reeks,
                            "minuten":  int(min_s) if min_s else 0,
                            "seconden": int(sec_s) if sec_s else 0,
                        })
        if not rows:
            raise ValueError("Geen tijden ingevoerd. Vul minstens een reeks in.")
        return pd.DataFrame(rows)

    def load_demo(self, demo_rows: list, py_table: dict):
        self._deelnemers.clear()
        self._reeksen.clear()
        self._tijden.clear()
        self._dl_tv.delete(*self._dl_tv.get_children())
        self._rk_lb.delete(0, "end")
        self._selected_reeks = None

        for naam, boottype, reeks, min_, sec in demo_rows:
            if not any(d["naam"] == naam for d in self._deelnemers):
                self._deelnemers.append({"naam": naam, "boottype": boottype})
                py = py_table.get(boottype, "?")
                self._dl_tv.insert("", "end", iid=naam,
                                   values=(naam, boottype, py))
            if reeks not in self._reeksen:
                self._reeksen.append(reeks)
                self._rk_lb.insert("end", f"  Reeks {reeks}")
            store = self._tijden.setdefault(reeks, {})
            store[naam] = [tk.StringVar(value=str(min_)),
                           tk.StringVar(value=str(sec))]

        self._refresh_tijden_panel()


# ==============================================================================
# Hoofd-applicatievenster
# ==============================================================================

class WedstrijdApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("WedstrijdCalculator  -  Portsmouth Yardstick Scoring")
        self.geometry("1400x860")
        self.minsize(1100, 700)

        self.py_table        = load_boat_py_table()
        self.invoer_pad      = None
        self.detail_df       = None
        self.samenvatting_df = None
        self._data_source    = "manual"

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()
        self._build_statusbar()

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=256, corner_radius=0)
        sb.grid(row=0, column=0, rowspan=2, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(99, weight=1)

        r = 0
        ctk.CTkLabel(sb, text="⛵", font=ctk.CTkFont(size=46)).grid(
            row=r, column=0, pady=(32, 0)); r += 1
        ctk.CTkLabel(sb, text="Wedstrijd\nCalculator",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     justify="center").grid(row=r, column=0, pady=(4, 20)); r += 1

        self._sb_section(sb, "WEDSTRIJD", r); r += 1
        self._wedstrijd_naam = ctk.CTkEntry(
            sb, placeholder_text="Naam van de wedstrijd...", height=36)
        self._wedstrijd_naam.grid(row=r, column=0, padx=16, pady=4, sticky="ew"); r += 1
        self._wedstrijd_naam.insert(0, "Voorjaarswedstrijd 2026")

        self._sb_section(sb, "GEGEVENSBRON", r); r += 1
        self.src_var = tk.StringVar(value="manual")
        ctk.CTkRadioButton(sb, text="Handmatig invoeren",
                           variable=self.src_var, value="manual",
                           command=self._on_src_change).grid(
            row=r, column=0, padx=18, pady=3, sticky="w"); r += 1
        ctk.CTkRadioButton(sb, text="CSV / Excel laden",
                           variable=self.src_var, value="file",
                           command=self._on_src_change).grid(
            row=r, column=0, padx=18, pady=3, sticky="w"); r += 1
        ctk.CTkButton(sb, text="Laad bestand",
                      command=self._pick_file, height=34).grid(
            row=r, column=0, padx=16, pady=(6, 2), sticky="ew"); r += 1
        self.lbl_file = ctk.CTkLabel(sb, text="",
                                      font=ctk.CTkFont(size=10),
                                      text_color="gray", wraplength=220)
        self.lbl_file.grid(row=r, column=0, padx=16, pady=(0, 4)); r += 1
        ctk.CTkButton(sb, text="Demo-data laden",
                      command=self._use_demo, height=34,
                      fg_color="transparent", border_width=1).grid(
            row=r, column=0, padx=16, pady=4, sticky="ew"); r += 1

        self._sb_section(sb, "OPTIES", r); r += 1
        self.schrap_var = tk.BooleanVar(value=True)
        ctk.CTkSwitch(sb, text="Schrap slechtste reeks",
                      variable=self.schrap_var,
                      font=ctk.CTkFont(size=12)).grid(
            row=r, column=0, padx=18, pady=8, sticky="w"); r += 1

        self._sb_section(sb, "ACTIES", r); r += 1
        self.btn_calc = ctk.CTkButton(
            sb, text="Bereken",
            command=self._start_calculation, height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=CLR_PRIMARY, hover_color=CLR_ACCENT)
        self.btn_calc.grid(row=r, column=0, padx=16, pady=6, sticky="ew"); r += 1
        self.btn_export = ctk.CTkButton(
            sb, text="Export naar Excel",
            command=self._export_excel, height=36, state="disabled",
            fg_color="transparent", border_width=1)
        self.btn_export.grid(row=r, column=0, padx=16, pady=4, sticky="ew"); r += 1

        self._sb_section(sb, "WEERGAVE", 99)
        self.theme_sw = ctk.CTkSwitch(sb, text="Donker thema",
                                      command=self._toggle_theme,
                                      font=ctk.CTkFont(size=12))
        self.theme_sw.select()
        self.theme_sw.grid(row=100, column=0, padx=18, pady=(6, 28), sticky="w")

    def _sb_section(self, parent, text: str, row: int):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color="gray").grid(
            row=row, column=0, padx=18, pady=(14, 2), sticky="w")

    def _build_main(self):
        main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=1)

        self.tabs = ctk.CTkTabview(main)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)

        self._t_invoer = self.tabs.add("Invoer")
        self._t_klass  = self.tabs.add("Klassement")
        self._t_detail = self.tabs.add("Detail per reeks")
        self._t_py     = self.tabs.add("PY-tabel")

        for t in (self._t_invoer, self._t_klass, self._t_detail, self._t_py):
            t.grid_rowconfigure(1, weight=1)
            t.grid_columnconfigure(0, weight=1)

        self._build_tab_invoer()
        self._build_tab_klassement()
        self._build_tab_detail()
        self._build_tab_py()

    def _tab_title(self, parent, text: str):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(size=17, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=8, pady=(10, 6))

    def _build_tab_invoer(self):
        self._t_invoer.grid_rowconfigure(0, weight=1)
        self._t_invoer.grid_columnconfigure(0, weight=1)
        self.invoer_tab = InvoerTab(self._t_invoer, self.py_table)
        self.invoer_tab.grid(row=0, column=0, sticky="nsew")

    def _build_tab_klassement(self):
        self._tab_title(self._t_klass, "Eindklassement")
        self._klass_subtitle = ctk.CTkLabel(
            self._t_klass, text="",
            font=ctk.CTkFont(size=12), text_color="gray")
        self._klass_subtitle.grid(row=0, column=0, sticky="e", padx=12)

        self.lbl_placeholder = ctk.CTkLabel(
            self._t_klass,
            text="Voer tijden in of laad een bestand,\ndan klik op Bereken.",
            font=ctk.CTkFont(size=14), text_color="gray", justify="center")
        self.lbl_placeholder.grid(row=1, column=0)

        self.tv_klass = DataTable(
            self._t_klass, columns=["Eindstand", "Naam", "Boottype", "Totaal punten"])
        self.tv_klass.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        self.tv_klass.grid_remove()

    def _build_tab_detail(self):
        self._tab_title(self._t_detail, "Detail per reeks")
        self._detail_scroll = ctk.CTkScrollableFrame(self._t_detail)
        self._detail_scroll.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        self._detail_scroll.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self._detail_scroll,
                     text="Nog geen resultaten beschikbaar.",
                     font=ctk.CTkFont(size=13), text_color="gray").grid(
            row=0, column=0, pady=40)

    def _refresh_detail_tab(self, detail):
        for w in self._detail_scroll.winfo_children():
            w.destroy()

        cols = ["Naam", "Boottype", "PY", "Min", "Sec",
                "Totaal sec", "Gecorr. tijd", "Rang", "Punten"]
        col_w = {
            "Naam": 120, "Boottype": 130, "PY": 70,
            "Min": 52, "Sec": 52, "Totaal sec": 90,
            "Gecorr. tijd": 110, "Rang": 55, "Punten": 65,
        }

        reeksen = sorted(detail["Reeks"].unique())
        for idx, reeks in enumerate(reeksen):
            hdr = ctk.CTkLabel(
                self._detail_scroll,
                text=f"  Reeks {reeks}",
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=CLR_PRIMARY, corner_radius=6,
                anchor="w", height=36)
            hdr.grid(row=idx * 2, column=0, sticky="ew",
                     padx=4, pady=(16 if idx > 0 else 4, 4))

            subset = detail[detail["Reeks"] == reeks].drop(columns=["Reeks"])
            tbl = DataTable(self._detail_scroll, columns=cols, row_height=28)
            tbl.grid(row=idx * 2 + 1, column=0, sticky="ew", padx=4, pady=(0, 2))
            tbl.set_columns(cols, col_w)

            medals = {1: "gold", 2: "silver", 3: "bronze"}
            for i, (_, row) in enumerate(subset.iterrows()):
                rang = int(row.get("Rang", 99))
                tag  = medals.get(rang, "odd" if i % 2 == 0 else "even")
                tbl.add_row([self._fmt(row.get(c)) for c in cols], tag=tag)

    def _build_tab_py(self):
        self._tab_title(self._t_py, "Boottypes en PY-waarden")
        cols = ["Boottype", "PY-waarde"]
        self.tv_py = DataTable(self._t_py, columns=cols)
        self.tv_py.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        self.tv_py.set_columns(cols, {"Boottype": 260, "PY-waarde": 120})
        for i, (bt, py) in enumerate(
                sorted(self.py_table.items(), key=lambda x: x[1])):
            self.tv_py.add_row([bt, f"{py:.0f}"],
                               tag="odd" if i % 2 == 0 else "even")

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, height=38, corner_radius=0)
        bar.grid(row=1, column=1, sticky="ew")
        bar.grid_propagate(False)
        self.lbl_status = ctk.CTkLabel(
            bar, text="Gereed.",
            font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status.pack(side="left", padx=16)
        self.progress = ctk.CTkProgressBar(bar, width=150, height=8)
        self.progress.pack(side="right", padx=16, pady=10)
        self.progress.set(0)

    def _on_src_change(self):
        if self.src_var.get() == "manual":
            self.tabs.set("Invoer")
        self._data_source = self.src_var.get()

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Selecteer invoerbestand",
            filetypes=[("CSV en Excel", "*.csv *.xlsx *.xls"),
                       ("Alle bestanden", "*.*")])
        if path:
            self.invoer_pad = path
            self.lbl_file.configure(text=Path(path).name)
            self.src_var.set("file")
            self._data_source = "file"
            self._status(f"Bestand: {Path(path).name}")

    def _use_demo(self):
        from wedstrijd_calculator import DEMO_DATA
        self.invoer_tab.load_demo(DEMO_DATA, self.py_table)
        self.src_var.set("manual")
        self._data_source = "manual"
        self._wedstrijd_naam.delete(0, "end")
        self._wedstrijd_naam.insert(0, "Voorjaarswedstrijd 2026")
        self.tabs.set("Invoer")
        self._status("Demo-data geladen in de invoertab.")

    def _start_calculation(self):
        self.btn_calc.configure(state="disabled", text="Bezig...")
        self.progress.set(0.1)
        self._status("Berekening gestart...")
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        try:
            self.after(0, lambda: self.progress.set(0.25))
            if self._data_source == "manual":
                df = self.invoer_tab.get_dataframe()
            else:
                df = load_race_data(self.invoer_pad)
            self.after(0, lambda: self.progress.set(0.45))
            df = calculate_elapsed_seconds(df)
            df = calculate_corrected_time_py(df, self.py_table)
            df = rank_each_race(df)
            df = calculate_points(df)
            if self.schrap_var.get():
                df = drop_worst_result(df)
            self.after(0, lambda: self.progress.set(0.75))
            detail, samenvatting = generate_summary_tables(
                df, use_drop=self.schrap_var.get())
            self.detail_df       = detail
            self.samenvatting_df = samenvatting
            self.after(0, lambda: self._display_results(detail, samenvatting))
        except Exception as exc:
            self.after(0, lambda e=exc: self._on_error(str(e)))
        finally:
            self.after(0, lambda: self.btn_calc.configure(
                state="normal", text="Bereken"))

    def _wedstrijd_naam_get(self) -> str:
        return self._wedstrijd_naam.get().strip() or "Wedstrijd"

    def _display_results(self, detail, samenvatting):
        self.progress.set(1.0)
        naam = self._wedstrijd_naam_get()
        self._klass_subtitle.configure(text=naam)

        reeks_cols = [c for c in samenvatting.columns
                      if c.startswith("R") and c[1:].isdigit()]
        klass_cols = (["Eindstand", "Naam", "Boottype"]
                      + reeks_cols
                      + ["Som alle", "Slechtste", "Totaal punten"])
        klass_w = {
            "Eindstand": 80, "Naam": 120, "Boottype": 130,
            "Som alle": 84, "Slechtste": 84, "Totaal punten": 110,
            **{r: 60 for r in reeks_cols},
        }
        self.tv_klass.set_columns(klass_cols, klass_w)
        self.tv_klass.clear()
        medals = {1: "gold", 2: "silver", 3: "bronze"}
        for i, (_, row) in enumerate(samenvatting.iterrows()):
            stand = int(row.get("Eindstand", i + 1))
            tag   = medals.get(stand, "odd" if i % 2 == 0 else "even")
            self.tv_klass.add_row(
                [self._fmt(row.get(c)) for c in klass_cols], tag=tag)

        self.lbl_placeholder.grid_remove()
        self.tv_klass.grid()
        self._refresh_detail_tab(detail)

        self.btn_export.configure(state="normal")
        n_d = len(samenvatting)
        n_r = len(reeks_cols)
        self._status(f"{naam}  -  {n_d} deelnemers, {n_r} reeksen verwerkt.")
        self.tabs.set("Klassement")

    def _export_excel(self):
        if self.detail_df is None:
            return
        naam = self._wedstrijd_naam_get()
        safe = "".join(c for c in naam if c.isalnum() or c in " _-").strip()
        init = f"{safe}.xlsx" if safe else "wedstrijd_resultaten.xlsx"
        path = filedialog.asksaveasfilename(
            title="Opslaan als Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel-werkmap", "*.xlsx")],
            initialfile=init)
        if path:
            try:
                export_to_excel(self.detail_df, self.samenvatting_df, path)
                self._status(f"Excel opgeslagen: {Path(path).name}")
                messagebox.showinfo("Opgeslagen",
                                    f"Resultaten opgeslagen als:\n{path}")
            except Exception as exc:
                self._on_error(str(exc))

    def _toggle_theme(self):
        ctk.set_appearance_mode(
            "Dark" if self.theme_sw.get() else "Light")

    def _status(self, msg: str):
        self.lbl_status.configure(text=msg)

    def _on_error(self, msg: str):
        self.progress.set(0)
        self._status(f"Fout: {msg}")
        messagebox.showerror("Fout", msg)

    @staticmethod
    def _fmt(v):
        if isinstance(v, float):
            return f"{v:.2f}"
        return v if v is not None else ""


# ==============================================================================
if __name__ == "__main__":
    app = WedstrijdApp()
    app.mainloop()
