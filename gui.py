import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import json
import os
from pathlib import Path
from create_document import create_supplementary_document,createMasterBOL
from sorterFunctions import orderPDF,convertBOLsToDic


# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS & PATHS
# ─────────────────────────────────────────────────────────────────────────────
APP_TITLE   = "BOL Document Sorter & Creator"
APP_VERSION = "1.0.0"

CONFIG_DIR    = Path(__file__).parent / "configs"
PROFILES_PATH = CONFIG_DIR / "profiles.json"

# Colour palette
CLR_BG       = "#F5F7FA"
CLR_PANEL    = "#FFFFFF"
CLR_BORDER   = "#DDE3EC"
CLR_ACCENT   = "#2563EB"          # blue
CLR_ACCENT_H = "#1D4ED8"          # hover blue
CLR_DANGER   = "#DC2626"
CLR_DANGER_H = "#B91C1C"
CLR_TEXT     = "#111827"
CLR_SUBTEXT  = "#6B7280"
CLR_LOG_BG   = "#F8FAFC"
CLR_LOG_FG   = "#1E293B"
CLR_SUCCESS  = "#16A34A"

FONT_TITLE   = ("Segoe UI", 13, "bold")
FONT_LABEL   = ("Segoe UI", 9)
FONT_LABEL_B = ("Segoe UI", 9, "bold")
FONT_ENTRY   = ("Segoe UI", 10)
FONT_LOG     = ("Consolas", 9)
FONT_BTN     = ("Segoe UI", 9, "bold")
FONT_SMALL   = ("Segoe UI", 8)

LOG_PLACEHOLDER = (
    "Output log will appear here after you press Run.\n\n"
    "  • Select an input PDF on the left\n"
    "  • Choose or create a profile\n"
    "  • Fill in BOL #, Loading #, and Date\n"
    "  • Tick Master BOL and/or Supp as needed\n"
    "  • Press  ▶ Run  to begin processing\n"
)


# ─────────────────────────────────────────────────────────────────────────────
#  PROFILE PERSISTENCE
# ─────────────────────────────────────────────────────────────────────────────
PROFILE_SCHEMA = [
    ("profile_name",          "Profile Name"),
    ("company_name",          "Company Name"),
    ("ship_name",             "Ship Name"),
    ("ship_address1",         "Ship Address 1"),
    ("ship_address2",         "Ship Address 2"),
    ("city",                  "City"),
    ("state",                 "State"),
    ("zip",                   "ZIP"),
    ("carrier_name",          "Carrier Name"),
    ("commodity_description", "Commodity Description"),
]


def _ensure_config():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not PROFILES_PATH.exists():
        PROFILES_PATH.write_text("[]", encoding="utf-8")


def load_profiles() -> list[dict]:
    _ensure_config()
    try:
        data = json.loads(PROFILES_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def save_profiles(profiles: list[dict]):
    _ensure_config()
    PROFILES_PATH.write_text(
        json.dumps(profiles, indent=2, ensure_ascii=False), encoding="utf-8"
    )


# ─────────────────────────────────────────────────────────────────────────────
#  REUSABLE STYLED WIDGETS
# ─────────────────────────────────────────────────────────────────────────────
def styled_button(parent, text, command, color=CLR_ACCENT, hover=CLR_ACCENT_H,
                  fg="white", width=None, pady=6):
    """Flat coloured button that changes shade on hover."""
    kw = dict(
        text=text, command=command,
        bg=color, fg=fg,
        font=FONT_BTN,
        relief="flat", bd=0, cursor="hand2",
        padx=14, pady=pady, activebackground=hover, activeforeground=fg,
    )
    if width:
        kw["width"] = width
    btn = tk.Button(parent, **kw)
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=color))
    return btn


def section_label(parent, text):
    """Bold section heading with a bottom divider."""
    f = tk.Frame(parent, bg=CLR_BG)
    tk.Label(f, text=text, font=FONT_LABEL_B, bg=CLR_BG, fg=CLR_TEXT).pack(anchor="w")
    tk.Frame(f, height=1, bg=CLR_BORDER).pack(fill="x", pady=(2, 6))
    return f


def card(parent, **kw):
    """White rounded-ish card frame."""
    defaults = dict(bg=CLR_PANEL, relief="flat", bd=0,
                    highlightthickness=1, highlightbackground=CLR_BORDER)
    defaults.update(kw)
    return tk.Frame(parent, **defaults)


# ─────────────────────────────────────────────────────────────────────────────
#  PROFILE DIALOG  (Add / Edit)
# ─────────────────────────────────────────────────────────────────────────────
class ProfileDialog(tk.Toplevel):
    def __init__(self, master, existing: dict | None = None):
        super().__init__(master)
        self.title("Edit Profile" if existing else "Add Profile")
        self.resizable(False, False)
        self.configure(bg=CLR_BG)
        self.grab_set()
        self.result: dict | None = None
        self._build(existing or {})
        self.wait_window()

    def _build(self, data: dict):
        pad = dict(padx=12, pady=4)
        wrap = tk.Frame(self, bg=CLR_BG, padx=20, pady=16)
        wrap.pack(fill="both", expand=True)

        tk.Label(wrap, text="Profile Details", font=FONT_TITLE,
                 bg=CLR_BG, fg=CLR_TEXT).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self._vars: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate(PROFILE_SCHEMA, start=1):
            tk.Label(wrap, text=label, font=FONT_LABEL, bg=CLR_BG,
                     fg=CLR_SUBTEXT, width=22, anchor="w").grid(
                row=i, column=0, sticky="w", **pad)
            var = tk.StringVar(value=data.get(key, ""))
            ent = tk.Entry(wrap, textvariable=var, font=FONT_ENTRY,
                           bg=CLR_PANEL, fg=CLR_TEXT, relief="flat",
                           bd=0, highlightthickness=1,
                           highlightbackground=CLR_BORDER,
                           highlightcolor=CLR_ACCENT, width=34)
            ent.grid(row=i, column=1, sticky="ew", **pad)
            self._vars[key] = var

        # Buttons
        bf = tk.Frame(wrap, bg=CLR_BG)
        bf.grid(row=len(PROFILE_SCHEMA) + 1, column=0, columnspan=2,
                pady=(14, 0), sticky="e")
        styled_button(bf, "Cancel", self.destroy,
                      color="#E5E7EB", hover="#D1D5DB",
                      fg=CLR_TEXT, width=10).pack(side="left", padx=(0, 6))
        styled_button(bf, "Save Profile", self._save, width=14).pack(side="left")

    def _save(self):
        values = {k: v.get().strip() for k, v in self._vars.items()}
        if not values["profile_name"]:
            messagebox.showwarning("Required Field",
                                   "Profile Name cannot be empty.", parent=self)
            return
        self.result = values
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 2 — PROFILES MANAGER
# ─────────────────────────────────────────────────────────────────────────────
class ProfilesTab(tk.Frame):
    def __init__(self, master, on_change_cb):
        super().__init__(master, bg=CLR_BG)
        self._on_change = on_change_cb   # called after any mutation
        self._profiles: list[dict] = []
        self._build()
        self.reload()

    # ── Layout ──
    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=CLR_BG, pady=12, padx=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Profile Manager", font=FONT_TITLE,
                 bg=CLR_BG, fg=CLR_TEXT).pack(side="left")

        # Main content: list left, detail right
        body = tk.Frame(self, bg=CLR_BG, padx=18, pady=4)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        # ── Left: list + action buttons ──
        left = tk.Frame(body, bg=CLR_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=1)

        tk.Label(left, text="Saved Profiles", font=FONT_LABEL_B,
                 bg=CLR_BG, fg=CLR_TEXT).grid(row=0, column=0, sticky="w", pady=(0, 4))

        list_card = card(left)
        list_card.grid(row=1, column=0, sticky="nsew")

        self._listbox = tk.Listbox(
            list_card, font=FONT_ENTRY,
            bg=CLR_PANEL, fg=CLR_TEXT,
            selectbackground=CLR_ACCENT, selectforeground="white",
            relief="flat", bd=0, activestyle="none",
            highlightthickness=0,
        )
        sb = ttk.Scrollbar(list_card, orient="vertical",
                           command=self._listbox.yview)
        self._listbox.config(yscrollcommand=sb.set)
        self._listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        sb.pack(side="right", fill="y")
        self._listbox.bind("<<ListboxSelect>>", self._on_select)

        btn_row = tk.Frame(left, bg=CLR_BG)
        btn_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        styled_button(btn_row, "+ Add", self._add, width=8).pack(side="left", padx=(0, 6))
        styled_button(btn_row, "Edit", self._edit, width=8,
                      color="#4B5563", hover="#374151").pack(side="left", padx=(0, 6))
        styled_button(btn_row, "Delete", self._delete, width=8,
                      color=CLR_DANGER, hover=CLR_DANGER_H).pack(side="left")

        # ── Right: detail view ──
        right = tk.Frame(body, bg=CLR_BG)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)

        tk.Label(right, text="Profile Details", font=FONT_LABEL_B,
                 bg=CLR_BG, fg=CLR_TEXT).grid(row=0, column=0, sticky="w", pady=(0, 4))

        detail_card = card(right)
        detail_card.grid(row=1, column=0, sticky="nsew")
        right.columnconfigure(0, weight=1)

        self._detail = tk.Text(
            detail_card, font=FONT_LOG,
            bg=CLR_LOG_BG, fg=CLR_LOG_FG,
            relief="flat", bd=0, state="disabled",
            padx=12, pady=10, wrap="word",
            highlightthickness=0,
        )
        self._detail.pack(fill="both", expand=True)

        # Storage path label
        path_row = tk.Frame(self, bg=CLR_BG, padx=18, pady=6)
        path_row.pack(fill="x")
        tk.Label(path_row, text=f"Storage: {PROFILES_PATH}",
                 font=FONT_SMALL, bg=CLR_BG, fg=CLR_SUBTEXT).pack(side="left")

    # ── Data helpers ──
    def reload(self):
        self._profiles = load_profiles()
        self._refresh_list()

    def _refresh_list(self):
        self._listbox.delete(0, "end")
        for p in self._profiles:
            self._listbox.insert("end", f"  {p.get('profile_name', '(unnamed)')}")
        self._show_detail(None)

    def _selected_index(self) -> int | None:
        sel = self._listbox.curselection()
        return sel[0] if sel else None

    def _on_select(self, _event=None):
        idx = self._selected_index()
        self._show_detail(self._profiles[idx] if idx is not None else None)

    def _show_detail(self, profile: dict | None):
        self._detail.config(state="normal")
        self._detail.delete("1.0", "end")
        if profile:
            for key, label in PROFILE_SCHEMA:
                val = profile.get(key, "") or "—"
                self._detail.insert("end", f"{label}:\n", "lbl")
                self._detail.insert("end", f"  {val}\n\n")
            self._detail.tag_config("lbl", font=FONT_LABEL_B, foreground=CLR_SUBTEXT)
        else:
            self._detail.insert("end", "Select a profile to view details.")
        self._detail.config(state="disabled")

    # ── CRUD ──
    def _add(self):
        dlg = ProfileDialog(self)
        if dlg.result:
            self._profiles.append(dlg.result)
            save_profiles(self._profiles)
            self._refresh_list()
            self._on_change()

    def _edit(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showinfo("Select Profile", "Please select a profile to edit.")
            return
        dlg = ProfileDialog(self, existing=self._profiles[idx])
        if dlg.result:
            self._profiles[idx] = dlg.result
            save_profiles(self._profiles)
            self._refresh_list()
            self._on_change()

    def _delete(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showinfo("Select Profile", "Please select a profile to delete.")
            return
        name = self._profiles[idx].get("profile_name", "this profile")
        if messagebox.askyesno("Confirm Delete",
                               f'Delete profile "{name}"?\nThis cannot be undone.'):
            self._profiles.pop(idx)
            save_profiles(self._profiles)
            self._refresh_list()
            self._on_change()

    def focus_add(self):
        """Called when redirected from Tab 1's '+ Add New' dropdown option."""
        self._add()


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 — MAIN BOL TOOL
# ─────────────────────────────────────────────────────────────────────────────
class MainTab(tk.Frame):
    def __init__(self, master, get_profiles_cb, switch_to_profiles_cb):
        super().__init__(master, bg=CLR_BG)
        self._get_profiles    = get_profiles_cb       # returns list[dict]
        self._switch_profiles = switch_to_profiles_cb # jumps to tab 2
        self._pdf_path        = tk.StringVar()
        self._profile_var     = tk.StringVar()
        self._bol_num         = tk.StringVar()
        self._pallet_count    = tk.StringVar()
        self._loading_num     = tk.StringVar()
        self._date_var        = tk.StringVar(value=_today())
        self._master_bol      = tk.BooleanVar(value=False)
        self._create_supp     = tk.BooleanVar(value=False)
        self._build()
        self.refresh_profiles()

    # ── Layout ──
    def _build(self):
        # ── Header bar ──
        hdr = tk.Frame(self, bg=CLR_ACCENT, padx=18, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=APP_TITLE, font=FONT_TITLE,
                 bg=CLR_ACCENT, fg="white").pack(side="left")
        tk.Label(hdr, text=f"v{APP_VERSION}", font=FONT_SMALL,
                 bg=CLR_ACCENT, fg="#BFDBFE").pack(side="right", padx=4)

        # ── Two-column body ──
        body = tk.Frame(self, bg=CLR_BG, padx=16, pady=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3, minsize=320)
        body.columnconfigure(1, weight=4, minsize=360)
        body.rowconfigure(0, weight=1)

        self._build_left(body)
        self._build_right(body)

    def _build_left(self, parent):
        left = tk.Frame(parent, bg=CLR_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # ── Input PDF ──
        section_label(left, "Input Document").pack(fill="x")
        pdf_card = card(left, padx=12, pady=10)
        pdf_card.pack(fill="x", pady=(0, 12))

        tk.Label(pdf_card, text="Source PDF (multi-BOL file)",
                 font=FONT_LABEL, bg=CLR_PANEL, fg=CLR_SUBTEXT).pack(anchor="w")

        pick_row = tk.Frame(pdf_card, bg=CLR_PANEL)
        pick_row.pack(fill="x", pady=(4, 0))

        self._pdf_entry = tk.Entry(
            pick_row, textvariable=self._pdf_path,
            font=FONT_SMALL, bg=CLR_BG, fg=CLR_TEXT,
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground=CLR_BORDER,
            highlightcolor=CLR_ACCENT, state="readonly", width=28,
        )
        self._pdf_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))

        styled_button(pick_row, "Browse…", self._browse_pdf,
                      color="#4B5563", hover="#374151",
                      width=9, pady=4).pack(side="left")

        # ── Profile ──
        section_label(left, "Profile").pack(fill="x")
        prof_card = card(left, padx=12, pady=10)
        prof_card.pack(fill="x", pady=(0, 12))

        tk.Label(prof_card, text="Select Profile",
                 font=FONT_LABEL, bg=CLR_PANEL, fg=CLR_SUBTEXT).pack(anchor="w")

        self._profile_combo = ttk.Combobox(
            prof_card, textvariable=self._profile_var,
            state="readonly", font=FONT_ENTRY, width=32,
        )
        self._profile_combo.pack(fill="x", pady=(4, 0))
        self._profile_combo.bind("<<ComboboxSelected>>", self._on_profile_selected)

        # ── BOL Fields ──
        section_label(left, "BOL Information").pack(fill="x")
        info_card = card(left, padx=12, pady=10)
        info_card.pack(fill="x", pady=(0, 12))

        fields = [
            ("BOL Number *",    self._bol_num),
            ("Loading Number",  self._loading_num),
            ("Pallet Count",    self._pallet_count),
            ("Date *",          self._date_var),
        ]
        for label, var in fields:
            tk.Label(info_card, text=label, font=FONT_LABEL,
                     bg=CLR_PANEL, fg=CLR_SUBTEXT).pack(anchor="w", pady=(4, 0))
            ent = tk.Entry(
                info_card, textvariable=var,
                font=FONT_ENTRY, bg=CLR_BG, fg=CLR_TEXT,
                relief="flat", bd=0,
                highlightthickness=1, highlightbackground=CLR_BORDER,
                highlightcolor=CLR_ACCENT,
            )
            ent.pack(fill="x", ipady=5)

        # ── Options ──
        section_label(left, "Output Options").pack(fill="x")
        opt_card = card(left, padx=12, pady=10)
        opt_card.pack(fill="x", pady=(0, 12))

        style = ttk.Style()
        style.configure("BOL.TCheckbutton", background=CLR_PANEL, font=FONT_LABEL)

        ttk.Checkbutton(opt_card, text="Create Master BOL",
                        variable=self._master_bol,
                        style="BOL.TCheckbutton").pack(anchor="w", pady=2)
        ttk.Checkbutton(opt_card, text="Create Supplement (Supp)",
                        variable=self._create_supp,
                        style="BOL.TCheckbutton").pack(anchor="w", pady=2)

        # ── Run Button ──
        run_row = tk.Frame(left, bg=CLR_BG)
        run_row.pack(fill="x", pady=(4, 0))
        self._run_btn = styled_button(
            run_row, "▶  Run", self._on_run,
            color=CLR_ACCENT, hover=CLR_ACCENT_H,
            pady=9,
        )
        self._run_btn.pack(fill="x")

    def _build_right(self, parent):
        right = tk.Frame(parent, bg=CLR_BG)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # Header row
        log_hdr = tk.Frame(right, bg=CLR_BG)
        log_hdr.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(log_hdr, text="Output Log", font=FONT_LABEL_B,
                 bg=CLR_BG, fg=CLR_TEXT).pack(side="left")
        styled_button(log_hdr, "Clear", self._clear_log,
                      color="#E5E7EB", hover="#D1D5DB",
                      fg=CLR_TEXT, width=7, pady=2).pack(side="right")

        # Log text widget
        log_card = card(right)
        log_card.grid(row=1, column=0, sticky="nsew")
        log_card.rowconfigure(0, weight=1)
        log_card.columnconfigure(0, weight=1)

        self._log = tk.Text(
            log_card, font=FONT_LOG,
            bg=CLR_LOG_BG, fg=CLR_LOG_FG,
            relief="flat", bd=0,
            padx=12, pady=10,
            state="disabled", wrap="word",
            highlightthickness=0,
            insertbackground=CLR_ACCENT,
        )
        log_sb = ttk.Scrollbar(log_card, orient="vertical",
                               command=self._log.yview)
        self._log.config(yscrollcommand=log_sb.set)
        self._log.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=2)
        log_sb.grid(row=0, column=1, sticky="ns")

        # Colour tags for log
        self._log.tag_config("info",    foreground=CLR_LOG_FG)
        self._log.tag_config("success", foreground=CLR_SUCCESS)
        self._log.tag_config("error",   foreground=CLR_DANGER)
        self._log.tag_config("hint",    foreground=CLR_SUBTEXT, font=FONT_SMALL)
        self._log.tag_config("sep",     foreground=CLR_BORDER)

        self._write_placeholder()

    # ── Log helpers ──
    def _write_placeholder(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.insert("end", LOG_PLACEHOLDER, "hint")
        self._log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def log(self, msg: str, tag: str = "info"):
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")
        self.update_idletasks()

    # ── Profile helpers ──
    def refresh_profiles(self):
        profiles = self._get_profiles()
        names = [p.get("profile_name", "(unnamed)") for p in profiles]
        names.append("+ Add New")
        self._profile_combo["values"] = names
        # Preserve selection if still valid
        cur = self._profile_var.get()
        if cur not in names:
            self._profile_var.set("")

    def _on_profile_selected(self, _event=None):
        if self._profile_var.get() == "+ Add New":
            self._profile_var.set("")
            self._switch_profiles()

    # ── Browse PDF ──
    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Select source PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self._pdf_path.set(path)

    # ── Validation ──
    def _validate(self) -> list[str]:
        errors = []
        if not self._pdf_path.get():
            errors.append("No input PDF selected.")
        if not self._profile_var.get() or self._profile_var.get() == "+ Add New":
            errors.append("No profile selected.")
        if not self._bol_num.get().strip():
            errors.append("BOL Number is required.")
        if not self._date_var.get().strip():
            errors.append("Date is required.")
        return errors

    # ── Run ──
    def _on_run(self):
        # 1. Validate
        errors = self._validate()
        if errors:
            messagebox.showerror(
                "Missing Required Fields",
                "Please fix the following before running:\n\n"
                + "\n".join(f"  • {e}" for e in errors),
            )
            return

        # 2. Prompt for output file
        output_path = filedialog.asksaveasfilename(
            title="Save output PDF as…",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"BOL_{self._bol_num.get().strip()}.pdf",
        )
        if not output_path:
            return  # user cancelled

        # 3. Gather values
        profiles      = self._get_profiles()
        profile_name  = self._profile_var.get()
        profile       = next(
            (p for p in profiles if p.get("profile_name") == profile_name), {}
        )
        input_pdf     = self._pdf_path.get()
        bol_number    = self._bol_num.get().strip()
        loading_number= self._loading_num.get().strip()
        date          = self._date_var.get().strip()
        pallet_count   = self._pallet_count.get().strip()
        create_master = self._master_bol.get()
        create_supp   = self._create_supp.get()

        # 4. Clear log & write header
        self._clear_log()
        self.log(f"{'─' * 48}", "sep")
        self.log(f"  BOL Processing Started", "info")
        self.log(f"{'─' * 48}", "sep")
        self.log(f"  Input    : {input_pdf}")
        self.log(f"  Profile  : {profile_name}")
        self.log(f"  BOL #    : {bol_number}")
        self.log(f"  Loading #: {loading_number}")
        self.log(f"  Date     : {date}")
        self.log(f"  Master   : {'Yes' if create_master else 'No'}")
        self.log(f"  Supp     : {'Yes' if create_supp else 'No'}")
        self.log(f"  Output   : {output_path}")
        self.log(f"{'─' * 48}", "sep")

        # ──────────────────────────────────────────────────────────────────
        #  ▼▼▼  BACKEND HOOK  ▼▼▼
        #  Replace this block with your BOL creation / sorting logic.
        #  All values are available as plain Python variables:
        #
        #    input_pdf      : str   — absolute path to source PDF
        #    output_path    : str   — absolute path for output PDF
        #    profile        : dict  — full profile dict (keys match PROFILE_SCHEMA)
        #    bol_number     : str
        #    loading_number : str
        #    date           : str
        #    create_master  : bool
        #    create_supp    : bool
        #
        #  Use self.log(message) to write lines to the output log.
        #  Use self.log(message, "success") / ("error") for coloured lines.
        # ──────────────────────────────────────────────────────────────────
        run_bol_processing(
            input_pdf=input_pdf,
            output_path=output_path,
            profile=profile,
            bol_number=bol_number,
            loading_number=loading_number,
            date=date,
            create_master=create_master,
            create_supp=create_supp,
            log=self.log,
            pallet_count=pallet_count,
        )
        # ──────────────────────────────────────────────────────────────────
        #  ▲▲▲  END BACKEND HOOK  ▲▲▲
        # ──────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
#  ██████████████████████████████████████████████████████████████████████████
#  BACKEND HOOK FUNCTION — replace the body of this function with your logic
#  ██████████████████████████████████████████████████████████████████████████
# ─────────────────────────────────────────────────────────────────────────────
def run_bol_processing(
    input_pdf: str,
    output_path: str,
    profile: dict,
    bol_number: str,
    loading_number: str,
    pallet_count: str,
    date: str,
    create_master: bool,
    create_supp: bool,
    log,          # callable: log(message, tag="info"|"success"|"error")
):

    sortedPDF = orderPDF(input_pdf)

    all_bols = ""        
    for index, page in enumerate(sortedPDF):
            all_bols += (page[2] + " ")

    bol_dic=convertBOLsToDic(sortedPDF)

    masterBOL=createMasterBOL(bol_number,loading_number,date,sortedPDF,all_bols,[profile.get("company_name",""),profile.get("ship_name",""),profile.get("ship_address1",""),profile.get("ship_address2",""),profile.get("city",""),profile.get("state",""),profile.get("zip",""),profile.get("carrier_name",""),profile.get("commodity_description","")],pallet_count)
    supp_file=create_supplementary_document(date,bol_number,loading_number,pallet_count,bol_dic)

    exportPDF(sortedPDF,PdfReader(input_pdf), output_path)
    #dcInfo=[row[4] for row in sortedPDF]
    #allWeights=[row[4][2] for row in sortedPDF]
    #allQty=[row[4][1] for row in sortedPDF]
    #totalWeight=sum([int(weight) for weight in allWeights])
    #totalQty=sum([int(qty) for qty in allQty])

    #if create_supp:
        #suppDoc=create_supplementary_document(date,bol_number,loading_number,str(totalQty),str(totalWeight),dcInfo)
    
    #if create_master:
       # masterBOL=createMasterBOL(bol_number,loading_number,date,profile.get("company_name",""),profile.get("ship_name",""),profile.get("ship_address1",""),profile.get("ship_address2",""),profile.get("city",""),profile.get("state",""),profile.get("zip",""),profile.get("carrier_name",""),profile.get("commodity_description",""))



    log("  [STUB] run_bol_processing() called — add your logic here.", "hint")
    log(f"  input_pdf      = {input_pdf}")
    log(f"  output_path    = {output_path}")
    log(f"  profile        = {profile}")
    log(f"  bol_number     = {bol_number}")
    log(f"  loading_number = {loading_number}")
    log(f"  date           = {date}")
    log(f"  create_master  = {create_master}")
    log(f"  create_supp    = {create_supp}")
    log("  [STUB] Done (no output written).", "hint")


# ─────────────────────────────────────────────────────────────────────────────
#  APPLICATION ROOT
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1020x660")
        self.minsize(820, 560)
        self.configure(bg=CLR_BG)
        self._profiles: list[dict] = load_profiles()
        self._build()

    def _build(self):
        # ttk notebook style
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",        background=CLR_BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=CLR_BORDER, foreground=CLR_SUBTEXT,
                        font=FONT_LABEL_B, padding=[14, 6])
        style.map("TNotebook.Tab",
                  background=[("selected", CLR_PANEL)],
                  foreground=[("selected", CLR_ACCENT)])
        style.configure("TCombobox", fieldbackground=CLR_BG, background=CLR_PANEL)

        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill="both", expand=True)

        # Tab 1
        self._main_tab = MainTab(
            self._notebook,
            get_profiles_cb=lambda: self._profiles,
            switch_to_profiles_cb=self._go_to_profiles,
        )
        self._notebook.add(self._main_tab, text="  BOL Tool  ")

        # Tab 2
        self._profiles_tab = ProfilesTab(
            self._notebook,
            on_change_cb=self._on_profiles_changed,
        )
        self._notebook.add(self._profiles_tab, text="  Profiles  ")

    def _go_to_profiles(self):
        self._notebook.select(1)
        self._profiles_tab.focus_add()

    def _on_profiles_changed(self):
        self._profiles = load_profiles()
        self._main_tab.refresh_profiles()


# ─────────────────────────────────────────────────────────────────────────────
#  UTILITIES
# ─────────────────────────────────────────────────────────────────────────────
def _today() -> str:
    from datetime import date
    return date.today().strftime("%m/%d/%Y")

def sort_function(file_path):
    """Function that executes when Sort button is clicked"""
    
    if not file_path:
        messagebox.showwarning("No File", "Please select a file first!")
        return
    
    return orderPDF(file_path)

    exportPDF(sortedPDF, PdfReader(file_path), export_path)
    dcInfo=[row[4] for row in sortedPDF]
    allWeights=[row[4][2] for row in sortedPDF]
    allQty=[row[4][1] for row in sortedPDF]
    totalWeight=sum([int(weight) for weight in allWeights])
    totalQty=sum([int(qty) for qty in allQty])
    print(dcInfo)
    #create_supplementary_document("03/12/26","1234567890","1",str(totalQty),str(totalWeight),dcInfo)
    print(f"Processing file: {file_path}")

def exportPDF(sortedPDF, reader, exportPath,supplementaryDoc=None,masterBOL=None):
    bol_list = ""
    bolDic = convertBOLsToDic(sortedPDF)
    try:
        output_pdf = PdfWriter()

        for index, page in enumerate(sortedPDF):
            output_pdf.add_page(reader.get_page(page[0]))
            bol_list += (page[2] + " ")

        if masterBOL:
            masterBOL=createMasterBOL()
            output_pdf.append(masterBOL)
            os.remove(masterBOL)

        with open(Path(exportPath), "wb") as f:
            print('Write to:' + exportPath)
            output_pdf.write(f)
    except Exception as error:
        print(error)




# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()