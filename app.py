import os
import sys
import json
import warnings
import traceback
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

warnings.filterwarnings("ignore")

import brand
import invoice
import process
import summary
from engine import MONTHS

SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "Sully")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")
CURRENCIES = ["ZAR", "USD", "EUR", "GBP"]

CHIPS = {
    "PROBLEM": ("#F8D7DA", "#842029"),
    "CHECK": ("#FBE38A", "#5C4700"),
    "LEFT OUT": ("#F7C99B", "#6B3300"),
    "ASSUMPTION": ("#CFD6F2", "#1F2A66"),
    "FIXED": ("#BFE3C7", "#0F4A25"),
    "NOTE": ("#EDEDED", "#424242"),
}
CHIP_ORDER = ["PROBLEM", "CHECK", "LEFT OUT", "ASSUMPTION", "FIXED", "NOTE"]
BANNERS = {
    "info": ("#EDEDED", "#424242"),
    "working": ("#FFF1BF", "#6B5300"),
    "ok": ("#2D6A35", "#FFFFFF"),
    "error": ("#F8D7DA", "#842029"),
}


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass


def to_number(text):
    cleaned = str(text).strip().replace(" ", "").replace(",", "")
    if not cleaned:
        return 0.0
    return float(cleaned)


def money(value):
    return f"{value:,.2f}"


def categorize(text, is_error=False):
    stripped = text.strip()
    if is_error:
        return "PROBLEM", stripped
    for prefix, label in (("CORRECTED:", "FIXED"), ("EXCLUDED", "LEFT OUT"), ("ASSUMED:", "ASSUMPTION")):
        if stripped.startswith(prefix):
            return label, stripped[len(prefix):].lstrip(": ").strip() if prefix.endswith(":") else stripped
    if "NOT USED" in stripped[:80]:
        return "LEFT OUT", stripped
    if stripped.startswith("NEW SUBSCRIPTION") or stripped.startswith("Merged the branch") or stripped.startswith("Subtropico 'Head Office'"):
        return "NOTE", stripped
    return "CHECK", stripped


class InvoiceDialog(tk.Toplevel):
    def __init__(self, parent, inv, subscriptions):
        super().__init__(parent)
        self.title("Invoice details")
        self.configure(bg=brand.TK_WHITE)
        self.resizable(False, False)
        self.inv = inv
        self.result = False
        self.vars = {
            "subscription": tk.StringVar(value=inv.subscription),
            "supplier": tk.StringVar(value=inv.supplier),
            "number": tk.StringVar(value=inv.number),
            "currency": tk.StringVar(value=inv.currency),
            "excl": tk.StringVar(value=f"{inv.excl:,.2f}"),
            "vat": tk.StringVar(value=f"{inv.vat:,.2f}"),
            "incl": tk.StringVar(value=f"{inv.incl:,.2f}"),
        }
        frame = ttk.Frame(self, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=inv.file, font=(brand.FONT, 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        rows = [("Subscription", "subscription", subscriptions), ("Supplier", "supplier", None), ("Invoice number", "number", None),
                ("Currency", "currency", CURRENCIES), ("Excl VAT", "excl", None), ("VAT", "vat", None), ("Incl VAT", "incl", None)]
        for i, (label, key, values) in enumerate(rows, 1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", pady=3)
            if values:
                ttk.Combobox(frame, textvariable=self.vars[key], values=values, width=28).grid(row=i, column=1, sticky="w", padx=8)
            else:
                ttk.Entry(frame, textvariable=self.vars[key], width=31).grid(row=i, column=1, sticky="w", padx=8)
        self.force_var = tk.IntVar(value=1 if inv.force_include else 0)
        ttk.Checkbutton(frame, text="Use anyway, even if the invoice is dated outside the selected month", variable=self.force_var).grid(
            row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))
        if inv.notes:
            ttk.Label(frame, text="\n".join(inv.notes), wraplength=380, foreground=brand.TK_GREY_MID).grid(row=9, column=0, columnspan=2, sticky="w", pady=(8, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=10, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="OK", style="Accent.TButton", command=self.accept).pack(side="right")
        self.transient(parent)
        self.grab_set()

    def accept(self):
        try:
            excl, vat, incl = (to_number(self.vars[k].get()) for k in ("excl", "vat", "incl"))
        except ValueError:
            messagebox.showwarning("Check the amounts", "The amounts must be numbers, for example 3,193.03", parent=self)
            return
        if not incl:
            incl = round(excl + vat, 2)
        self.inv.force_include = bool(self.force_var.get())
        if (excl, vat, incl) != (self.inv.excl, self.inv.vat, self.inv.incl) or self.vars["subscription"].get().strip() != self.inv.subscription:
            self.inv.components = []
            self.inv.usd = None
        self.inv.subscription = self.vars["subscription"].get().strip()
        self.inv.supplier = self.vars["supplier"].get().strip()
        self.inv.number = self.vars["number"].get().strip()
        self.inv.currency = self.vars["currency"].get().strip().upper() or "ZAR"
        self.inv.excl, self.inv.vat, self.inv.incl = excl, vat, incl
        self.result = True
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sully")
        self.geometry(f"{min(1320, self.winfo_screenwidth() - 40)}x{min(820, self.winfo_screenheight() - 90)}+20+10")
        self.configure(bg=brand.TK_WHITE)
        self.settings = load_settings()
        self.list_var = tk.StringVar(value=self.settings.get("list", ""))
        self.book_var = tk.StringVar(value=self.settings.get("book", ""))
        today = datetime.date.today()
        self.month_var = tk.StringVar(value=MONTHS[today.month - 2 if today.month > 1 else 11])
        self.year_var = tk.StringVar(value=str(today.year if today.month > 1 else today.year - 1))
        self.last_output = None
        self.invoices = []
        self.group_cache = (None, [g[0] for g in summary.BASE_GROUPS])
        self.images = []
        self.set_style()
        self.set_icon()
        self.build()

    def set_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=(brand.FONT, 10), background=brand.TK_WHITE, foreground=brand.TK_GREY_DARK)
        style.configure("TFrame", background=brand.TK_WHITE)
        style.configure("TLabel", background=brand.TK_WHITE)
        style.configure("TLabelframe", background=brand.TK_WHITE, bordercolor=brand.TK_GREY_LIGHT)
        style.configure("TLabelframe.Label", background=brand.TK_WHITE, foreground=brand.TK_GREEN_DARK, font=(brand.FONT, 10, "bold"))
        style.configure("TButton", padding=(10, 4), background=brand.TK_GREY_LIGHT, foreground=brand.TK_GREY_DARK)
        style.map("TButton", background=[("active", brand.TK_GREY_MID), ("disabled", "#F0F0F0")], foreground=[("disabled", brand.TK_GREY_MID)])
        style.configure("Accent.TButton", background=brand.TK_GREEN, foreground=brand.TK_WHITE, font=(brand.FONT, 11, "bold"), padding=(16, 6))
        style.map("Accent.TButton", background=[("active", brand.TK_GREEN_DARK), ("disabled", brand.TK_GREY_LIGHT)], foreground=[("disabled", brand.TK_GREY_MID)])
        style.configure("Treeview", rowheight=22, fieldbackground=brand.TK_WHITE, borderwidth=0)
        style.configure("Treeview.Heading", background=brand.TK_GREEN, foreground=brand.TK_WHITE, font=(brand.FONT, 9, "bold"))
        style.map("Treeview.Heading", background=[("active", brand.TK_GREEN_DARK)])
        style.map("Treeview", background=[("selected", brand.TK_GREEN)], foreground=[("selected", brand.TK_WHITE)])
        style.configure("TNotebook", background=brand.TK_WHITE, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(14, 6), font=(brand.FONT, 10, "bold"), background=brand.TK_GREY_LIGHT, foreground=brand.TK_GREY_DARK)
        style.map("TNotebook.Tab", background=[("selected", brand.TK_GREEN)], foreground=[("selected", brand.TK_WHITE)])

    def set_icon(self):
        try:
            icon = tk.PhotoImage(file=brand.resource_path(os.path.join("assets", "app_icon.png")))
            self.images.append(icon)
            self.iconphoto(True, icon)
        except Exception:
            pass

    def table(self, parent, columns, widths, anchors, height=8):
        holder = ttk.Frame(parent)
        tree = ttk.Treeview(holder, columns=[c[0] for c in columns], show="headings", height=height)
        for (key, heading), width, anchor in zip(columns, widths, anchors):
            tree.heading(key, text=heading)
            tree.column(key, width=width, anchor=anchor, stretch=False)
        bar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        xbar = ttk.Scrollbar(holder, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=bar.set, xscrollcommand=xbar.set)
        bar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)
        tree.tag_configure("odd", background="#F6F6F6")
        tree.tag_configure("bu", background=brand.TK_GREY_LIGHT, font=(brand.FONT, 10, "bold"))
        tree.tag_configure("total", background="#FBEDB8", font=(brand.FONT, 10, "bold"))
        tree.tag_configure("zero", foreground="#A0A0A0")
        return holder, tree

    def build(self):
        header = tk.Frame(self, bg=brand.TK_GREEN_DARK)
        header.pack(fill="x")
        try:
            logo = tk.PhotoImage(file=brand.resource_path(os.path.join("assets", "logo_white.png"))).subsample(4, 4)
            self.images.append(logo)
            tk.Label(header, image=logo, bg=brand.TK_GREEN_DARK).pack(side="left", padx=(18, 14), pady=10)
        except Exception:
            tk.Label(header, text="SUBTROPICO", bg=brand.TK_GREEN_DARK, fg=brand.TK_WHITE, font=(brand.FONT, 18, "bold")).pack(side="left", padx=18, pady=10)
        titles = tk.Frame(header, bg=brand.TK_GREEN_DARK)
        titles.pack(side="left")
        tk.Label(titles, text="Sully", bg=brand.TK_GREEN_DARK, fg=brand.TK_WHITE, font=(brand.FONT, 16, "bold")).pack(anchor="w")
        tk.Label(titles, text="Subscription user counts, invoice split and updated workbook", bg=brand.TK_GREEN_DARK, fg=brand.TK_GREY_LIGHT, font=(brand.FONT, 10)).pack(anchor="w")
        tk.Frame(self, bg=brand.TK_YELLOW, height=4).pack(fill="x")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=12, pady=8)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew")

        files = ttk.LabelFrame(left, text="1. Files")
        files.pack(fill="x")
        files.columnconfigure(0, weight=1)
        ttk.Label(files, text="This month's user list:").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 0))
        ttk.Entry(files, textvariable=self.list_var, width=46).grid(row=1, column=0, sticky="ew", padx=(8, 4))
        ttk.Button(files, text="Browse...", command=self.pick_list).grid(row=1, column=1, padx=(0, 8))
        ttk.Label(files, text="Latest '2. Beperk Verhalings' workbook:").grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 0))
        ttk.Entry(files, textvariable=self.book_var, width=46).grid(row=3, column=0, sticky="ew", padx=(8, 4))
        ttk.Button(files, text="Browse...", command=self.pick_book).grid(row=3, column=1, padx=(0, 8))
        pick = ttk.Frame(files)
        pick.grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 8))
        ttk.Label(pick, text="Month:").pack(side="left")
        ttk.Combobox(pick, textvariable=self.month_var, values=MONTHS, state="readonly", width=13).pack(side="left", padx=(6, 4))
        ttk.Combobox(pick, textvariable=self.year_var, values=[str(y) for y in range(2025, 2036)], state="readonly", width=6).pack(side="left")
        self.month_var.trace_add("write", lambda *args: self.refresh_tree() if hasattr(self, "tree") else None)
        self.year_var.trace_add("write", lambda *args: self.refresh_tree() if hasattr(self, "tree") else None)

        inv = ttk.LabelFrame(left, text="2. Invoices (optional)")
        inv.pack(fill="both", expand=True, pady=(8, 0))
        columns = ("sub", "number", "date", "cur", "incl", "used")
        widths = (116, 112, 70, 36, 76, 44)
        headings = ("Subscription", "Invoice no.", "Date", "Cur.", "Incl VAT", "Used?")
        holder = ttk.Frame(inv)
        holder.pack(fill="both", expand=True, padx=8, pady=(6, 2))
        self.tree = ttk.Treeview(holder, columns=columns, show="headings", height=9, selectmode="extended")
        for c, w, h in zip(columns, widths, headings):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="e" if c == "incl" else "w", stretch=False)
        bar = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        self.tree.tag_configure("unused", foreground="#A0A0A0")
        self.tree.bind("<Double-1>", lambda e: self.edit_invoice())
        row = ttk.Frame(inv)
        row.pack(fill="x", padx=8, pady=(2, 2))
        ttk.Button(row, text="Add PDFs...", command=self.add_invoices).pack(side="left")
        ttk.Button(row, text="Edit", command=self.edit_invoice).pack(side="left", padx=4)
        ttk.Button(row, text="Remove", command=self.remove_invoices).pack(side="left")
        ttk.Label(inv, text="Double-click an invoice to check or change it.", foreground=brand.TK_GREY_MID, font=(brand.FONT, 8)).pack(anchor="w", padx=8, pady=(0, 6))

        action = ttk.Frame(left)
        action.pack(fill="x", pady=(10, 0))
        self.create_btn = ttk.Button(action, text="Create workbook", style="Accent.TButton", command=self.run_create)
        self.create_btn.pack(side="left")
        self.open_btn = ttk.Button(action, text="Open folder", command=self.open_folder, state="disabled")
        self.open_btn.pack(side="left", padx=8)

        self.banner_frame = tk.Frame(right, bg=BANNERS["info"][0])
        self.banner_frame.pack(fill="x", pady=(0, 8))
        self.banner_title = tk.Label(self.banner_frame, text="", anchor="w", justify="left", font=(brand.FONT, 11, "bold"))
        self.banner_title.pack(fill="x", padx=12, pady=(8, 0))
        self.banner_detail = tk.Label(self.banner_frame, text="", anchor="w", justify="left", font=(brand.FONT, 9), wraplength=640)
        self.banner_detail.pack(fill="x", padx=12, pady=(0, 8))
        self.banner_frame.bind("<Configure>", lambda e: self.banner_detail.configure(wraplength=max(200, e.width - 28)))

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)
        self.build_overview()
        self.build_checks()
        self.build_branch()
        self.build_split()
        self.build_full()
        self.banner("info", "Ready", "Choose the user list and the latest '2. Beperk Verhalings' workbook, pick the month, add the invoice PDFs if you have them, "
                    "then click 'Create workbook'. Your files are never changed: the updated workbook is saved as a new file next to your workbook.")

    def build_overview(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Overview")
        self.tab_overview = tab
        cards = ttk.Frame(tab)
        cards.pack(fill="x", padx=8, pady=(10, 6))
        self.cards = {}
        for key, caption in (("users", "Users read"), ("seats", "Microsoft 365 seats"), ("branches", "Branches"), ("invoices", "Invoices used")):
            box = tk.Frame(cards, bg=brand.TK_GREY_LIGHT)
            box.pack(side="left", fill="x", expand=True, padx=(0, 8))
            number = tk.Label(box, text="-", bg=brand.TK_GREY_LIGHT, fg=brand.TK_GREEN_DARK, font=(brand.FONT, 22, "bold"))
            number.pack(pady=(8, 0))
            tk.Label(box, text=caption, bg=brand.TK_GREY_LIGHT, fg=brand.TK_GREY_DARK, font=(brand.FONT, 9)).pack(pady=(0, 8))
            self.cards[key] = number
        tables = ttk.Frame(tab)
        tables.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        left = ttk.LabelFrame(tables, text="Microsoft 365 seats per entity")
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        holder, self.entity_tree = self.table(left, [("name", "Entity"), ("now", "This month"), ("before", "Last month"), ("change", "Change")],
                                              (212, 68, 68, 50), ("w", "e", "e", "e"), height=11)
        holder.pack(fill="both", expand=True, padx=6, pady=6)
        right = ttk.LabelFrame(tables, text="Adobe, Read AI and ChatGPT users")
        right.pack(side="left", fill="both", expand=True)
        holder, self.other_tree = self.table(right, [("name", "Subscription"), ("total", "Total"), ("vs", "VS"), ("sma", "SMA"), ("bpk", "BPK")],
                                             (128, 48, 36, 36, 38), ("w", "e", "e", "e", "e"), height=11)
        holder.pack(fill="both", expand=True, padx=6, pady=6)

    def build_checks(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Please check")
        self.tab_checks = tab
        self.checks_text = tk.Text(tab, wrap="word", font=(brand.FONT, 10), bg=brand.TK_WHITE, fg=brand.TK_GREY_DARK, relief="flat",
                                   padx=14, pady=10, cursor="arrow", spacing1=2, spacing3=6)
        bar = ttk.Scrollbar(tab, orient="vertical", command=self.checks_text.yview)
        self.checks_text.configure(yscrollcommand=bar.set, state="disabled")
        self.checks_text.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        for label, (background, foreground) in CHIPS.items():
            self.checks_text.tag_configure(f"chip_{label}", background=background, foreground=foreground, font=(brand.FONT, 8, "bold"))
            self.checks_text.tag_configure(f"body_{label}", lmargin1=0, lmargin2=0)
        self.checks_text.tag_configure("empty", foreground=brand.TK_GREY_MID, font=(brand.FONT, 10, "italic"))
        self.checks_text.tag_configure("heading", font=(brand.FONT, 10, "bold"), foreground=brand.TK_GREEN_DARK, spacing1=8)

    def build_branch(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="By branch")
        self.tab_branch = tab
        self.branch_holder = None
        self.branch_tree = None
        self.branch_hint = ttk.Label(tab, text="The number of users per business unit and branch appears here after you create the workbook.", foreground=brand.TK_GREY_MID)
        self.branch_hint.pack(padx=14, pady=14, anchor="w")

    def build_split(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Invoice split")
        self.tab_split = tab
        self.split_hint = ttk.Label(tab, text="Add invoice PDFs and create the workbook to see how each invoice is divided over the branches.", foreground=brand.TK_GREY_MID)
        self.split_hint.pack(padx=14, pady=14, anchor="w")
        self.split_area = ttk.Frame(tab)
        top = ttk.LabelFrame(self.split_area, text="What was divided (per subscription and licence type)")
        top.pack(fill="x", padx=8, pady=(8, 4))
        holder, self.block_tree = self.table(top, [("title", "Block"), ("users", "Users"), ("excl", "Excl VAT"), ("vat", "VAT"), ("incl", "Incl VAT")],
                                             (330, 52, 92, 84, 92), ("w", "e", "e", "e", "e"), height=6)
        holder.pack(fill="x", padx=6, pady=6)
        self.split_bottom = ttk.LabelFrame(self.split_area, text="Incl VAT per business unit and branch")
        self.split_bottom.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        self.split_holder = None
        self.split_tree = None

    def build_full(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Full report")
        self.tab_full = tab
        self.text = scrolledtext.ScrolledText(tab, wrap="none", font=("Consolas", 10), bg="#FAFAFA", fg=brand.TK_GREY_DARK, relief="flat")
        self.text.pack(fill="both", expand=True)
        self.text.configure(state="disabled")

    def banner(self, kind, title, detail=""):
        background, foreground = BANNERS[kind]
        self.banner_frame.configure(bg=background)
        self.banner_title.configure(text=title, bg=background, fg=foreground)
        self.banner_detail.configure(text=detail, bg=background, fg=foreground)
        self.update_idletasks()

    def set_full(self, text):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", text)
        self.text.configure(state="disabled")

    def say(self, message, kind="info", detail=""):
        lines = message.strip().splitlines()
        self.banner(kind, lines[0] if lines else "", detail or (lines[1] if len(lines) > 1 and kind != "error" else ""))
        if kind == "error":
            self.set_full(message)
            self.notebook.select(self.tab_full)

    def render_checks(self, errors, items):
        entries = [categorize(e, True) for e in errors] + [categorize(i) for i in items]
        order = {label: n for n, label in enumerate(CHIP_ORDER)}
        entries.sort(key=lambda entry: order[entry[0]])
        self.checks_text.configure(state="normal")
        self.checks_text.delete("1.0", "end")
        if not entries:
            self.checks_text.insert("end", "Nothing to check. Everything lined up.", "empty")
        for label, message in entries:
            self.checks_text.insert("end", f" {label} ", f"chip_{label}")
            self.checks_text.insert("end", "  " + message + "\n", f"body_{label}")
        self.checks_text.configure(state="disabled")
        attention = sum(1 for label, _ in entries if label in ("PROBLEM", "CHECK", "LEFT OUT", "ASSUMPTION"))
        self.notebook.tab(self.tab_checks, text=f"Please check ({attention})" if attention else "Please check")

    def fill(self, tree, rows):
        tree.delete(*tree.get_children())
        for n, (values, tags) in enumerate(rows):
            tree.insert("", "end", values=values, tags=tags + (("odd",) if n % 2 and not tags else ()))

    def show_branch_table(self, heads, rows):
        if self.branch_holder is not None:
            self.branch_holder.destroy()
        self.branch_hint.pack_forget()
        short = {"Fresh Service Agent": "Fresh Svc"}
        columns = [("name", "Business unit / branch")] + [(f"c{i}", short.get(h, h)) for i, h in enumerate(heads)]
        widths = [200] + [92] * len(heads)
        self.branch_holder, self.branch_tree = self.table(self.tab_branch, columns, widths, ["w"] + ["e"] * len(heads), height=20)
        self.branch_holder.pack(fill="both", expand=True, padx=8, pady=8)
        data = []
        for kind, label, values in rows:
            shown = [str(v) if v else "-" for v in values]
            if kind == "branch":
                data.append((("    " + label, *shown), ()))
            else:
                data.append(((label, *shown), (kind,)))
        self.fill(self.branch_tree, data)

    def show_split(self, result):
        if self.split_holder is not None:
            self.split_holder.destroy()
            self.split_holder = None
        if not result.split_blocks:
            self.split_area.pack_forget()
            self.split_hint.pack(padx=14, pady=14, anchor="w")
            return
        self.split_hint.pack_forget()
        self.split_area.pack(fill="both", expand=True)
        blocks = []
        for b in result.split_blocks:
            if b["error"]:
                blocks.append(((b["title"], "", "", "", b["error"]), ()))
            else:
                blocks.append(((b["title"], b["users"], money(b["excl"]), money(b["vat"]), money(b["incl"])), ()))
        self.fill(self.block_tree, blocks)
        columns = [("name", "Business unit / branch")] + [(f"c{i}", h) for i, h in enumerate(result.split_heads)]
        widths = [200] + [112] * len(result.split_heads)
        self.split_holder, self.split_tree = self.table(self.split_bottom, columns, widths, ["w"] + ["e"] * len(result.split_heads), height=12)
        self.split_holder.pack(fill="both", expand=True, padx=6, pady=6)
        rows = []
        for kind, label, values in result.split_rows:
            shown = [money(v) if v else "-" for v in values]
            rows.append(((("    " + label) if kind == "branch" else label, *shown), () if kind == "branch" else (kind,)))
        self.fill(self.split_tree, rows)

    def show_result(self, result):
        stats = result.stats
        for key in self.cards:
            self.cards[key].configure(text=f"{stats.get(key, '-'):,}" if isinstance(stats.get(key), int) else "-")
        if stats:
            self.cards["invoices"].configure(text=f"{stats['invoices']}" + (f" of {stats['invoices_total']}" if stats["invoices_total"] != stats["invoices"] else ""))
        entity_rows = []
        for name, now, before in result.entities:
            change = "" if before is None else (f"{now - before:+d}" if now != before else "-")
            entity_rows.append(((name, now, "" if before is None else before, change), ("zero",) if now == 0 else ()))
        self.fill(self.entity_tree, entity_rows)
        self.fill(self.other_tree, [((name, total, vs, sma, bpk), ()) for name, total, vs, sma, bpk in result.others])
        if result.branch_rows:
            self.show_branch_table(result.branch_heads, result.branch_rows)
        self.show_split(result)
        items = list(result.notes) + list(result.warnings) + list(result.invoice_notes)
        self.render_checks(result.errors, items)
        self.set_full(result.text)
        attention = sum(1 for label, _ in [categorize(e, True) for e in result.errors] + [categorize(i) for i in items]
                        if label in ("PROBLEM", "CHECK", "LEFT OUT", "ASSUMPTION"))
        if result.errors:
            self.banner("error", "The workbook could not be created", f"{len(result.errors)} problem(s) must be fixed first. See the 'Please check' tab.")
            self.notebook.select(self.tab_checks)
        elif result.out_path:
            detail = result.out_path + (f"\n{attention} item(s) to look at on the 'Please check' tab." if attention else "")
            self.banner("ok", f"Done. The updated workbook for {result.title} was saved.", detail)
            self.notebook.select(self.tab_overview)

    def remember(self):
        self.settings["list"] = self.list_var.get()
        self.settings["book"] = self.book_var.get()
        save_settings(self.settings)

    def pick_list(self):
        start = os.path.dirname(self.list_var.get()) or None
        path = filedialog.askopenfilename(title="Select the user list", initialdir=start, filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.list_var.set(path)
            month, year = process.guess_month(path)
            if month:
                self.month_var.set(MONTHS[month - 1])
            if year:
                self.year_var.set(str(year))

    def pick_book(self):
        start = os.path.dirname(self.book_var.get()) or None
        path = filedialog.askopenfilename(title="Select the latest 2. Beperk Verhalings workbook", initialdir=start, filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.book_var.set(path)

    def month_year(self):
        return MONTHS.index(self.month_var.get()) + 1, int(self.year_var.get())

    def busy(self, on):
        self.create_btn.configure(state="disabled" if on else "normal")
        self.configure(cursor="watch" if on else "")
        self.update_idletasks()

    def subscriptions(self):
        path = self.list_var.get().strip()
        if os.path.isfile(path):
            stamp = (path, os.path.getmtime(path))
            if self.group_cache[0] != stamp:
                try:
                    data = summary.read_user_data(path)
                    if not data.errors:
                        self.group_cache = (stamp, [g[0] for g in data.groups])
                except Exception:
                    pass
        return self.group_cache[1]

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        month, year = self.month_year()
        for i, inv in enumerate(self.invoices):
            used = invoice.in_month(inv, month, year)
            self.tree.insert("", "end", iid=str(i), tags=() if used else ("unused",), values=(
                inv.subscription or "(choose)", inv.number, inv.date.strftime("%d/%m/%Y") if inv.date else inv.date_text,
                inv.currency, f"{inv.incl:,.2f}", "Yes" if used else "NO"))

    def add_invoices(self):
        start = os.path.dirname(self.settings.get("invoices", "")) or None
        paths = filedialog.askopenfilenames(title="Select the invoice PDFs", initialdir=start, filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        self.settings["invoices"] = paths[0]
        self.remember()
        self.busy(True)
        notes = []
        try:
            for n, path in enumerate(paths, 1):
                self.banner("working", f"Reading invoice {n} of {len(paths)}", os.path.basename(path))
                inv = invoice.parse_invoice(path)
                self.invoices.append(inv)
                for note in inv.notes:
                    notes.append(f"{inv.file}: {note}")
            self.refresh_tree()
            month, year = self.month_year()
            items = notes + invoice.invoice_warnings(self.invoices, month, year)
            self.render_checks([], items)
            self.banner("info", f"{len(paths)} invoice(s) added",
                        "Check the amounts in the list above (double-click an invoice to change it). Notes about the invoices are on the 'Please check' tab.")
            self.notebook.select(self.tab_checks)
        except Exception:
            self.say("Something went wrong while reading the invoices:\n\n" + traceback.format_exc(), "error")
        finally:
            self.busy(False)

    def edit_invoice(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Edit invoice", "Click an invoice in the list first.")
            return
        inv = self.invoices[int(selected[0])]
        dialog = InvoiceDialog(self, inv, self.subscriptions())
        self.wait_window(dialog)
        if dialog.result:
            self.refresh_tree()

    def remove_invoices(self):
        selected = sorted((int(i) for i in self.tree.selection()), reverse=True)
        for i in selected:
            del self.invoices[i]
        self.refresh_tree()

    def run_create(self):
        list_path = self.list_var.get().strip()
        book_path = self.book_var.get().strip()
        if not os.path.isfile(list_path):
            messagebox.showwarning("Missing file", "Please choose the user list file.")
            return
        if not os.path.isfile(book_path):
            messagebox.showwarning("Missing file", "Please choose the latest '2. Beperk Verhalings' workbook.")
            return
        month, year = self.month_year()
        self.remember()
        self.banner("working", "Reading the files and calculating...", "This can take a few seconds.")
        self.busy(True)
        try:
            result = process.run_full(list_path, book_path, month, year, self.invoices)
            if result.out_path:
                self.last_output = result.out_path
                self.open_btn.configure(state="normal")
            self.show_result(result)
        except PermissionError:
            self.say("A file could not be opened or saved.", "error", "If it is open in Excel, close it and try again.")
        except Exception:
            self.say("Something went wrong:\n\n" + traceback.format_exc(), "error")
        finally:
            self.busy(False)

    def open_folder(self):
        if self.last_output and os.path.exists(self.last_output):
            os.startfile(os.path.dirname(self.last_output))


def selftest(args):
    list_path, book_path, month, year, out = args[0], args[1], int(args[2]), int(args[3]), args[4]
    invoices = [invoice.parse_invoice(p) for p in args[5:]]
    text, saved = process.run(list_path, book_path, month, year, invoices, out)
    print(text)
    print("WROTE", saved)


if __name__ == "__main__":
    if len(sys.argv) >= 6 and sys.argv[1] == "--selftest":
        selftest(sys.argv[2:])
    else:
        App().mainloop()
