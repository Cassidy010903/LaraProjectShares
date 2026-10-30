import re
import os
import datetime
import pdfplumber

SUPPLIERS = [
    ("vodacom", "Vodacom", "Microsoft 365"),
    ("canva", "Canva", "Canva"),
    ("adobe", "Adobe", "Adobe"),
    ("openai", "OpenAI", "Chat GPT"),
    ("chatgpt", "OpenAI", "Chat GPT"),
    ("freshworks", "Freshworks", "Fresh Service Agent"),
    ("freshservice", "Freshworks", "Fresh Service Agent"),
    ("read.ai", "Read AI", "Read AI"),
    ("read ai", "Read AI", "Read AI"),
    ("microsoft", "Microsoft", "Microsoft 365"),
]

MONTH_NAMES = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}

MONTH_NAMES_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

NUMBER = re.compile(r"(?<![\d.,])(\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2})(?!\d)")


class Invoice:
    def __init__(self, path):
        self.path = path
        self.file = os.path.basename(path)
        self.supplier = ""
        self.subscription = ""
        self.number = ""
        self.date = None
        self.date_text = ""
        self.currency = "ZAR"
        self.excl = 0.0
        self.vat = 0.0
        self.incl = 0.0
        self.notes = []
        self.found = False
        self.lines = {}
        self.components = []
        self.usd = None
        self.rate = None
        self.force_include = False
        self.counted = True
        self.assumptions = []
        self.unit_prices = {}
        self.standard_seats = 0
        self.premium_seats = 0


def to_float(text):
    return float(text.replace(",", ""))


def half_up(value):
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(repr(round(value, 8))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def in_month(inv, month, year):
    return inv.force_include or inv.date is None or (inv.date.year, inv.date.month) == (year, month)


def read_pdf_text(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join((page.extract_text() or "") for page in pdf.pages)
    text = re.sub(r"[    ]", " ", text)
    return text


def plain(text):
    return re.sub(r"[^\x20-\x7e\n]", "", text)


def dashed(text):
    return plain(re.sub(r"[\x00�]", "-", text))


def find_triple(values):
    best = None
    unique = sorted(set(values))
    for a in unique:
        if a <= 0:
            continue
        vat = round(a * 0.15, 2)
        for b in unique:
            if abs(b - vat) > 0.02:
                continue
            for c in unique:
                if abs(c - (a + b)) <= 0.015 and (best is None or c > best[2]):
                    best = (a, b, c)
    return best


def parse_date(text):
    match = re.search(r"(\d{1,2})/(\d{1,2})/(20\d\d)", text)
    if match:
        d, m, y = (int(x) for x in match.groups())
        try:
            return datetime.date(y, m, d), match.group(0)
        except ValueError:
            pass
    match = re.search(r"\b(\d{1,2})-([A-Za-z]{3})-(20\d\d)", text)
    if match and match.group(2).lower() in MONTH_NAMES:
        try:
            return datetime.date(int(match.group(3)), MONTH_NAMES[match.group(2).lower()], int(match.group(1))), match.group(0)
        except ValueError:
            pass
    match = re.search(r"\b([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(20\d\d)", text)
    if match and match.group(1)[:3].lower() in MONTH_NAMES:
        try:
            return datetime.date(int(match.group(3)), MONTH_NAMES[match.group(1)[:3].lower()], int(match.group(2))), match.group(0)
        except ValueError:
            pass
    return None, ""


def find_invoice_date(text):
    labels = r"(?:Tax Invoice Date|Invoice Date|Date of issue|Document Date|Date)"
    for match in re.finditer(labels + r"[ \t:]+([^\n]{0,40})", text, re.I):
        date, shown = parse_date(match.group(1))
        if date:
            return date, shown
    return parse_date(text)


def find_invoice_number(text):
    match = re.search(r"Invoice[ \t]*(?:Number|No\.?|#)[ \t]*[:#]?[ \t]*([A-Za-z0-9][A-Za-z0-9\-/]+)(?:[ \t]([0-9]{3,5})\b)?", text, re.I)
    if not match:
        match = re.search(r"Billing Number[ \t]+([A-Za-z0-9\-]+)", text, re.I)
        return match.group(1) if match else ""
    number = match.group(1)
    if match.group(2):
        number += " " + match.group(2)
    return number


def detect_currency(text, incl):
    shown = f"{incl:,.2f}"
    for line in text.splitlines():
        if shown in line:
            if re.search(r"USD|\$", line):
                return "USD"
            if re.search(r"ZAR|\bR\s?\d", line):
                return "ZAR"
    if re.search(r"USD|\$", text) and not re.search(r"ZAR|\bR\s?\d", text):
        return "USD"
    return "ZAR"


def parse_invoice(path):
    inv = Invoice(path)
    try:
        raw = read_pdf_text(path)
        text = plain(raw)
    except Exception as e:
        inv.notes.append(f"The PDF could not be read ({type(e).__name__}). Please type the amounts in.")
        return inv
    lower = text.lower()
    for keyword, supplier, subscription in SUPPLIERS:
        if keyword in lower:
            inv.supplier, inv.subscription = supplier, subscription
            break
    if inv.supplier == "Adobe":
        for name, key in (("Acrobat Standard DC", "std"), ("Acrobat Pro", "pro"), ("Creative Cloud", "cc")):
            found = re.search(re.escape(name) + r"[A-Za-z ]*?\s+(\d+)\s+EA\s+([\d,]+\.\d{2})", text)
            if found:
                inv.lines[key] = int(found.group(1))
                inv.unit_prices[key] = to_float(found.group(2))
    if inv.supplier == "Vodacom":
        for found in re.finditer(r"([^\n]*Professional Services)\([\d/]+\s*-\s+([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}", text):
            amount = to_float(found.group(2))
            following = text[found.end():found.end() + 90]
            entity = "SUBTROPICO INTERNATIONAL PTA" if "INTERNATIONAL" in following else "SUBTROPICO"
            inv.assumptions.append(
                f"ASSUMED: this Vodacom invoice includes R{amount:,.2f} excl VAT for 'Office 365 Professional Services' billed to {entity}. "
                "It is not a monthly licence charge, and nothing in the user list or the other invoices explains it. "
                "Until it is known where it belongs, it is split over the users like the rest of the Microsoft 365 amount (an assumption).")
        inv.notes += inv.assumptions
    if inv.supplier == "OpenAI":
        found = re.search(r"Remaining time on (\d+)\s+ChatGPT Business Subscription", text)
        if found:
            inv.standard_seats = int(found.group(1))
        found = re.search(r"Remaining time on ChatGPT Business Premium Subscription[^\n]*?\s(\d+)\s+15%", text)
        if found:
            inv.premium_seats = int(found.group(1))
    inv.number = find_invoice_number(dashed(raw))
    inv.date, inv.date_text = find_invoice_date(text)
    zar = re.search(r"Total Charges \(excluding VAT\)\s+ZAR\s+([\d,]+\.\d{2})\s+Total VAT\s+ZAR\s+([\d,]+\.\d{2})\s+Total Charges \(including VAT\)\s+ZAR\s+([\d,]+\.\d{2})", text)
    triple = None
    if zar:
        triple = tuple(to_float(x) for x in zar.groups())
        inv.currency = "ZAR"
        rate = re.search(r"1\s*USD\s*=\s*([\d.]+)\s*ZAR", text)
        inv.notes.append("Rand amounts taken from the tax information on the invoice" + (f" (1 USD = {float(rate.group(1)):.4f} ZAR)." if rate else "."))
    if triple is None:
        rand_numbers = [to_float(m.group(1)) for m in re.finditer(r"\bR\s?(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})", text)]
        triple = find_triple(rand_numbers) if rand_numbers else None
        if triple:
            inv.currency = "ZAR"
            inv.notes.append("Rand amounts taken from the invoice (it is billed in another currency but shows rand amounts).")
    if triple is None:
        values = [to_float(m.group(1)) for m in NUMBER.finditer(text)]
        triple = find_triple(values)
        if triple:
            inv.currency = detect_currency(text, triple[2])
            if inv.currency == "USD":
                converted = re.search(r"VAT[^\n]*\n\s*R\s?([\d,]+\.\d{2})", text)
                if converted and triple[1] > 0:
                    zar_vat = to_float(converted.group(1))
                    rate = zar_vat / triple[1]
                    inv.usd = triple
                    inv.rate = rate
                    excl_zar = half_up(triple[0] * rate)
                    triple = (excl_zar, zar_vat, half_up(excl_zar + zar_vat))
                    inv.currency = "ZAR"
                    inv.notes.append(f"Rand amounts calculated with the exchange rate printed on the invoice (VAT ${inv.usd[1]:,.2f} = R{zar_vat:,.2f}, so 1 USD = R{rate:.4f}). "
                                     f"Dollar amounts: excl {inv.usd[0]:,.2f}, VAT {inv.usd[1]:,.2f}, total {inv.usd[2]:,.2f}.")
                else:
                    inv.notes.append("This invoice is in US dollars and shows no rand amount. The split will be in USD, or type the rand amount from the bank statement.")
    if triple:
        inv.excl, inv.vat, inv.incl = triple
        inv.found = True
        import components
        inv.components = components.build(inv, text)
    else:
        inv.notes.append("The amounts could not be found automatically. Please type them in.")
    return inv


def allocate(total_cents, weights):
    total_weight = sum(weights)
    if total_weight <= 0:
        return [0] * len(weights)
    base = [(2 * total_cents * w + total_weight) // (2 * total_weight) for w in weights]
    difference = total_cents - sum(base)
    biggest = max(range(len(weights)), key=lambda i: (weights[i], -i))
    base[biggest] += difference
    return base


def cents(value):
    return int(round(float(value) * 100))


class Split:
    def __init__(self, subscription, currency):
        self.subscription = subscription
        self.currency = currency
        self.licence = ""
        self.licence_key = None
        self.note = ""
        self.invoices = []
        self.excl = 0.0
        self.vat = 0.0
        self.incl = 0.0
        self.users = {}
        self.total_users = 0
        self.rows = {}
        self.error = ""


def invoice_pieces(inv):
    if inv.components:
        return inv.components
    return [{"key": None, "label": "", "excl": inv.excl, "vat": inv.vat, "incl": inv.incl}]


def build_splits(data, invoices):
    groups = collections_ordered()
    for inv in invoices:
        for piece in invoice_pieces(inv):
            block = groups.setdefault((inv.subscription, piece["key"], inv.currency), {"invoices": [], "excl": 0.0, "vat": 0.0, "incl": 0.0})
            if inv not in block["invoices"]:
                block["invoices"].append(inv)
            block["excl"] += piece["excl"]
            block["vat"] += piece["vat"]
            block["incl"] += piece["incl"]
    splits = []
    names = [g[0] for g in data.groups]
    for (subscription, key, currency), block in groups.items():
        split = Split(subscription, currency)
        split.invoices = block["invoices"]
        split.excl = round(block["excl"], 2)
        split.vat = round(block["vat"], 2)
        split.incl = round(block["incl"], 2)
        if subscription not in names:
            split.error = f"'{subscription or '(none)'}' is not a subscription in the user list."
            splits.append(split)
            continue
        members = dict(data.groups)[subscription]
        if key and key in data.keys:
            split.licence_key = key
            split.licence = data.detail_name[key]
            split.users = {k: data.counts[k][key] for k in data.order}
            if sum(split.users.values()) <= 0:
                split.note = (f"{subscription}: nobody on the user list has {split.licence}, so that amount was divided over all {subscription} users instead.")
                split.licence_key, split.licence = None, ""
                split.users = {k: data.group_count(k, subscription) for k in data.order}
        else:
            split.users = {k: data.group_count(k, subscription) for k in data.order}
            if len(members) > 1:
                split.licence = "all licence types"
        split.total_users = sum(split.users.values())
        if split.total_users <= 0:
            split.error = f"No users have {subscription} on the user list, so the invoice cannot be split."
            splits.append(split)
            continue
        weights = [split.users[key] for key in data.order]
        excl = allocate(cents(split.excl), weights)
        vat = allocate(cents(split.vat), weights)
        for key, e, v in zip(data.order, excl, vat):
            split.rows[key] = (e / 100, v / 100, (e + v) / 100)
        splits.append(split)
    return splits


def collections_ordered():
    import collections
    return collections.OrderedDict()


def invoice_warnings(invoices, month, year):
    out = []
    seen = {}
    for inv in invoices:
        label = f"{inv.file}"
        if not inv.subscription:
            out.append(f"{label}: no subscription chosen.")
        if inv.incl and abs(inv.excl + inv.vat - inv.incl) > 0.02:
            out.append(f"{label}: Excl VAT + VAT does not add up to the total ({inv.excl:,.2f} + {inv.vat:,.2f} is not {inv.incl:,.2f}).")
        if not in_month(inv, month, year):
            out.append(f"{label}: NOT USED. The invoice is dated {inv.date.strftime('%d/%m/%Y')}, which is not in {MONTH_NAMES_FULL[month - 1]} {year}, "
                       "so it is left out of the split and the workbook figures. (Double-click it and tick 'Use anyway' to include it.)")
        if inv.number:
            if inv.number in seen:
                out.append(f"{label}: invoice number {inv.number} was already added ({seen[inv.number]}).")
            seen[inv.number] = inv.file
    return out
