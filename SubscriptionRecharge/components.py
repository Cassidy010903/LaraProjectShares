import re

MICROSOFT_PRODUCTS = [
    (r"power bi premium", "microsoft power bi premium"),
    (r"power bi pro", "microsoft power bi pro"),
    (r"teams premium", "teams premium"),
    (r"copilot", "copilot"),
    (r"business premium", "business premium"),
    (r"business standard", "business standard"),
    (r"\bf3\b", "f3"),
]


def to_float(text):
    return float(text.replace(",", ""))


def piece(key, label, excl, vat=None):
    return {"key": key, "label": label, "excl": excl, "vat": vat, "incl": None}


def microsoft_key(text):
    lower = text.lower()
    keys = {key for pattern, key in MICROSOFT_PRODUCTS if re.search(pattern, lower)}
    return keys.pop() if len(keys) == 1 else None


def adobe(text):
    found = []
    for name, key in (("Acrobat Standard DC", "adobe acrobat standard dc"), ("Acrobat Pro", "adobe acrobat pro"), ("Creative Cloud", "adobe creative suite")):
        line = re.search(re.escape(name) + r"[A-Za-z ]*?\s+(\d+)\s+EA\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+15\.00%\s+([\d,]+\.\d{2})", text)
        if line:
            found.append(piece(key, name, to_float(line.group(2)), to_float(line.group(3))))
    return found


def vodacom(text):
    header = re.compile(r"^[SU]\d+/\d+\s.*?([\d,]+\.\d{2})\s+0\.00\s+([\d,]+\.\d{2})\s*$")
    blocks = []
    current = None
    for line in text.splitlines():
        stripped = line.strip()
        match = header.match(stripped)
        if match:
            current = {"amount": to_float(match.group(2)), "text": stripped}
            blocks.append(current)
            continue
        if current is not None:
            if stripped.startswith("Total Charges"):
                current = None
                continue
            current["text"] += "\n" + stripped
    found = []
    for block in blocks:
        lower = block["text"].lower()
        if "professional services" in lower:
            found.append(piece(None, "Office 365 Professional Services", block["amount"]))
        elif "business premium" in lower:
            found.append(piece("business premium", "M365 Business Premium SaaS Backup Bundle", block["amount"]))
        elif "business standard" in lower:
            found.append(piece("business standard", "M365 Business Standard SaaS Backup Bundle", block["amount"]))
        elif re.search(r"\bf3\b", lower):
            found.append(piece("f3", "Office 365 F3", block["amount"]))
        elif "keepit" in lower:
            found.append(piece(None, "Keepit for Microsoft 365", block["amount"]))
        else:
            found.append(piece(None, "Other Vodacom line", block["amount"]))
    return found


def build(inv, text):
    import invoice
    pieces = []
    if inv.supplier == "Adobe":
        pieces = adobe(text)
    elif inv.supplier == "OpenAI":
        pieces = [piece("chat gpt premium" if "premium" in text.lower() else "chat gpt", "ChatGPT", inv.excl, inv.vat)]
    elif inv.supplier == "Microsoft":
        key = microsoft_key(text)
        if key:
            pieces = [piece(key, "Microsoft licence", inv.excl, inv.vat)]
    elif inv.supplier == "Vodacom":
        pieces = vodacom(text)
    if not pieces or abs(sum(p["excl"] for p in pieces) - inv.excl) > 0.02:
        return []
    vat_given = all(p["vat"] is not None for p in pieces)
    if not vat_given or abs(sum(p["vat"] for p in pieces) - inv.vat) > 0.02:
        shares = invoice.allocate(invoice.cents(inv.vat), [invoice.cents(p["excl"]) for p in pieces])
        for p, share in zip(pieces, shares):
            p["vat"] = share / 100
    for p in pieces:
        p["incl"] = round(p["excl"] + p["vat"], 2)
    return pieces
