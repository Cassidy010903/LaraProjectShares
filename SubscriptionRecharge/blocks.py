import fx

DATE_STYLE = {"Adobe": "%d.%m.%y", "Read AI": "%d/%m/%y", "Chat GPT": "%d/%m/%y"}


def short_date(day, subscription):
    return day.strftime(DATE_STYLE[subscription]) if day else None


def join_numbers(numbers):
    numbers = [n for n in numbers if n]
    if len(numbers) > 1 and all("-" in n for n in numbers):
        prefixes = {n.rsplit("-", 1)[0] for n in numbers}
        if len(prefixes) == 1:
            return numbers[0] + "/" + "/".join(n.rsplit("-", 1)[1] for n in numbers[1:])
    return "/".join(numbers)


def prepare(invoices):
    figures = {}
    notes = []
    for subscription in ("Adobe", "Read AI", "Chat GPT"):
        items = [i for i in invoices if i.subscription == subscription]
        if not items:
            continue
        dated = [i.date for i in items if i.date]
        figure = {"date": short_date(min(dated), subscription) if dated else None,
                  "number": join_numbers([i.number for i in items]), "usd": None, "rand": None, "estimated": False}
        if subscription == "Adobe":
            prices = {}
            for inv in items:
                prices.update(inv.unit_prices)
            figure["prices"] = prices
            figures[subscription] = figure
            continue
        usd_total = 0.0
        rand_total = 0.0
        complete = True
        used = []
        exact = []
        for inv in items:
            if inv.usd:
                usd_total += inv.usd[2]
                rand_total += inv.incl
                exact.append(inv.rate)
            elif inv.currency == "USD":
                usd_total += inv.incl
                looked_up = fx.daily_rate(inv.date) if inv.date else None
                if looked_up:
                    rand_total += inv.incl * looked_up[0]
                    figure["estimated"] = True
                    used.append(f"{looked_up[0]:.4f} on {inv.date.strftime('%d/%m/%Y')}")
                else:
                    complete = False
            elif inv.currency == "ZAR":
                rand_total += inv.incl
        if usd_total:
            figure["usd"] = round(usd_total, 2)
        if complete:
            figure["rand"] = round(rand_total, 2)
        figures[subscription] = figure
        if figure["estimated"] and complete:
            notes.append(f"{subscription}: the rand total on the '2026' sheet (R{figure['rand']:,.2f}) is an ESTIMATE. The dollar invoices (USD {figure['usd']:,.2f} incl VAT) "
                         f"were converted at the daily ECB exchange rate ({'; '.join(used)}). The real rand amount is on the bank statement and was, in past months, "
                         "about 2.5% to 3.5% higher (the bank's own rate). To use the real amount, double-click the invoice, set Currency to ZAR and type the rand amounts.")
        if exact:
            notes.append(f"{subscription}: the rand total on the '2026' sheet (R{figure['rand']:,.2f}) is calculated from the exchange rate printed on each of the "
                         f"{len(exact)} dollar invoice(s) (rates {min(exact):.4f} to {max(exact):.4f}), so it is not an estimate.")
        if usd_total and not complete:
            notes.append(f"{subscription}: the exchange rate could not be looked up (no internet?), so the rand total on the '2026' sheet was not changed. "
                         "Type the rand amounts from the bank statement into the invoice (Currency ZAR) to fill it in.")
        if subscription == "Chat GPT":
            notes.append(f"Chat GPT: the figures filled in are the {len(items)} OpenAI invoice(s) you uploaded. If the month's regular renewal invoice is not one of them, upload it too.")
    return figures, notes
