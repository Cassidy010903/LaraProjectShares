import json
import urllib.request

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Sully/1.0", "Accept": "application/json"}
URL = "https://api.frankfurter.dev/v1/{date}?base={base}&symbols=ZAR"
_cache = {}


def daily_rate(day, base="USD"):
    key = (day.isoformat(), base)
    if key in _cache:
        return _cache[key]
    result = None
    try:
        request = urllib.request.Request(URL.format(date=day.isoformat(), base=base), headers=HEADERS)
        with urllib.request.urlopen(request, timeout=8) as response:
            payload = json.loads(response.read().decode("utf-8"))
        result = (float(payload["rates"]["ZAR"]), payload.get("date", day.isoformat()))
    except Exception:
        result = None
    _cache[key] = result
    return result
