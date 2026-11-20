import io, json, os, datetime
import requests
from openpyxl import load_workbook

PUBLIC_URL = os.environ.get("PUBLIC_URL", "https://disk.yandex.ru/d/OqpN0fbaPia95A")
API = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
SNAP = "data/latest.json"
CHANGES = "data/changes.json"


def fetch_xlsx() -> bytes:
    local = os.environ.get("LOCAL_XLSX")  # для теста без сети
    if local:
        return open(local, "rb").read()
    r = requests.get(API, params={"public_key": PUBLIC_URL}, timeout=30)
    r.raise_for_status()
    f = requests.get(r.json()["href"], timeout=120)
    f.raise_for_status()
    return f.content


def norm(v):
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return v.isoformat()
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def read_cells(data: bytes) -> dict:
    wb = load_workbook(io.BytesIO(data), data_only=True)
    out = {}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and str(c.value).strip() != "":
                    cells[c.coordinate] = norm(c.value)
        out[ws.title] = cells
    return out


def diff(old: dict, new: dict) -> list:
    res = []
    for sheet in sorted(set(old) | set(new)):
        o, n = old.get(sheet, {}), new.get(sheet, {})
        for cell in sorted(set(o) | set(n)):
            if o.get(cell) != n.get(cell):
                res.append({"sheet": sheet, "cell": cell,
                            "old": o.get(cell), "new": n.get(cell)})
    return res


def main():
    os.makedirs("data", exist_ok=True)
    new = read_cells(fetch_xlsx())
    first = not os.path.exists(SNAP)
    old = {} if first else json.load(open(SNAP, encoding="utf-8"))
    changes = [] if first else diff(old, new)
    now = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")
    json.dump({"checked_at_utc": now, "first_run": first,
               "has_changes": bool(changes), "changes": changes},
              open(CHANGES, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    json.dump(new, open(SNAP, "w", encoding="utf-8"), ensure_ascii=False,
              indent=1, sort_keys=True)
    print(f"sheets={len(new)} cells={sum(len(v) for v in new.values())} "
          f"changes={len(changes)} first_run={first}")


if __name__ == "__main__":
    main()
