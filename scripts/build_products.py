#!/usr/bin/env python3
"""Build products_full.json (+ product images) from the product spreadsheet.

Source is a public Google Sheet exported as XLSX (via --sheet-id / SHEET_ID env)
or a local .xlsx file (via --xlsx). Columns are resolved by header name, not
position, because the sheet's column order has changed between versions.

Image downloads are incremental: the Drive link for a product changes whenever
its image changes, so we remember each row's link in sync-manifest.json and only
re-download an image when its link changed (or the file is missing). The full
products_full.json is rewritten every run; rely on git to decide whether anything
actually changed before committing.
"""
import argparse
import hashlib
import io
import json
import os
import re
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_PATH = os.path.join(ROOT, "products_full.json")
MANIFEST_PATH = os.path.join(ROOT, "sync-manifest.json")
IMG_REL_DIR = os.path.join("uploads", "products")
IMG_DIR = os.path.join(ROOT, IMG_REL_DIR)

UA = "Mozilla/5.0 (compatible; potachips-sync/1.0)"

# spreadsheet header text -> canonical field name
HEADER_MAP = {
    "STT": "stt",
    "Tên sản phẩm": "name",
    "BRAND": "brand",
    "GIÁ NIÊM YẾT": "shopee_price",
    "GIÁ BÁN WEBSITE": "web_price",
    "Ảnh sản phẩm": "img_url",
    "Nhóm sản phẩm": "category",
    "Flashsale": "flashsale",
    "Link shopee": "shopee_link",
}
REQUIRED = ("stt", "name", "brand", "shopee_price", "web_price", "img_url", "category")


def fetch_xlsx(sheet_id):
    url = "https://docs.google.com/spreadsheets/d/%s/export?format=xlsx" % sheet_id
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return io.BytesIO(r.read())


def load_rows(src):
    import openpyxl

    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb.active
    rows = list(ws.values)
    if not rows:
        raise SystemExit("empty spreadsheet")
    header = rows[0]
    idx = {}
    for i, h in enumerate(header):
        key = HEADER_MAP.get(str(h).strip() if h is not None else "")
        if key and key not in idx:
            idx[key] = i
    missing = [k for k in REQUIRED if k not in idx]
    if missing:
        raise SystemExit("Missing columns %s. Sheet headers: %s" % (missing, list(header)))
    out = []
    for r in rows[1:]:
        if r[idx["stt"]] is None:
            continue
        out.append({k: r[i] for k, i in idx.items()})
    return out


def num(v):
    return int(v) if isinstance(v, (int, float)) else 0


def as_bool(v):
    if v is True:
        return True
    return str(v).strip().lower() in ("true", "1", "yes", "x", "có")


def drive_key(link):
    """Stable identity for a Drive link (the file id), used for change detection."""
    s = str(link or "")
    m = re.search(r"/d/([A-Za-z0-9_-]+)", s) or re.search(r"[?&]id=([A-Za-z0-9_-]+)", s)
    if m:
        return m.group(1)
    return "sha:" + hashlib.sha1(s.encode("utf-8")).hexdigest()[:16] if s else ""


def download_image(link, dest):
    fid = drive_key(link)
    if not fid or fid.startswith("sha:"):
        raise ValueError("no drive file id in link: %s" % link)
    url = "https://drive.google.com/uc?export=download&id=%s" % fid
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=120) as r:
        data = r.read()
    # large files return an HTML virus-scan interstitial with a confirm token
    low = data[:400].lower()
    if low.startswith(b"<!doctype html") or b"<html" in low:
        m = re.search(rb"confirm=([0-9A-Za-z_-]+)", data)
        if m:
            url2 = ("https://drive.usercontent.google.com/download?id=%s"
                    "&export=download&confirm=%s") % (fid, m.group(1).decode())
            req2 = urllib.request.Request(url2, headers={"User-Agent": UA})
            with urllib.request.urlopen(req2, timeout=120) as r:
                data = r.read()
    if not (data[:8].startswith(b"\x89PNG") or data[:3] == b"\xff\xd8\xff"):
        raise ValueError("not a PNG/JPEG (%d bytes)" % len(data))
    tmp = dest + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, dest)
    return len(data)


def build(rows, force_images=False, skip_images=False):
    os.makedirs(IMG_DIR, exist_ok=True)
    old = {}
    if os.path.exists(MANIFEST_PATH):
        old = json.load(open(MANIFEST_PATH)).get("rows", {})

    products = []
    manifest_rows = {}
    downloaded = failed = 0
    seen = set()

    for r in rows:
        pid = num(r["stt"])
        seen.add(pid)
        web = num(r["web_price"])
        shopee = num(r["shopee_price"])
        shopee = shopee if shopee > 100 else 0
        save = round((shopee - web) / shopee * 100) if (shopee > 0 and web > 0) else 0
        link = str(r.get("shopee_link") or "").strip()
        img_url = str(r["img_url"] or "").strip()
        img_rel = "%s/%s.png" % (IMG_REL_DIR.replace(os.sep, "/"), pid)

        products.append({
            "id": pid,
            "name": str(r["name"] or "").strip(),
            "brand": str(r["brand"] or "").strip(),
            "category": str(r["category"] or "").strip(),
            "webPrice": web,
            "shopeePrice": shopee,
            "savePct": save,
            "shopeeLink": link if link.startswith("http") else "",
            "img": img_rel,
            "flashSale": as_bool(r.get("flashsale")),
        })

        key = drive_key(img_url)
        dest = os.path.join(ROOT, img_rel)
        prev_key = old.get(str(pid), {}).get("img")
        need = force_images or prev_key != key or not os.path.exists(dest)
        downloaded_ok = False
        if need and img_url and not skip_images:
            try:
                download_image(img_url, dest)
                downloaded_ok = True
                downloaded += 1
            except Exception as e:
                print("WARN id=%s image failed: %s" % (pid, e), file=sys.stderr)
                failed += 1
        # Only record the link as current when the on-disk file provably matches
        # it: either we just downloaded it, or it was already current and exists.
        if downloaded_ok or (prev_key == key and os.path.exists(dest)):
            stored_key = key
        else:
            stored_key = ""  # unknown/stale -> forces re-download next real run
        manifest_rows[str(pid)] = {"img": stored_key}

    removed = 0
    for old_id in old:
        if int(old_id) not in seen:
            p = os.path.join(IMG_DIR, "%s.png" % old_id)
            if os.path.exists(p):
                os.remove(p)
                removed += 1

    products.sort(key=lambda p: p["id"])
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, separators=(",", ":"))
    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump({"rows": manifest_rows}, f, ensure_ascii=False, indent=2, sort_keys=True)

    print("products=%d images_downloaded=%d images_failed=%d images_removed=%d"
          % (len(products), downloaded, failed, removed))
    return failed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet-id")
    ap.add_argument("--xlsx", help="local .xlsx instead of fetching the sheet")
    ap.add_argument("--force-images", action="store_true")
    ap.add_argument("--skip-images", action="store_true", help="build JSON only")
    a = ap.parse_args()

    if a.xlsx:
        src = a.xlsx
    else:
        sid = a.sheet_id or os.environ.get("SHEET_ID")
        if not sid:
            raise SystemExit("provide --xlsx, --sheet-id, or SHEET_ID env")
        src = fetch_xlsx(sid)

    rows = load_rows(src)
    failed = build(rows, force_images=a.force_images, skip_images=a.skip_images)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
