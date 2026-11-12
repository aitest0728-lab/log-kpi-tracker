#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download the ECOM-MFNAS Accumulated Delivery Lead Time (Order Date) report from TKO-MFNAS."""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.environ.get("MFNAS_REPORT_DIR", HERE)

NAS_BASE_URL = "https://tko-mfnas.hktv.com.hk:5001"
NAS_REPORT_FOLDER = "ECOM-MFNAS_hktvAccumulatedDeliveryLeadTimeReport"
NAS_ROOT_CANDIDATES = [r for r in (
    os.environ.get("ROLLING_NAS_ROOT", "").strip().rstrip("/"),
    "/ECOM-MFNAS/ECOM-MFNAS",
    "/ECOM-MFNAS",
    "/OIX-MFNAS/OIX-MFNAS",
    "/OIX-MFNAS/OIX-MFNAS/ECOM-MFNAS",
    "/OIX-MFNAS/ECOM-MFNAS",
    "/OIX-MFNAS",
) if r]
NAS_CONNECT_TIMEOUT = 10
NAS_READ_TIMEOUT = 120


def _load_env_file():
    loaded = []
    for folder in (HERE, os.path.dirname(HERE)):
        for name in ("mfnas.env", "rolling_orders.env", "crab.env"):
            path = os.path.join(folder, name)
            if os.path.exists(path):
                with open(path, encoding="utf-8-sig") as fh:
                    for line in fh:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            key, value = line.split("=", 1)
                            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))
                loaded.append(path)
    return loaded


_load_env_file()

NAS_USERNAME = os.environ.get("CRAB_NAS_USER", "")
NAS_PASSWORD = os.environ.get("CRAB_NAS_PASS", "")


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def mmm_dd(d):
    return d.strftime("%b-%d")


def report_path(order_date):
    return os.path.join(OUT_DIR, f"Rolling Order Count ({mmm_dd(order_date)}).csv")


class NASClient:
    """Synology File Station Web API: login, list, download."""

    def __init__(self):
        import requests
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.s = requests.Session()
        self.sid = None

    def _get(self, path, params, **kw):
        return self.s.get(
            f"{NAS_BASE_URL}/webapi/{path}",
            params=params,
            verify=False,
            timeout=(NAS_CONNECT_TIMEOUT, NAS_READ_TIMEOUT),
            **kw,
        )

    def login(self):
        if not NAS_USERNAME or not NAS_PASSWORD:
            raise RuntimeError(
                "CRAB_NAS_USER / CRAB_NAS_PASS not set (env var, mfnas.env, rolling_orders.env, or crab.env)"
            )
        log("MFNAS: login")
        data = self._get("auth.cgi", {
            "api": "SYNO.API.Auth",
            "version": 3,
            "method": "login",
            "account": NAS_USERNAME,
            "passwd": NAS_PASSWORD,
            "session": "FileStation",
            "format": "sid",
        }).json()
        if not data.get("success"):
            err = data.get("error", {})
            raise RuntimeError(f"[NAS] login failed: {err if err else 'unsuccessful response'}")
        self.sid = data["data"]["sid"]
        log("MFNAS: login successful")

    def logout(self):
        if not self.sid:
            return
        try:
            self._get("auth.cgi", {
                "api": "SYNO.API.Auth",
                "version": 1,
                "method": "logout",
                "session": "FileStation",
                "_sid": self.sid,
            })
        except Exception:
            pass
        self.sid = None

    def list_folder(self, folder_path, quiet=False):
        data = self._get("entry.cgi", {
            "api": "SYNO.FileStation.List",
            "version": 2,
            "method": "list",
            "folder_path": folder_path,
            "_sid": self.sid,
        }).json()
        if not data.get("success"):
            if not quiet:
                log(f"[NAS] cannot list '{folder_path}': {data.get('error', data)}")
            return None
        return [{
            "name": f.get("name", ""),
            "path": f.get("path"),
            "isdir": bool(f.get("isdir")),
        } for f in data["data"].get("files", [])]

    def list_shares(self):
        data = self._get("entry.cgi", {
            "api": "SYNO.FileStation.List",
            "version": 2,
            "method": "list_share",
            "_sid": self.sid,
        }).json()
        return [s.get("path") for s in data.get("data", {}).get("shares", [])] if data.get("success") else []

    def download(self, remote_path, local_path):
        log(f"[NAS] downloading {os.path.basename(remote_path)} ...")
        r = self._get("entry.cgi", {
            "api": "SYNO.FileStation.Download",
            "version": 2,
            "method": "download",
            "path": remote_path,
            "mode": "open",
            "_sid": self.sid,
        }, stream=True)
        r.raise_for_status()
        if "json" in (r.headers.get("Content-Type") or "").lower():
            raise RuntimeError(f"[NAS] download refused: {r.text[:300]}")
        os.makedirs(os.path.dirname(local_path) or ".", exist_ok=True)
        total, last = 0, time.time()
        with open(local_path, "wb") as f:
            for chunk in r.iter_content(65536):
                if chunk:
                    f.write(chunk)
                    total += len(chunk)
                    if time.time() - last > 10:
                        log(f"[NAS] ... {total / 1048576:.0f} MB")
                        last = time.time()
        log(f"[NAS] saved {local_path} ({total / 1048576:.1f} MB)")
        return total


def find_date_folder(nas, t0s):
    for root in NAS_ROOT_CANDIDATES:
        path = f"{root}/{NAS_REPORT_FOLDER}/{t0s}"
        if nas.list_folder(path, quiet=True) is not None:
            log(f"MFNAS: folder found: {path}")
            return path
        if nas.list_folder(f"{root}/{NAS_REPORT_FOLDER}", quiet=True) is not None:
            log(f"MFNAS: report folder found at {root}/{NAS_REPORT_FOLDER}, but no {t0s} folder yet")
            return path
    log(f"MFNAS: '{NAS_REPORT_FOLDER}' not at the expected paths - searching the shares ...")
    queue = [(s, 1) for s in nas.list_shares()]
    seen = 0
    while queue and seen < 300:
        path, depth = queue.pop(0)
        seen += 1
        for item in nas.list_folder(path, quiet=True) or []:
            if item["name"] == NAS_REPORT_FOLDER:
                log(f"MFNAS: report folder found by search: {item['path']}")
                return f"{item['path']}/{t0s}"
            if item["isdir"] and depth < 3:
                queue.append((item["path"], depth + 1))
    return None


def download_report(t0, final_path, poll_attempts, poll_interval):
    t0s = t0.strftime("%Y%m%d")
    os.makedirs(os.path.dirname(final_path) or ".", exist_ok=True)
    nas = NASClient()
    nas.login()
    try:
        folder, pick = None, None
        for attempt in range(1, poll_attempts + 1):
            folder = folder or find_date_folder(nas, t0s)
            files = [
                f for f in (nas.list_folder(folder, quiet=True) if folder else None) or []
                if not f["isdir"] and f["name"].lower().endswith(".csv")
            ]
            if files:
                files.sort(key=lambda f: f["name"])
                pick = files[-1]
                if len(files) > 1:
                    log(f"MFNAS: WARNING {len(files)} csv files in the folder, using the newest: {pick['name']}")
                break
            log(f"MFNAS: no csv in {folder or NAS_REPORT_FOLDER + '/' + t0s} yet (attempt {attempt}/{poll_attempts})")
            if attempt < poll_attempts:
                time.sleep(poll_interval)
        if not pick:
            shares = nas.list_shares()
            raise RuntimeError(
                f"MFNAS: no csv found for {t0s}. Visible shares: {shares}. "
                "Set ROLLING_NAS_ROOT if the report folder sits somewhere else."
            )
        log(f"MFNAS: {pick['name']}  ->  {os.path.basename(final_path)}")
        tmp = final_path + ".part"
        size = nas.download(pick["path"], tmp)
        if size < 1024:
            raise RuntimeError(f"MFNAS: downloaded file is only {size} bytes")
        os.replace(tmp, final_path)
        print(f"REMOTE={pick['name']}")
        print(f"LOCAL={final_path}")
        print(f"BYTES={size}")
        return final_path, pick["name"], size
    finally:
        nas.logout()


def main():
    ap = argparse.ArgumentParser(description="Download the MFNAS Order Date report")
    ap.add_argument("--date", help="T-0 = the run date YYYY-MM-DD (default today). The report holds the orders of the day before it.")
    ap.add_argument("--poll-attempts", type=int, default=int(os.environ.get("ROLLING_NAS_POLL_ATTEMPTS", "6")))
    ap.add_argument("--poll-interval", type=int, default=int(os.environ.get("ROLLING_NAS_POLL_INTERVAL", "300")))
    args = ap.parse_args()

    t0 = datetime.strptime(args.date, "%Y-%m-%d") if args.date else datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    order_date = t0 - timedelta(days=1)
    dest = report_path(order_date)
    log(f"T-0={t0:%Y-%m-%d}  ->  Order Date (T-1)={order_date:%Y-%m-%d}")
    download_report(t0, dest, args.poll_attempts, args.poll_interval)
    log("download-only: done")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log(f"ERROR: {ex}")
        sys.exit(1)
