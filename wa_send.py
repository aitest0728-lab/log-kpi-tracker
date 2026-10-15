#!/usr/bin/env python3
"""
wa_send.py  (KPI Dashboard v9.1)

Posts ONE image + caption to a WhatsApp group through WhatsApp Web, using the persistent, already-logged-in Chromium
profile (WA_SESSION_DIR) - the same profile / group / hide-window logic kpi_pipeline.py's cost-report download uses.

No import of kpi_pipeline (it insists on Tableau credentials at import time), so it also works on its own:

    python3 wa_send.py --image logs/delay_report/delay_20261007.png --caption-file logs/delay_report/delay_20261007.txt
    python3 wa_send.py --image x.png --caption "hello" --group "My test group"     # try it on a private test group first

The browser is opened ONLY by wa_common.wa_session() (one profile, one lock, one fingerprint). The session_dir / hide_mode /
headless / user_agent / virtual_screen arguments below are kept so old callers still work, but they are ignored - set
WA_SESSION_DIR / WA_HIDE_MODE / ... in the environment (see wa_common.py). Own environment: WA_TARGET_GROUP.
"""
import argparse
import contextlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

try:
    import fcntl                      # Linux / WSL / macOS
except ImportError:                   # Windows: no cross-process lock, the jobs are scheduled far apart anyway
    fcntl = None

DEFAULT_SESSION_DIR = os.environ.get("WA_SESSION_DIR", "/home/chipanl/whatsapp_session_2")
DEFAULT_GROUP = os.environ.get("WA_TARGET_GROUP", "LOG 區頭 x Head office")
def _default_hide_mode():
    m = os.environ.get("WA_HIDE_MODE", "auto").strip().lower()
    if os.environ.get("WA_VIRTUAL_DISPLAY") == "0":
        m = "off"
    if m == "auto":
        m = "xvfb" if sys.platform.startswith("linux") else "headless"
    return m


# The WhatsApp session (lock, hidden display, browser flags, clean close) is wa_common's job - never launch Chrome here.
sys.path.insert(0, os.environ.get("WA_COMMON_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session"))
import wa_common


# -------------------------------------------------------------------------------------------------- page helpers
def dismiss_wa_popups(page, escape=True):
    """One-time WhatsApp Web promo tooltips/dialogs float over the UI and make Playwright refuse to click.
    escape=False once a chat is open: Escape would CLOSE that chat (and with it the attach button)."""
    for _ in range(4):
        if escape:
            page.keyboard.press("Escape")
            time.sleep(0.3)
        clicked = False
        for sel in ['[role="tooltip"] [aria-label*="Close" i]', '[role="tooltip"] [aria-label*="Dismiss" i]',
                    '[role="dialog"] [aria-label*="Close" i]', '[role="dialog"] button:has-text("OK")',
                    '[role="dialog"] button:has-text("Got it")', 'span[data-icon="x-alt"]']:
            try:
                loc = page.locator(sel).first
                if loc.is_visible():
                    loc.click(timeout=1500)
                    clicked = True
                    time.sleep(0.5)
                    break
            except Exception:
                continue
        if not clicked:
            break


COMPOSER_SELECTORS = ['#main footer div[contenteditable="true"]', '#main footer [role="textbox"]']


def chat_is_writable(page, timeout_ms=6000):
    """True when the open chat has a message box. A group the account has left / been removed from (or a stale copy of it)
    shows 'You can't send messages to this group because you're no longer a member' and has NO message box."""
    return _first_visible(page, COMPOSER_SELECTORS, timeout_ms) is not None


def _try_matching_chats(page, title_sel, group):
    """Open every chat-list entry titled `group` in turn (there can be several: a removed/old copy + the live one) and
    stop at the first one we can actually post to. Returns True if such a chat is now open."""
    n = page.locator(title_sel).count()
    for i in range(n):
        try:
            page.locator(title_sel).nth(i).click(timeout=6000)
            page.wait_for_selector("#main", timeout=30000)
        except Exception:
            continue
        time.sleep(2)
        ok = chat_is_writable(page)
        print(f"  chat {i + 1}/{n} titled {group!r}: " + ("can send messages  <-- using this one" if ok else
              "READ-ONLY (account is not a member of this copy)"))
        if ok:
            return True
    return False


def open_group(page, group, shot):
    """Open the group we can POST to. 1) chat list  2) search box. If several chats share the title, each is tried."""
    title_sel = "#pane-side span[title=" + json.dumps(group, ensure_ascii=False) + "]"
    if _try_matching_chats(page, title_sel, group):
        return
    dismiss_wa_popups(page)
    box = None
    for sel in ['div[contenteditable="true"][data-tab="3"]', '[aria-label="Search input textbox"]',
                '[aria-label="Search or start a new chat"]', '#side div[role="textbox"]']:
        loc = page.locator(sel).first
        try:
            loc.wait_for(state="visible", timeout=4000)
            box = loc
            break
        except Exception:
            continue
    if box is None:
        shot(page, "search")
        raise RuntimeError("Could not find WhatsApp's search box.")
    try:
        box.click(timeout=5000)
    except Exception:
        dismiss_wa_popups(page)
        box.click(force=True, timeout=5000)
    box.fill(group)
    time.sleep(2.5)
    if _try_matching_chats(page, title_sel, group):
        return
    shot(page, "group")
    raise RuntimeError(
        f"The only chat(s) titled {group!r} in this WhatsApp Web profile are READ-ONLY ('you're no longer a member'). If new "
        f"messages from that group keep arriving here, the account IS still a member and this profile's cached group state is "
        f"stale (typical after the account was removed and re-added): re-link the profile - move {DEFAULT_SESSION_DIR!r} aside, "
        f"run  WA_HIDE_MODE=off python3 wa_send.py --list-groups  and scan the QR code (phone: Settings > Linked devices). "
        f"If no new messages arrive, re-add the account to the group first.")


def _first_visible(page, selectors, timeout_ms=4000):
    end = time.time() + timeout_ms / 1000
    while time.time() < end:
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() and loc.is_visible():
                    return loc
            except Exception:
                continue
        time.sleep(0.25)
    return None


ATTACH_SELECTORS = ['#main footer [aria-label="Attach"]', '#main footer span[data-icon="plus-rounded"]',
                    '#main footer span[data-icon="plus"]', '#main footer span[data-icon="clip"]',
                    '#main footer button[title="Attach"]', '#main footer [title="Attach"]']
CAPTION_SELECTORS = ['div[contenteditable="true"][aria-label*="caption" i]',
                     '[data-testid="media-caption-input-container"] div[contenteditable="true"]',
                     'div[role="dialog"] div[contenteditable="true"]',
                     'div[contenteditable="true"][data-lexical-editor="true"]:not(#main footer *)']
SEND_SELECTORS = ['[aria-label="Send"][role="button"]', 'div[role="button"][aria-label="Send"]',
                  'button[aria-label="Send"]', 'span[data-icon="send"]', 'span[data-icon="wds-ic-send-filled"]']


def attach_image(page, image_path, shot):
    """Attach button -> 'Photos & videos' -> file chooser. Falls back to feeding the hidden <input type=file> directly."""
    attach = _first_visible(page, ATTACH_SELECTORS, 8000)
    if attach is None:
        shot(page, "attach")
        raise RuntimeError("Could not find WhatsApp's attach (+) button.")
    attach.click(timeout=5000)
    time.sleep(1)
    try:
        import re
        with page.expect_file_chooser(timeout=6000) as fc:
            page.get_by_text(re.compile(r"Photos\s*(&|and)\s*videos", re.I)).first.click(timeout=4000)
        fc.value.set_files(image_path)
        return
    except Exception:
        pass
    try:                               # fallback: hidden image input that the attach menu creates
        page.locator('input[type="file"][accept*="image"]').first.set_input_files(image_path, timeout=6000)
        return
    except Exception:
        shot(page, "attach_file")
        raise RuntimeError("Attached nothing: neither the 'Photos & videos' chooser nor the hidden file input worked.")


def type_caption(page, caption, shot):
    box = _first_visible(page, CAPTION_SELECTORS, 30000)     # preview of the picture has to render first
    if box is None:
        shot(page, "caption")
        raise RuntimeError("The picture preview / caption box never appeared.")
    box.click(timeout=5000)
    lines = caption.split("\n")
    for i, line in enumerate(lines):
        if line:
            page.keyboard.insert_text(line)                  # insert_text handles emoji + 區/頭 without key mapping
        if i < len(lines) - 1:
            page.keyboard.press("Shift+Enter")               # plain Enter would SEND the message
    time.sleep(0.5)


def click_send(page, shot):
    btn = _first_visible(page, SEND_SELECTORS, 8000)
    try:
        if btn is None:
            raise RuntimeError("no send button")
        btn.click(timeout=5000)
    except Exception:
        page.keyboard.press("Enter")                         # last resort from inside the caption box


def wait_until_delivered(page, shot, max_wait_s=120):
    """Closing the browser while the upload is still going would drop the picture: wait for the preview to close and the
    outgoing message's 'pending' clock icon to disappear. Returns True if confirmed."""
    time.sleep(3)
    end = time.time() + max_wait_s
    while time.time() < end:
        preview_open = False
        try:
            preview_open = bool(page.locator(", ".join(CAPTION_SELECTORS)).count()) and \
                _first_visible(page, CAPTION_SELECTORS, 300) is not None
        except Exception:
            pass
        pending = 0
        try:
            pending = page.locator('#main span[data-icon="msg-time"]').count()
        except Exception:
            pass
        if not preview_open and pending == 0:
            time.sleep(3)
            return True
        time.sleep(1.5)
    shot(page, "unconfirmed")
    return False


# ------------------------------------------------------------------------------------------------------ main API
@contextlib.contextmanager
def wa_page(session_dir, hide_mode, headless, user_agent, virtual_screen, shot):
    """Shared wa_common session; yields the WhatsApp Web page once the chat list is loaded.
    (all arguments except `shot` are ignored - see the module docstring)"""
    with wa_common.wa_session() as (ctx, page):
        try:
            dismiss_wa_popups(page)
            yield page
        except Exception:
            shot(page, "error")
            raise


def _make_shot(debug_dir):
    Path(debug_dir).mkdir(parents=True, exist_ok=True)

    def shot(page, tag):
        try:
            p = os.path.join(debug_dir, f"_debug_whatsapp_{tag}.png")
            page.screenshot(path=p, full_page=True)
            print(f"  📸 screenshot: {p}")
        except Exception:
            pass
    return shot


def send_image_with_caption(image_path, caption, group=None, session_dir=None, hide_mode=None, headless=False,
                            user_agent=None, virtual_screen="1920x1080x24", debug_dir="./logs/delay_report"):
    """Open WhatsApp Web, open `group` (the copy we can post to), send `image_path` with `caption`. Raises on any failure
    (screenshots of the failing screen go to debug_dir/_debug_whatsapp_<tag>.png). Returns True when delivery was confirmed."""
    image_path = str(Path(image_path).resolve())
    if not os.path.exists(image_path):
        raise FileNotFoundError(image_path)
    group = group or DEFAULT_GROUP
    session_dir = session_dir or DEFAULT_SESSION_DIR
    hide_mode = hide_mode or _default_hide_mode()
    shot = _make_shot(debug_dir)
    print(f"🚀 Opening WhatsApp Web (session {session_dir!r}) to post to {group!r}...")
    with wa_page(session_dir, hide_mode, headless, user_agent, virtual_screen, shot) as page:
        open_group(page, group, shot)
        dismiss_wa_popups(page, escape=False)
        attach_image(page, image_path, shot)
        type_caption(page, caption, shot)
        click_send(page, shot)
        ok = wait_until_delivered(page, shot)
        print("  ✅ message sent." if ok else "  ⚠️ message was sent but delivery could not be confirmed (see screenshot).")
        return ok


def send_image_to_groups(image_path, caption, groups, debug_dir="./logs/delay_report"):
    """Post the same image + caption to several groups through ONE WhatsApp session (one Chrome start instead of one per group).
    Returns {group: None (delivered) | Exception}. Raises only if the session itself cannot be opened
    (SessionExpired / FingerprintChanged / RuntimeError) - then nothing was sent."""
    image_path = str(Path(image_path).resolve())
    if not os.path.exists(image_path):
        raise FileNotFoundError(image_path)
    shot = _make_shot(debug_dir)
    results = {}
    with wa_page(None, None, False, None, None, shot) as page:
        for g in groups:
            print(f"🚀 Posting to {g!r}...")
            try:
                open_group(page, g, shot)
                dismiss_wa_popups(page, escape=False)
                attach_image(page, image_path, shot)
                type_caption(page, caption, shot)
                click_send(page, shot)
                ok = wait_until_delivered(page, shot)
                print("  ✅ message sent." if ok else "  ⚠️ sent, but delivery could not be confirmed (see screenshot).")
                results[g] = None
            except Exception as e:
                results[g] = e
                shot(page, f"group_{len(results)}")
                for _ in range(2):                           # leave a half-open picture preview before the next group
                    try:
                        page.keyboard.press("Escape")
                        time.sleep(0.5)
                    except Exception:
                        pass
    return results


def list_matching_chats(group=None, session_dir=None, hide_mode=None, headless=False, user_agent=None,
                        virtual_screen="1920x1080x24", debug_dir="./logs/delay_report"):
    """Diagnostic (sends nothing): print every chat titled `group` that this session can see and whether it can post."""
    group = group or DEFAULT_GROUP
    session_dir = session_dir or DEFAULT_SESSION_DIR
    hide_mode = hide_mode or _default_hide_mode()
    shot = _make_shot(debug_dir)
    title_sel = "#pane-side span[title=" + json.dumps(group, ensure_ascii=False) + "]"
    with wa_page(session_dir, hide_mode, headless, user_agent, virtual_screen, shot) as page:
        n = page.locator(title_sel).count()
        print(f"Chat list: {n} chat(s) titled {group!r}")
        for i in range(n):
            page.locator(title_sel).nth(i).click(timeout=6000)
            page.wait_for_selector("#main", timeout=30000)
            time.sleep(2)
            subtitle = ""
            try:
                subtitle = page.locator("#main header").inner_text(timeout=2000).replace("\n", " | ")[:160]
            except Exception:
                pass
            ok = chat_is_writable(page)
            print(f"  #{i + 1}: {'CAN SEND' if ok else 'READ-ONLY'}  header: {subtitle}")
            shot(page, f"list_{i + 1}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Post an image + caption to a WhatsApp group via WhatsApp Web.")
    ap.add_argument("--image")
    ap.add_argument("--list-groups", action="store_true", help="diagnostic: list every chat with the group's title and "
                                                               "whether this session can post to it (sends nothing)")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--caption")
    g.add_argument("--caption-file")
    ap.add_argument("--group", default=None, help=f"default: WA_TARGET_GROUP ({DEFAULT_GROUP!r})")
    a = ap.parse_args()
    if a.list_groups:
        list_matching_chats(group=a.group)
        sys.exit(0)
    if not a.image or (a.caption is None and not a.caption_file):
        ap.error("--image and --caption/--caption-file are required (unless --list-groups)")
    cap = a.caption if a.caption is not None else Path(a.caption_file).read_text(encoding="utf-8")
    send_image_with_caption(a.image, cap, group=a.group)
