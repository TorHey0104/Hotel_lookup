import html
import os
import re
import glob
import tempfile
from typing import Dict, Any, List
import importlib.util
import pythoncom

# Basic renderer used across modes
def render_with_signature(
    body_text: str,
    signature_entry: dict,
    body_is_html: bool = False,
    forward_html: str = "",
    forward_is_html: bool = False,
    base_style: str = "font-family:'Aptos',sans-serif; font-size:12pt; line-height:1.4;",
) -> dict:
    sig_html = signature_entry.get("html", "") if isinstance(signature_entry, dict) else ""
    sig_txt = signature_entry.get("text", "") if isinstance(signature_entry, dict) else ""

    link_pattern = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")

    def linkify_text(txt: str) -> str:
        parts = []
        idx = 0
        combined_pattern = re.compile(r"\[([^\]]+)\]\(([^)]+)\)|<a\s+[^>]*href=[\"']?([^>\"']+)[\"']?[^>]*>(.*?)</a>", re.IGNORECASE | re.DOTALL)
        for m in combined_pattern.finditer(txt):
            pre = txt[idx:m.start()]
            if pre:
                parts.append(html.escape(pre).replace("\n", "<br>"))
            label = ""
            url_raw = ""
            if m.group(1) is not None:
                label = m.group(1)
                url_raw = m.group(2)
            else:
                label = m.group(4)
                url_raw = m.group(3)
            url_raw = (url_raw or "").strip()
            if not url_raw.lower().startswith(("http://", "https://")):
                url_raw = "https://" + url_raw
            url = html.escape(url_raw)
            label_safe = html.escape(label or url_raw)
            parts.append(f'<a href="{url}">{label_safe}</a>')
            idx = m.end()
        tail = txt[idx:]
        if tail:
            parts.append(html.escape(tail).replace("\n", "<br>"))
        return "".join(parts) if parts else html.escape(txt).replace("\n", "<br>")

    def to_html(txt: str, allow_links: bool = False) -> str:
        return linkify_text(txt) if allow_links else html.escape(txt).replace("\n", "<br>")

    def looks_like_html(txt: str) -> bool:
        lowered = txt.lower()
        return any(tag in lowered for tag in ("<html", "<body", "<table", "<div", "<p", "<a "))

    def wrap_block(content: str) -> str:
        return f"<div style='white-space:pre-wrap; {base_style}'>{content}</div>"

    def extract_body_fragment(html_text: str) -> str:
        lowered = html_text.lower()
        body_start = lowered.find("<body")
        if body_start == -1:
            return html_text
        start = lowered.find(">", body_start)
        if start == -1:
            return html_text
        body_end = lowered.rfind("</body>")
        if body_end == -1:
            return html_text[start + 1 :]
        return html_text[start + 1 : body_end]

    body_has_links = bool(link_pattern.search(body_text)) if not body_is_html else False
    user_block = body_text if body_is_html else wrap_block(to_html(body_text, allow_links=True))

    sig_block = ""
    if sig_html:
        sig_block = wrap_block(sig_html if looks_like_html(sig_html) else to_html(sig_html, allow_links=True))
    elif sig_txt:
        sig_block = wrap_block(to_html(sig_txt, allow_links=True))

    forward_block = ""
    if forward_html:
        if forward_is_html or looks_like_html(forward_html):
            forward_block = extract_body_fragment(forward_html)
        else:
            forward_block = wrap_block(to_html(forward_html, allow_links=True))

    if forward_block or sig_block or body_is_html or forward_is_html or sig_html or body_has_links:
        html_parts = [user_block]
        if sig_block:
            html_parts.append(sig_block)
        if forward_block:
            html_parts.append(forward_block)
        html_body = "<br><br>".join([p for p in html_parts if p])
        if not forward_html and "<html" not in html_body.lower():
            html_body = f"<!DOCTYPE html><html><body style=\"{base_style}\">{html_body}</body></html>"
        return {"html": html_body}

    combined = body_text
    if sig_txt:
        combined += "\n\n" + sig_txt
    if forward_html:
        combined += "\n\n" + forward_html
    return {"text": combined}


def get_outlook_app(force_refresh: bool = False):
    """Return a cached Outlook Application COM object (or create it on first use)."""
    global _outlook_app

    if "_outlook_app" not in globals():
        globals()["_outlook_app"] = None
    _cached = globals()["_outlook_app"]
    if _cached is not None and not force_refresh:
        return _cached

    import win32com.client  # type: ignore[import-untyped]

    try:
        globals()["_outlook_app"] = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        try:
            gen_path = win32com.client.gencache.GetGeneratePath()
            if os.path.isdir(gen_path):
                import shutil
                shutil.rmtree(gen_path, ignore_errors=True)
        except Exception:
            pass
        globals()["_outlook_app"] = win32com.client.Dispatch("Outlook.Application")
    return globals()["_outlook_app"]


def save_forward(item) -> Dict[str, Any]:
    """Capture HTML/attachments from an Outlook MailItem-like object."""
    forward_template = {"subject": "", "body_text": "", "attachments": [], "temp_dir": "", "is_html": False}
    forward_template["subject"] = f"FW: {getattr(item, 'Subject', '')}"
    html_body = getattr(item, "HTMLBody", "") or ""
    plain_body = getattr(item, "Body", "") or ""
    forward_template["body_text"] = html_body if html_body else plain_body
    forward_template["is_html"] = bool(html_body)
    temp_dir = tempfile.mkdtemp(prefix="forward_src_")
    forward_template["temp_dir"] = temp_dir
    atts = getattr(item, "Attachments", None)
    if atts:
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            try:
                save_path = os.path.join(temp_dir, att.FileName)
                att.SaveAsFile(save_path)
                forward_template["attachments"].append(save_path)
            except Exception:
                pass
    return forward_template
