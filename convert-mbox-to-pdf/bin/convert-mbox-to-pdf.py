#!/usr/bin/env python3
"""
mbox → PDF. Spec: Perl converter + RobustMbox I/O.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import os
import json
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.header import decode_header
from email.message import Message
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, TextIO, Tuple, Set

from PIL import Image as PILImage
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.pagesizes import A0, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

from robust_mbox import RobustMbox, write_mbox

PILImage.MAX_IMAGE_PIXELS = 800_000_000

CMD_TIMEOUT = 24 * 60 * 60
QPDF_TIMEOUT = CMD_TIMEOUT
PDF_VERSION = (1, 7)


def _raise_pypdf_limits() -> None:
    cap = 512 * 1024 * 1024
    for modname, attr in (
            ("pypdf.filters", "MAX_DECLARED_STREAM_LENGTH"),
            ("pypdf.generic._data_structures", "MAX_DECLARED_STREAM_LENGTH"),
            ("pypdf._utils", "MAX_DECLARED_STREAM_LENGTH"),
    ):
        try:
            mod = __import__(modname, fromlist=[attr])
            if hasattr(mod, attr):
                setattr(mod, attr, cap)
        except Exception:
            pass
    try:
        from pypdf import apply_configuration  # type: ignore
        apply_configuration(maximum_declared_stream_length=cap)
    except Exception:
        pass


_raise_pypdf_limits()

DIN_A = {
    "A0": A0, "A1": A1, "A2": A2, "A3": A3, "A4": A4,
    "A5": A5, "A6": A6, "A7": A7, "A8": A8, "A9": A9, "A10": A10,
}
PERL_REF_W = 210.0 * 300.0 / 25.4
DARK_BLUE = HexColor("#1a365d")
LIGHT_GRAY = HexColor("#e2e8f0")
MUTED_GREY = "#718096"
DEDUP_POSTFIX = "-deduplicated.mbox"
SKIPPED_POSTFIX = "-skipped.mbox"
RATIO_COLON = "\u2236"
BAR_OPTICAL_NUDGE = 0
TITLE_BAR_TEXT_NUDGE = 3.0      # image / montage / PDF title text; + = lower
ATTACH_HEADING_RULE_NUDGE = 6.0  # + = rule lower, heading unchanged
SPLIT_LIMIT_BYTES = 8 * 1024 * 1024 * 1024

CONVERTIBLE = {
    "doc", "docx", "odt", "rtf", "xls", "xlsx", "ods", "csv",
    "ppt", "pptx", "odp", "html", "htm", "epub",
}
IMAGE_EXT = {
    "jpeg": "jpg", "jpg": "jpg", "png": "png", "gif": "gif",
    "bmp": "bmp", "webp": "webp", "heic": "heic", "heif": "heic",
    "tiff": "tiff", "tif": "tiff", "avif": "avif",
}
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".heic", ".heif", ".tif", ".tiff", ".avif"}
EMOJI_MAP = {
    "✝": "†", "😊": ":)", "😉": ";)", "😂": ":D", "😄": ":D", "😃": ":D",
    "🙂": ":)", "😔": ":(", "😢": ":(", "😭": ":'(", "😡": ">:(",
    "😠": ">:(", "🤔": "hmm", "🙄": "roll eyes", "😮": ":O", "😲": ":O",
    "😱": "!!!", "🥰": "<3", "😍": "<3", "❤️": "<3", "💕": "<3", "💖": "<3",
    "👍": "+1", "👎": "-1", "👏": "clap", "🙏": "please",
    "🔥": "fire", "💯": "100", "✅": "[ok]", "❌": "[x]", "⭐": "*", "🌟": "*",
}

debug = False
quiet = False
verbose = False
fast_level = 0
finalize = False
prefer_plain_text = True
email_to_name: Dict[str, str] = {}
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
FONT_ITALIC = "Helvetica-Oblique"
log_ctx: Dict[str, Any] = {"id": None, "date": "", "subject": "", "from": ""}
log_convert: Optional[TextIO] = None
log_errors: Optional[TextIO] = None


@dataclass
class Attachment:
    original_name: str
    saved_path: Path
    size: int
    type: str
    converted_from: Optional[str] = None
    sha256: Optional[str] = None
    converted_size: Optional[int] = None
    converted_path: Optional[Path] = None


@dataclass
class Layout:
    width: float
    height: float
    image_dpi: int
    margin: float
    line_height: float
    font_body: float
    font_header: float
    font_subject: float
    font_bar: float
    font_attach: float
    font_empty: float
    leading: float
    label_x: float
    name_x: float
    email_x: float
    quote_indent: float
    bar_width: float
    header_rule: float
    attach_bar_h: float


@dataclass
class ProcessedEmail:
    count: int
    date_str: str
    from_str: str
    to_str: str
    cc_str: str
    subject: str
    body_text: str
    original_plain: str
    original_html: str
    date_raw: str = ""
    date_dt: Optional[datetime] = None
    images: List[Attachment] = field(default_factory=list)
    pdf_attachments: List[Attachment] = field(default_factory=list)
    all_attachments: List[Attachment] = field(default_factory=list)


def q(s: Any) -> str:
    return "'" + str(s).replace("'", "\\'") + "'"


def _ctx_prefix() -> str:
    if log_ctx.get("id") is None:
        return ""
    return f"[Email #{log_ctx['id']} | {log_ctx.get('date') or ''} | {(log_ctx.get('subject') or '')[:60]}] "


def _write_log_files(line: str, review: bool) -> None:
    if log_convert is not None:
        try:
            log_convert.write(line + "\n")
            log_convert.flush()
        except Exception:
            pass
    if review and log_errors is not None:
        try:
            log_errors.write(line + "\n")
            log_errors.flush()
        except Exception:
            pass


def log(level: str, message: str, review: bool = False) -> None:
    level = level.upper()
    message = _ctx_prefix() + message
    if level == "ERROR":
        review = True
        line = f"ERROR: {message}"
        print(line, file=sys.stderr)
        _write_log_files(line, True)
        return
    if level == "WARNING":
        line = f"WARNING: {message}"
        print(line)
        _write_log_files(line, review)
        return
    if level == "DEBUG":
        line = f"DEBUG: {message}"
        if debug:
            print(line)
        _write_log_files(line, False)
        return
    if level == "VERBOSE":
        line = f"VERBOSE: {message}"
        if verbose or debug:
            print(line)
        _write_log_files(line, False)
        return
    line = message
    if not quiet:
        print(line)
    _write_log_files(("INFO: " + message) if not message.startswith("INFO:") else message, False)


def set_log_ctx(count: Optional[int], date: str = "", subject: str = "", frm: str = "") -> None:
    log_ctx["id"] = count
    log_ctx["date"] = date
    log_ctx["subject"] = subject
    log_ctx["from"] = frm


def parse_email_datetime(msg: Message) -> Tuple[str, Optional[datetime]]:
    raw = (msg.get("Date") or "").strip()
    try:
        return raw, parsedate_to_datetime(raw)
    except Exception:
        return raw, None


def archive_when_label(raw: str, dt: Optional[datetime]) -> str:
    raw = (raw or "").strip()
    if not raw or dt is None:
        return "unknown date"
    raw_l = raw.lower()
    has_year = bool(re.search(r"\b(?:19|20)\d{2}\b", raw))
    has_month = bool(
        re.search(
            r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b",
            raw_l,
        )
        or re.search(r"\b(0?[1-9]|1[0-2])[./-]", raw)
    )
    has_time = bool(re.search(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", raw))
    has_sec = bool(re.search(r"\b\d{1,2}:\d{2}:\d{2}\b", raw))
    has_day = has_month and bool(re.search(r"\b(0?[1-9]|[12]\d|3[01])\b", raw))
    if not has_year:
        return "unknown date"
    if not has_month:
        return f"{dt.year}"
    if not has_day:
        return f"{dt.month:02d}.{dt.year}"
    day = f"{dt.day:02d}.{dt.month:02d}.{dt.year}"
    if not has_time:
        return day
    hm = f"{dt.hour:02d}{RATIO_COLON}{dt.minute:02d}"
    if not has_sec:
        return f"{day} {hm}"
    return f"{day} {hm}{RATIO_COLON}{dt.second:02d}"


def _parse_stamp(text: str) -> Optional[datetime]:
    text = (text or "").strip()
    if not text:
        return None
    text = text.replace("∶", ":")
    for fmt in (
            "%Y:%m:%d %H:%M:%S",
            "%Y-%m-%d %H:%M:%S",
            "%Y:%m:%d %H:%M",
            "%Y-%m-%d %H:%M",
            "%Y:%m:%d",
            "%Y-%m-%d",
            "%Y",
    ):
        try:
            sample = text[:19] if len(text) >= 19 and ":" in text[10:12] else text
            return datetime.strptime(sample, fmt)
        except Exception:
            continue
    try:
        return parsedate_to_datetime(text)
    except Exception:
        return None


def attachment_internal_mtime(path: Path) -> Optional[datetime]:
    head = read_head(path, 8)
    suffix = path.suffix.lower()
    looks_image = suffix in IMAGE_SUFFIXES or head.startswith(
        (b"\xff\xd8\xff", b"\x89PNG", b"GIF8", b"BM")
    )
    looks_pdf = suffix == ".pdf" or head.startswith(b"%PDF-")
    if looks_image:
        rc, out = run_cmd([
            magick_bin(),
            f"{path}[0]" if suffix == ".gif" else str(path),
            "-format",
            "%[EXIF:DateTimeOriginal]\n%[EXIF:DateTimeDigitized]\n%[EXIF:DateTime]",
            "info:",
        ])
        if rc == 0:
            for line in out.splitlines():
                dt = _parse_stamp(line.strip())
                if dt:
                    return dt
        return None
    if looks_pdf:
        try:
            meta = PdfReader(str(path)).metadata
            for attr in ("modification_date", "creation_date"):
                val = getattr(meta, attr, None) if meta else None
                if isinstance(val, datetime):
                    return val.replace(tzinfo=None) if val.tzinfo else val
                dt = _parse_stamp(str(val) if val else "")
                if dt:
                    return dt
        except Exception:
            return None
    return None


def apply_archive_mtime(dest: Path, email: ProcessedEmail) -> None:
    dt = attachment_internal_mtime(dest) or email.date_dt
    if dt is None:
        return
    if dt.tzinfo is not None:
        dt = dt.replace(tzinfo=None)
    try:
        ts = dt.timestamp()
    except Exception:
        return
    os.utime(dest, (ts, ts))


def prepare_archive_root(rundir: Path) -> Path:
    root = rundir / "attachments"
    root.mkdir(parents=True, exist_ok=True)
    return root


def archive_email(email: ProcessedEmail, archive_root: Path) -> None:
    label = archive_when_label(email.date_raw, email.date_dt)
    dest_dir = archive_root / f"Email #{email.count} ({label})"
    dest_dir.mkdir(parents=True, exist_ok=True)
    if email.body_text:
        p = dest_dir / "03_converted.txt"
        p.write_text(email.body_text, encoding="utf-8")
        apply_archive_mtime(p, email)
    if email.original_plain:
        p = dest_dir / "01_original.txt"
        p.write_text(email.original_plain, encoding="utf-8")
        apply_archive_mtime(p, email)
    if email.original_html:
        p = dest_dir / "02_original.html"
        p.write_text(email.original_html, encoding="utf-8")
        apply_archive_mtime(p, email)
    for i, att in enumerate(email.all_attachments, 1):
        if not att.saved_path.exists():
            continue
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", att.original_name)
        dest = dest_dir / f"{i}. {safe}"
        try:
            shutil.copy2(att.saved_path, dest)
            apply_archive_mtime(dest, email)
        except Exception as e:
            log("ERROR", f"archive copy {q(att.original_name)}: {e}")


def mtime_stamp(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d-%m-%Y-%H-%M-%S")


def run_id_now() -> str:
    return datetime.now().strftime("%d-%m-%Y-%H-%M-%S")


def make_rundir(output_path: Path) -> Path:
    stamp = run_id_now()
    parent = output_path / "run"
    parent.mkdir(parents=True, exist_ok=True)
    rundir = parent / f"run-{stamp}"
    i = 1
    while rundir.exists():
        rundir = parent / f"run-{stamp}_{i}"
        i += 1
    rundir.mkdir(parents=True, exist_ok=True)
    log("INFO", f"Run directory: {q(rundir)}")
    return rundir


def link_into_root(output_path: Path, rundir: Path, name: str) -> None:
    src = rundir / name
    dest = output_path / name
    if not src.exists():
        return
    if dest.is_symlink():
        dest.unlink()
    elif dest.exists():
        log("WARNING", f"cannot symlink {q(dest)}: real file/dir exists, rundir keeps {q(src)}")
        return
    try:
        dest.symlink_to(src)
        log("VERBOSE", f"symlink {q(dest)} -> {q(src)}")
    except Exception as e:
        log("WARNING", f"cannot symlink {q(dest)}: {e}")


def cleanup_stale_part_symlinks(output_path: Path, stem: str, suffix: str, keep: Set[str]) -> None:
    for path in list(output_path.glob(f"{stem}-part*{suffix}")) + list(
            output_path.glob(f"{stem}-part*-master{suffix}")
    ):
        if path.name in keep:
            continue
        if path.is_symlink():
            path.unlink()
            log("INFO", f"removed stale part symlink {q(path)}")

def setup_log_files(log_dir: Path) -> None:
    global log_convert, log_errors
    log_dir.mkdir(parents=True, exist_ok=True)
    log_convert = open(log_dir / "convert.log", "w", encoding="utf-8")
    log_errors = open(log_dir / "errors.log", "w", encoding="utf-8")


def close_log_files() -> None:
    global log_convert, log_errors
    for fh in (log_convert, log_errors):
        if fh is not None:
            try:
                fh.close()
            except Exception:
                pass
    log_convert = None
    log_errors = None


def require_tool(names: List[str], what: str) -> str:
    for name in names:
        path = shutil.which(name)
        if path:
            log("VERBOSE", f"Found {what}: {q(path)}")
            return path
    log("ERROR", f"Required tool for {what} not found (tried: {', '.join(names)})")
    sys.exit(1)


def _pid_stopped(pid: int) -> bool:
    try:
        with open(f"/proc/{pid}/stat", "r", encoding="utf-8") as f:
            raw = f.read()
        state = raw[raw.rfind(")") + 1:].split()[0]
        return state == "T"
    except Exception:
        return False


def run_cmd(cmd: List[str], timeout: int = CMD_TIMEOUT) -> Tuple[int, str]:
    shown = " ".join(q(x) for x in cmd)
    log("DEBUG", f"exec: {shown}")
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception as e:
        log("ERROR", f"command failed: {e}")
        if debug:
            traceback.print_exc()
        return 1, str(e)
    remaining = float(timeout)
    slice_s = 0.5
    try:
        while True:
            try:
                stdout, stderr = p.communicate(timeout=slice_s)
                raw = (stdout or b"") + (stderr or b"")
                out = raw.decode("utf-8", errors="replace")
                if debug and out.strip():
                    log("DEBUG", f"exit {p.returncode} out={out.strip()[:800]}")
                else:
                    log("DEBUG", f"exit {p.returncode}")
                return p.returncode or 0, out
            except subprocess.TimeoutExpired:
                if p.poll() is not None:
                    stdout, stderr = p.communicate()
                    raw = (stdout or b"") + (stderr or b"")
                    return p.returncode or 0, raw.decode("utf-8", errors="replace")
                if _pid_stopped(p.pid):
                    continue
                remaining -= slice_s
                if remaining <= 0:
                    p.kill()
                    try:
                        p.communicate(timeout=10)
                    except Exception:
                        pass
                    log("ERROR", f"command timeout after {timeout}s: {shown}")
                    return 1, f"timeout after {timeout}s: {shown}"
    except Exception as e:
        log("ERROR", f"command failed: {e}")
        if debug:
            traceback.print_exc()
        try:
            p.kill()
        except Exception:
            pass
        return 1, str(e)


def file_sha(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def sha_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_head(path: Path, n: int = 16) -> bytes:
    try:
        with open(path, "rb") as f:
            return f.read(n)
    except Exception:
        return b""


def unix_file_mime(path: Path) -> str:
    tool = shutil.which("file")
    if not tool:
        return ""
    rc, out = run_cmd([tool, "--brief", "--mime-type", str(path)])
    return out.strip().split("\n")[0] if rc == 0 else ""


def cache_stem(src: Path) -> str:
    return src.stem if src.suffix.lower() == ".mbox" else src.name


def deduplicated_path(src: Path) -> Path:
    return src.with_name(cache_stem(src) + DEDUP_POSTFIX)


def skipped_path(src: Path) -> Path:
    return src.with_name(cache_stem(src) + SKIPPED_POSTFIX)


def new_pdf_writer() -> PdfWriter:
    writer = PdfWriter()
    try:
        writer._header = b"%PDF-1.7"
    except Exception:
        pass
    return writer


def new_canvas(path: str, page_size) -> rl_canvas.Canvas:
    try:
        return rl_canvas.Canvas(path, pagesize=page_size, pdfVersion=PDF_VERSION)
    except TypeError:
        return rl_canvas.Canvas(path, pagesize=page_size)


def looks_like_email_bytes(raw: bytes) -> bool:
    sample = raw.lstrip()[:2048]
    if sample.startswith(b"From "):
        return True
    head = sample.replace(b"\r\n", b"\n")
    return bool(re.search(br"(?im)^(from|received|subject|mime-version|date|to):", head))


def extract_nested_message(part: Message) -> Optional[Message]:
    inner = part.get_payload()
    if isinstance(inner, list):
        for item in inner:
            if isinstance(item, Message):
                return item
    raw = part.get_payload(decode=True)
    if not raw:
        return None
    if not looks_like_email_bytes(raw):
        return None
    try:
        return BytesParser(policy=policy.default).parsebytes(raw)
    except Exception as e:
        log("DEBUG", f"rfc822 bytes parse: {e}")
        return None


def resolve_input_mboxes_fast(input_files: List[str], use_dedup: bool) -> List[Path]:
    resolved: List[Path] = []
    for raw in input_files:
        src = Path(raw)
        log_mbox_family_counts(src)
        if src.name.endswith(DEDUP_POSTFIX) or src.name.endswith(SKIPPED_POSTFIX):
            log("ERROR", f"Refusing input that is a cache mbox: {q(src)}")
            sys.exit(1)
        if not src.exists():
            log("ERROR", f"mbox not found: {q(src)}")
            sys.exit(1)
        if use_dedup:
            dest = deduplicated_path(src)
            if dest.exists():
                log("INFO", f"list: using existing cache {q(dest)}")
                resolved.append(dest)
                continue
            log("INFO", f"list: no cache, using source {q(src)}")
        resolved.append(src)
    return resolved


def list_field(s: str) -> str:
    return re.sub(r"[\t\r\n]+", " ", (s or "")).strip()


def list_route(msg: Message) -> str:
    frm = clean_text(decode_header_value(msg, "From"))
    to = clean_text(decode_header_value(msg, "To"))
    frm = re.sub(r"<(.+?)>", r"\1", frm)
    frm = re.sub(r'^"(.+?)"\s*', r"\1 ", frm)
    to = re.sub(r'^"(.+?)"\s*', r"\1 ", to)
    to = re.sub(r"<(.+?)>", r"\1", to)
    return f"{list_field(frm)} → {list_field(to)}"


def list_stamp(msg: Message) -> str:
    try:
        return parsedate_to_datetime(msg.get("Date") or "").strftime("%d-%m-%Y-%H-%M-%S")
    except Exception:
        return "Unknown-date"


def list_attachment_names(msg: Message) -> str:
    names: List[str] = []
    seen: set[str] = set()
    for part in msg.walk():
        ctype = (part.get_content_type() or "").lower()
        if ctype.startswith("multipart/") or ctype in ("text/plain", "text/html"):
            continue
        raw = extract_filename(part) or ""
        name = raw.strip().strip("\"'")
        if ctype.startswith("message/rfc822"):
            name = name if name and name != "attachment.bin" else "attached-email"
        elif not name or name == "attachment.bin":
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        names.append(name)
    return ", ".join(names)


def print_email_list(
        messages: List[Message], year: int, selector: str, needle: str = "",
) -> None:
    try:
        clauses = parse_selector(selector)
    except ValueError as e:
        log("ERROR", str(e))
        sys.exit(1)
    needle = (needle or "").lower()
    n = len(messages)
    for i, msg in enumerate(messages, 1):
        if year != -1:
            try:
                if parsedate_to_datetime(msg.get("Date") or "").year != year:
                    continue
            except Exception:
                pass
        if not index_selected(i, n, clauses):
            continue
        subj = list_field(clean_text(decode_header_value(msg, "Subject", "(no subject)")))
        atts = list_field(list_attachment_names(msg))
        line = f"{i}\t{list_stamp(msg)}\t{list_route(msg)}\t{subj}\t{atts}"
        if needle and needle not in line.lower():
            continue
        print(line)


def declared_role(ext: str, mime: str) -> str:
    mime = (mime or "").lower()
    ext = (ext or "").lower()
    if mime == "application/pdf" or ext == "pdf":
        return "pdf"
    if mime.startswith("image/") or ext in IMAGE_EXT:
        return "image"
    if mime.startswith("video/"):
        return "video"
    if mime in ("text/vcard", "text/x-vcard", "text/directory") or ext == "vcf":
        return "vcf"
    if mime == "text/calendar" or ext == "ics":
        return "ics"
    if mime == "message/rfc822" or ext == "eml":
        return "eml"
    if (
            ext in CONVERTIBLE
            or mime.startswith("application/msword")
            or mime.startswith("application/vnd.ms-")
            or mime.startswith("application/vnd.openxmlformats")
            or mime.startswith("application/vnd.oasis.opendocument")
            or mime in ("application/rtf", "text/rtf", "application/vnd.ms-excel")
    ):
        return "office"
    if mime in ("application/zip", "application/x-zip-compressed") or ext == "zip":
        return "zip"
    return ""


def sniff_kind(path: Path, declared_ext: str = "", declared_mime: str = "") -> str:
    head = read_head(path, 16)
    mime = unix_file_mime(path)
    ext = (declared_ext or path.suffix.lstrip(".")).lower()

    def warn_mismatch(detected: str) -> str:
        role = declared_role(ext, declared_mime) or declared_role(ext, mime)
        if role and role != detected:
            log(
                "WARNING",
                f"magic role {q(detected)} != declared role {q(role)} "
                f"name={q(path.name)} mime={q(declared_mime or mime)}",
                review=True,
            )
        return detected

    if head.startswith(b"%PDF-") or mime == "application/pdf":
        return warn_mismatch("pdf")
    if head.startswith(b"\xff\xd8\xff") or mime in ("image/jpeg", "image/jpg"):
        return warn_mismatch("image")
    if head.startswith(b"\x89PNG") or mime == "image/png":
        return warn_mismatch("image")
    if head.startswith(b"GIF8") or mime == "image/gif":
        return warn_mismatch("image")
    if head.startswith(b"BM") and mime.startswith("image/"):
        return warn_mismatch("image")
    if head[:4] == b"RIFF" and b"WEBP" in read_head(path, 16):
        return warn_mismatch("image")
    if mime.startswith("image/"):
        return warn_mismatch("image")
    if mime.startswith("video/"):
        return warn_mismatch("video")
    if head.startswith(b"BEGIN:VCARD") or mime in ("text/vcard", "text/x-vcard"):
        return warn_mismatch("vcf")
    peek = read_head(path, 64).upper()
    if peek.startswith(b"BEGIN:VCARD"):
        return warn_mismatch("vcf")
    if peek.startswith(b"BEGIN:VCALENDAR") or mime == "text/calendar":
        return warn_mismatch("ics")
    if head.startswith(b"From ") or mime == "message/rfc822":
        return warn_mismatch("eml")
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return warn_mismatch("office")
    if head.startswith(b"PK\x03\x04") or mime in (
            "application/zip",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.oasis.opendocument.text",
            "application/vnd.oasis.opendocument.spreadsheet",
    ):
        inner = office_zip_kind(path)
        if inner:
            return warn_mismatch("office")
        if mime.startswith("application/vnd.") or ext in CONVERTIBLE:
            return warn_mismatch("office")
        return warn_mismatch("zip")

    if ext in IMAGE_EXT or declared_mime.startswith("image/"):
        log("INFO", f"type guessed from extension/name only: {q(path.name)} ext={q(ext)}")
        return "image"
    if ext == "pdf" or declared_mime == "application/pdf":
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "pdf"
    if ext in CONVERTIBLE:
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "office"
    if ext == "zip":
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "zip"
    if ext == "ics":
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "ics"
    if ext == "vcf":
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "vcf"
    if ext == "eml":
        log("INFO", f"type guessed from extension/name only: {q(path.name)}")
        return "eml"
    log("WARNING", f"unknown magic for {q(path.name)} head={head[:12]!r} mime={q(mime)}", review=True)
    return "other"


def office_zip_kind(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "mimetype" in names:
                mt = zf.read("mimetype").decode("utf-8", "replace")
                if "opendocument" in mt:
                    return True
            if "[Content_Types].xml" in names:
                return True
    except Exception:
        return False
    return False


def norm_body(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]
    text = "\n".join(lines).strip()
    return re.sub(r"\n{3,}", "\n\n", text)


def email_fingerprint(email: ProcessedEmail) -> Tuple[str, str]:
    body_h = hashlib.sha256(norm_body(email.body_text).encode("utf-8", "replace")).hexdigest()
    att_bits = []
    for att in sorted(email.all_attachments, key=lambda a: a.original_name.lower()):
        if att.saved_path.exists():
            digest = att.sha256 or file_sha(att.saved_path)
            att.sha256 = digest
            att_bits.append(f"{att.original_name}:{att.size}:{digest}")
    att_h = hashlib.sha256("\n".join(att_bits).encode()).hexdigest()
    return body_h, att_h


def is_empty_email(email: ProcessedEmail) -> bool:
    return not norm_body(email.body_text) and not email.all_attachments


def warn_empties(email: ProcessedEmail) -> None:
    body_empty = not norm_body(email.body_text)
    subj_empty = not (email.subject or "").strip() or email.subject.strip() == "(no subject)"
    natt = len(email.all_attachments)
    if body_empty and natt == 0:
        log("WARNING", "empty body and no attachments", review=True)
    if subj_empty and body_empty:
        log("WARNING", f"empty subject and empty body (attachments: {natt})", review=True)


def filter_duplicate_emails(emails: List[ProcessedEmail], action: str) -> List[ProcessedEmail]:
    seen_strong: Dict[Tuple[str, str], int] = {}
    seen_body: Dict[str, int] = {}
    seen_att: Dict[str, int] = {}
    kept: List[ProcessedEmail] = []
    empty_body = hashlib.sha256(b"").hexdigest()
    for email in emails:
        set_log_ctx(email.count, email.date_str, email.subject, email.from_str)
        warn_empties(email)
        body_h, att_h = email_fingerprint(email)
        strong = (body_h, att_h)
        empty = is_empty_email(email)
        first_strong = seen_strong.get(strong)
        first_body = seen_body.get(body_h)
        first_att = seen_att.get(att_h) if email.all_attachments else None
        log("DEBUG", f"fp body={body_h[:12]} att={att_h[:12]} empty={empty}")
        if not empty and first_strong is not None:
            log("WARNING", f"strong duplicate of email #{first_strong} (body+attachments)", review=True)
            if action == "skip":
                continue
        elif not empty and first_body is not None and body_h != empty_body:
            log("INFO", f"likely duplicate of #{first_body} (same body, different attachments)")
        elif first_att is not None:
            log("INFO", f"likely duplicate of #{first_att} (same attachments, different body)")
        if not empty and strong not in seen_strong:
            seen_strong[strong] = email.count
        if body_h not in seen_body:
            seen_body[body_h] = email.count
        if email.all_attachments and att_h not in seen_att:
            seen_att[att_h] = email.count
        kept.append(email)
    set_log_ctx(None)
    return kept


def make_layout(page_size: Tuple[float, float], image_dpi: int) -> Layout:
    w, h = page_size
    s = w / PERL_REF_W
    return Layout(
        width=w, height=h, image_dpi=image_dpi,
        margin=50.0 * s, line_height=50.0 * s,
        font_body=32.0 * s, font_header=36.0 * s, font_subject=42.0 * s,
        font_bar=36.0 * s, font_attach=26.0 * s, font_empty=90.0 * s,
        leading=40.0 * s, label_x=30.0 * s, name_x=250.0 * s, email_x=1500.0 * s,
        quote_indent=30.0 * s, bar_width=4.0 * s, header_rule=12.0 * s,
        attach_bar_h=50.0 * s,
    )


def sw(text: str, font: str, size: float) -> float:
    return pdfmetrics.stringWidth(text or "", font, size)


def wrap_words(text: str, font: str, size: float, max_width: float) -> List[str]:
    words = (text or "").split()
    if not words:
        return [""]
    lines: List[str] = []
    cur = ""
    for word in words:
        test = f"{cur} {word}".strip()
        if sw(test, font, size) <= max_width or not cur:
            if sw(word, font, size) > max_width and not cur:
                buf = ""
                for ch in word:
                    if buf and sw(buf + ch, font, size) > max_width:
                        lines.append(buf)
                        buf = ch
                    else:
                        buf += ch
                cur = buf
            else:
                cur = test
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines or [""]


def replace_emojis(text: str) -> str:
    if not text:
        return text
    for emoji, repl in EMOJI_MAP.items():
        text = text.replace(emoji, repl)
    return re.sub(r"[\U0001F000-\U0001FFFF]", "", text)


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = replace_emojis(text)
    text = re.sub(r"[\u00A0\u2000-\u200B\u2028\u2029\uFEFF]", " ", text)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text

def process_plain_text_blockquotes(text: str) -> str:
    if not text:
        return ""
    result: List[str] = []
    level = 0
    for line in text.split("\n"):
        if not line.strip():
            result.append("")
            continue
        content = line
        new_level = 0
        while re.match(r"^\s*>\s*", content):
            new_level += 1
            content = re.sub(r"^\s*>\s*", "", content, count=1)
        while level < new_level:
            result.append("__BLOCKQUOTE_START__")
            level += 1
        while level > new_level:
            result.append("__BLOCKQUOTE_END__")
            level -= 1
        result.append(content)
    while level > 0:
        result.append("__BLOCKQUOTE_END__")
        level -= 1
    return "\n".join(result)

#
# def process_plain_text_blockquotes(text: str) -> str:
#     if not text:
#         return ""
#     result: List[str] = []
#     level = 0
#     for line in text.split("\n"):
#         content = line
#         new_level = 0
#         while re.match(r"^\s*>\s*", content):
#             new_level += 1
#             content = re.sub(r"^\s*>\s*", "", content, count=1)
#         while level < new_level:
#             result.append("__BLOCKQUOTE_START__")
#             level += 1
#         while level > new_level:
#             result.append("__BLOCKQUOTE_END__")
#             level -= 1
#         result.append(content)
#     while level > 0:
#         result.append("__BLOCKQUOTE_END__")
#         level -= 1
#     return "\n".join(result)


def handle_text(text: str) -> str:
    if not text:
        return ""
    return process_plain_text_blockquotes(text.replace("\r", ""))

def process_html(html: str) -> str:
    if not html:
        return ""

    html = re.sub(r"=\r?\n", "", html)
    html = re.sub(r"=\s+$", "", html, flags=re.MULTILINE)
    html = re.sub(r"=([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), html)

    html = re.sub(r"<br\s*/?>", "<br>", html, flags=re.IGNORECASE)
    html = re.sub(r"\s*<br>\s*$", "<br>", html, flags=re.IGNORECASE | re.MULTILINE)
    html = re.sub(r"^\s*<br>\s*", "", html, flags=re.IGNORECASE | re.MULTILINE)
    html = re.sub(r"<br>", "\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<br>\s*=20\s*<br>", "\n\n", html, flags=re.IGNORECASE)

    html = re.sub(r"</p>", "\n\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<blockquote[^>]*>", "\n__BLOCKQUOTE_START__", html, flags=re.IGNORECASE)
    html = re.sub(r"</blockquote>", "\n__BLOCKQUOTE_END__", html, flags=re.IGNORECASE)
    html = re.sub(r"</div>", "\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<div[^>]*>", "\n", html, flags=re.IGNORECASE)

    html = re.sub(r"</p>", "\n\n", html, flags=re.IGNORECASE)
    html = re.sub(r"</h[1-6]>", "\n\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<h[1-6][^>]*>", "\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<head[^>]*>.*?</head>", "", html, count=1, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.IGNORECASE | re.DOTALL)

    html = re.sub(
        r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        r"\2 (\1)",
        html,
        flags=re.IGNORECASE,
    )

    html = re.sub(r"<[^>]+>", "", html)

    html = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), html)
    html = html.replace("&quot;", '"')
    html = html.replace("&amp;", "&")
    html = html.replace("&lt;", "<")
    html = html.replace("&gt;", ">")
    html = html.replace("&nbsp;", " ")

    html = html.replace("\r\n", "\n")
    html = re.sub(r"\n+[^\S\n\r]*\n+[^\S\n\r]*\n+[^\S\n\r]*\n*", "\n\n", html)
    html = html.strip()
    html = re.sub(r"[ \t]+", " ", html)
    html = re.sub(r"^[ \t]+", "", html, flags=re.MULTILINE)
    html = re.sub(r"[ \t]+$", "", html, flags=re.MULTILINE)

    empty_quote = re.compile(
        r"\n?__BLOCKQUOTE_START__[ \t]*(?:\n[ \t]*)*__BLOCKQUOTE_END__\n?",
    )
    while True:
        nxt = empty_quote.sub("", html)
        if nxt == html:
            break
        html = nxt

    return process_plain_text_blockquotes(html)

def normalize_name(name: str) -> str:
    if not name:
        return ""
    name = name.strip().strip("\"'")
    if "," in name:
        parts = name.split(",", 1)
        if len(parts) == 2:
            name = f"{parts[1].strip()} {parts[0].strip()}"
    return name.strip()


def split_name_and_email(text: str) -> Tuple[str, str]:
    at_pos = text.rfind("@")
    if at_pos == -1:
        return text.strip(), ""
    email_start = at_pos
    while email_start > 0 and text[email_start - 1] not in " \t":
        email_start -= 1
    return text[:email_start].strip(), text[email_start:].strip()


def parse_email_list(input_str: str) -> List[Dict[str, str]]:
    if not input_str or not input_str.strip():
        return []
    recipients: List[str] = []
    current = ""
    seen_at = False
    for char in input_str:
        if char == ",":
            if seen_at:
                if current.strip():
                    recipients.append(current.strip())
                current = ""
                seen_at = False
            else:
                current += char
        else:
            current += char
            if char == "@":
                seen_at = True
    if current.strip():
        recipients.append(current.strip())
    result = []
    for rec in recipients:
        name, email = split_name_and_email(rec)
        if email:
            result.append({"name": normalize_name(name), "email": email})
    return result


def resolve_display_name(raw_name: str, email: str) -> str:
    if not email:
        return "DEBUG_MISSING_16666EMAIL666611" if debug else ""
    raw_name = normalize_name(raw_name)
    lower = email.lower().strip()
    if lower in email_to_name:
        display = email_to_name[lower]
    elif debug:
        display = "DEBUG_MISSING_NAME_44444444444444"
    else:
        cleaned = raw_name.strip().strip("\"'")
        display = "" if cleaned == email else cleaned
    if debug and display == "":
        display = "DEBUG_MISSING_NAME_44444444444444"
    return display


def decode_bytes(part: Message) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        return ""
    charset = part.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def extract_filename(part: Message) -> str:
    disp = part.get("Content-Disposition", "") or ""
    ctype = part.get("Content-Type", "") or ""
    name = ""
    parts = re.findall(r"filename\*(\d*)\*?=['\"]?([^\"';\r\n]+)", disp, flags=re.I)
    if parts:
        parts = sorted(parts, key=lambda x: int(x[0] or 0))
        name = "".join(p[1] for p in parts)
        name = re.sub(r"^[^']*'[^']*'", "", name)
        name = re.sub(r"%([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), name)
    if not name:
        m = re.search(r"filename\*=([^']+)'[^']*'([^;]+)", disp, flags=re.I)
        if m:
            name = re.sub(r"%([0-9A-Fa-f]{2})", lambda x: chr(int(x.group(1), 16)), m.group(2))
    if not name:
        m = re.search(r"filename=['\"]?([^\"';\r\n]+)", disp, flags=re.I)
        if m:
            name = m.group(1)
    if not name:
        parts = re.findall(r"name\*(\d*)\*?=['\"]?([^\"';\r\n]+)", ctype, flags=re.I)
        if parts:
            parts = sorted(parts, key=lambda x: int(x[0] or 0))
            name = "".join(p[1] for p in parts)
            name = re.sub(r"%([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), name)
    if not name:
        m = re.search(r"name=['\"]?([^\"';\r\n]+)", ctype, flags=re.I)
        if m:
            name = m.group(1)
    if name and re.search(r"=\?", name):
        try:
            decoded = decode_header(name)
            name = "".join(
                t[0].decode(t[1] or "utf-8", errors="replace") if isinstance(t[0], bytes) else t[0]
                for t in decoded
            )
        except Exception:
            pass
    name = re.sub(r"%([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), name)
    return (name or "").strip().strip("\"'") or "attachment.bin"


def ext_of(name: str) -> str:
    if "." not in (name or ""):
        return ""
    return name.rsplit(".", 1)[-1].lower()


def safe_stem(original: str) -> str:
    stem = Path(original).stem if original else "attachment"
    stem = re.sub(r"[^A-Za-z0-9._+-]", "_", stem)
    return (stem or "attachment")[:80]


def save_part_file(
        part: Message, ext: str, temp_dir: Path, kind: str, email_count: int,
) -> Optional[Attachment]:
    try:
        payload = part.get_payload(decode=True)
        if payload is None:
            log("WARNING", f"empty payload, ext={q(ext)}", review=True)
            return None
        original = extract_filename(part) or f"attachment.{ext}"
        digest = sha_bytes(payload)
        fname = f"{email_count}_{digest[:12]}_{safe_stem(original)}.{ext or 'bin'}"
        sub = "images" if kind == "image" else "pdf_attachments"
        path = temp_dir / sub / fname
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            path = temp_dir / sub / f"{email_count}_{digest}_{safe_stem(original)}.{ext or 'bin'}"
        path.write_bytes(payload)
        att = Attachment(original, path, path.stat().st_size, kind, sha256=digest)
        log("VERBOSE", f"saved {kind}: {q(path)} as {q(original)} ({att.size} bytes)")
        return att
    except Exception as e:
        log("ERROR", f"Failed to save attachment: {e}")
        if debug:
            traceback.print_exc()
        return None


def convert_document_to_pdf(
        input_path: Path, original_name: str, converter: str, temp_dir: Path,
) -> Optional[Attachment]:
    outdir = temp_dir / "converted"
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        converter, "--headless", "--invisible", "--convert-to", "pdf",
        "--outdir", str(outdir), str(input_path),
    ]
    log("INFO", f"calling open office: {' '.join(q(x) for x in cmd)}")
    log("INFO", f"LibreOffice source {q(input_path)} ({input_path.stat().st_size} bytes) name={q(original_name)}")
    rc, out = run_cmd(cmd)
    pdf_path = outdir / (input_path.stem + ".pdf")
    if rc != 0 or not pdf_path.exists():
        log("ERROR", f"LibreOffice failed to convert {q(original_name)} path={q(input_path)}: {out[:500]}")
        return None
    att = Attachment(original_name, pdf_path, pdf_path.stat().st_size, "pdf", ext_of(original_name), file_sha(pdf_path))
    log("INFO", f"LibreOffice wrote {q(pdf_path)} ({att.size} bytes)")
    return att


def unfold_lines(raw: str) -> str:
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n[ \t]", "", raw)


def calendar_addendum(raw: str, filename: str) -> str:
    raw = unfold_lines(raw)
    fields: Dict[str, List[str]] = {k: [] for k in (
        "SUMMARY", "DTSTART", "DTEND", "LOCATION", "DESCRIPTION",
        "ORGANIZER", "STATUS", "ATTENDEE",
    )}
    for line in raw.split("\n"):
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        params = key.split(";")
        name = params[0].upper()
        tz = ""
        for p in params[1:]:
            if p.upper().startswith("TZID="):
                tz = p.split("=", 1)[1]
        if tz and name in ("DTSTART", "DTEND"):
            val = f"{val} ({tz})"
        if name in fields:
            fields[name].append(val.strip())
    lines = ["", "—— Kalender ——", f"Datei: {filename}"]
    mapping = [
        ("SUMMARY", "Titel"), ("DTSTART", "Beginn"), ("DTEND", "Ende"),
        ("LOCATION", "Ort"), ("ORGANIZER", "Organisator"),
        ("ATTENDEE", "Teilnehmer"), ("STATUS", "Status"),
        ("DESCRIPTION", "Beschreibung"),
    ]
    for key, label in mapping:
        vals = [v for v in fields[key] if v]
        if vals:
            lines.append(f"{label}: {'; '.join(vals[:12])}")
    return "\n".join(lines) if len(lines) > 3 else ""


def vcard_addendum(raw: str, filename: str) -> str:
    raw = unfold_lines(raw)
    got: Dict[str, List[str]] = {k: [] for k in (
        "FN", "N", "ORG", "TITLE", "TEL", "EMAIL", "ADR", "URL", "NOTE",
    )}
    for line in raw.split("\n"):
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        name = key.split(";", 1)[0].upper()
        if name in got and val.strip():
            got[name].append(val.strip().replace(";", " ").strip())
    lines = ["", "—— Kontakt ——", f"Datei: {filename}"]
    if got["FN"]:
        lines.append(f"Name: {got['FN'][0]}")
    elif got["N"]:
        lines.append(f"Name: {got['N'][0]}")
    for key, label in (("ORG", "Organisation"), ("TITLE", "Titel"),
                       ("TEL", "Telefon"), ("EMAIL", "E-Mail"),
                       ("ADR", "Adresse"), ("URL", "URL"), ("NOTE", "Notiz")):
        if got[key]:
            lines.append(f"{label}: {'; '.join(got[key][:8])}")
    return "\n".join(lines) if len(lines) > 3 else ""


def zip_addendum(att: Attachment) -> str:
    digest = att.sha256 or file_sha(att.saved_path)
    att.sha256 = digest
    return (
        f"\n—— Archiv ——\nDatei: {att.original_name}\n"
        f"Größe: {att.size} Bytes\nSHA-256: {digest}\n"
    )


def placeholder_label(att: Attachment) -> str:
    if att.type == "image":
        return "Image"
    if att.type == "pdf" and att.converted_from:
        return "Document"
    if att.type == "pdf":
        return "PDF"
    if att.type == "ics":
        return "Appointment"
    if att.type == "vcf":
        return "Business Card"
    if att.type == "zip":
        return "Archive"
    if att.converted_from or ext_of(att.original_name) in CONVERTIBLE:
        return "Document"
    return "Attachment"


def iter_logical_parts(msg: Message):
    if not msg.is_multipart():
        yield msg
        return
    payload = msg.get_payload()
    if not isinstance(payload, list):
        return
    for part in payload:
        if not isinstance(part, Message):
            continue
        pct = (part.get_content_type() or "").lower()
        if pct.startswith("message/rfc822"):
            yield part
        elif part.is_multipart():
            yield from iter_logical_parts(part)
        else:
            yield part


def handle_part(
        part: Message, temp_dir: Path, converter: str,
        buckets: Dict[str, Any], email_count: int,
) -> None:
    pct = (part.get_content_type() or "").lower()
    original = extract_filename(part)
    ext = ext_of(original)
    log("DEBUG", f"part ct={q(pct)} name={q(original)} ext={q(ext)}")

    if pct.startswith("multipart/"):
        return
    if pct == "text/plain":
        raw = decode_bytes(part)
        buckets["orig_plain"] += raw + "\n\n"
        clean = handle_text(raw)
        if clean:
            buckets["plain"] += clean + "\n\n"
        return
    if pct == "text/html":
        raw = decode_bytes(part)
        buckets["orig_html"] += raw + "\n\n"
        clean = process_html(raw)
        if clean:
            buckets["html"] += clean + "\n\n"
        return

    nested = extract_nested_message(part)
    rfc822 = pct.startswith("message/rfc822") or pct == "message/news"
    if nested is not None and (rfc822 or isinstance(part.get_payload(), list)):
        log("VERBOSE", f"attached email from {q(original)} ct={q(pct)}")
        walk_message(nested, temp_dir, converter, 1, buckets, email_count)
        return
    if rfc822:
        log("WARNING", f"attached email had no parseable payload: {q(original)}", review=True)
        return

    guess_ext = ext or "bin"
    if pct.startswith("image/"):
        guess_ext = IMAGE_EXT.get(pct.split("/", 1)[-1], ext or "jpg")
    folder_kind = "image" if (
            pct.startswith("image/") or ext in IMAGE_EXT or guess_ext in IMAGE_EXT.values()
    ) else "other"
    att = save_part_file(part, guess_ext, temp_dir, folder_kind, email_count)
    if not att:
        nested = extract_nested_message(part)
        if nested is not None:
            log("VERBOSE", f"attached email recovered after empty payload: {q(original)}")
            walk_message(nested, temp_dir, converter, 1, buckets, email_count)
        return

    kind = sniff_kind(att.saved_path, ext, pct)
    dest_sub = "images" if kind == "image" else "pdf_attachments"
    if att.saved_path.parent.name != dest_sub:
        dest = temp_dir / dest_sub / att.saved_path.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists():
            dest = temp_dir / dest_sub / f"{att.saved_path.stem}_{att.sha256[:12]}{att.saved_path.suffix}"
        try:
            att.saved_path = att.saved_path.replace(dest)
        except Exception as e:
            log("WARNING", f"could not move attachment into {q(dest_sub)}: {e}")
    att.type = kind if kind != "office" else "other"
    buckets["all"].append(att)

    if kind == "pdf":
        att.type = "pdf"
        buckets["pdfs"].append(att)
        return
    if kind == "image":
        att.type = "image"
        buckets["images"].append(att)
        return
    if kind == "office":
        if fast_level >= 3:
            att.type = "other"
            att.converted_from = ext or "office"
            return
        converted = convert_document_to_pdf(att.saved_path, att.original_name, converter, temp_dir)
        if converted:
            att.converted_from = ext or "office"
            att.converted_size = converted.size
            att.converted_path = converted.saved_path
            buckets["pdfs"].append(converted)
        return
    if kind == "eml":
        att.type = "eml"
        nested_msg = extract_nested_message(part)
        if nested_msg is None:
            try:
                nested_msg = BytesParser(policy=policy.default).parsebytes(att.saved_path.read_bytes())
            except Exception as e:
                log("DEBUG", f"eml file parse: {e}")
        if nested_msg is not None:
            try:
                walk_message(nested_msg, temp_dir, converter, 1, buckets, email_count)
            except Exception as e:
                log("ERROR", f"Failed to parse attached email: {e}")
        return
    if kind == "ics":
        att.type = "ics"
        try:
            raw = att.saved_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            raw = ""
        add = calendar_addendum(raw, att.original_name)
        if add:
            buckets["plain"] += add + "\n\n"
            buckets["orig_plain"] += add + "\n\n"
        return
    if kind == "vcf":
        att.type = "vcf"
        try:
            raw = att.saved_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            raw = ""
        add = vcard_addendum(raw, att.original_name)
        if add:
            buckets["plain"] += add + "\n\n"
            buckets["orig_plain"] += add + "\n\n"
        return
    if kind == "zip":
        att.type = "zip"
        add = zip_addendum(att)
        buckets["plain"] += add + "\n\n"
        buckets["orig_plain"] += add + "\n\n"
        return
    if kind == "video":
        att.type = "video"
        log("VERBOSE", f"video saved for hash only: {q(original)}")
        return
    if pct in ("message/delivery-status", "message/disposition-notification"):
        log("WARNING", f"delivery-status stored as other: {q(original)}", review=True)
        return
    log("WARNING", f"unknown attachment hashed: {q(original)} (type: {pct})", review=True)


def walk_message(
        msg: Message, temp_dir: Path, converter: str,
        nested: int, buckets: Dict[str, Any], email_count: int,
) -> None:
    if nested == 0:
        buckets.update({
            "plain": "", "html": "", "orig_plain": "", "orig_html": "",
            "images": [], "pdfs": [], "all": [],
        })
    if not msg.is_multipart():
        handle_part(msg, temp_dir, converter, buckets, email_count)
        return
    for part in iter_logical_parts(msg):
        handle_part(part, temp_dir, converter, buckets, email_count)


def decode_header_value(msg: Message, key: str, default: str = "") -> str:
    raw = msg.get(key, default) or default
    try:
        decoded = decode_header(raw)
        raw = "".join(
            t[0].decode(t[1] or "utf-8", errors="replace") if isinstance(t[0], bytes) else t[0]
            for t in decoded
        )
    except Exception:
        pass
    return raw


def format_msg_date(msg: Message) -> str:
    try:
        return parsedate_to_datetime(msg.get("Date", "") or "").strftime("%d.%m.%Y  %H:%M")
    except Exception:
        return (msg.get("Date") or "Unknown date")[:30]


def process_single_message(
        msg: Message, count: int, temp_dir: Path, converter: str,
) -> ProcessedEmail:
    subject = clean_text(decode_header_value(msg, "Subject", "(no subject)"))
    date_raw, date_dt = parse_email_datetime(msg)
    date_str = format_msg_date(msg)
    from_str = clean_text(decode_header_value(msg, "From"))
    set_log_ctx(count, date_str, subject or "(no subject)", from_str)
    buckets: Dict[str, Any] = {}
    walk_message(msg, temp_dir, converter, 0, buckets, count)
    if prefer_plain_text:
        text = buckets["plain"] if buckets["plain"].strip() else buckets["html"]
    else:
        text = buckets["html"] if buckets["html"].strip() else buckets["plain"]
    text = clean_text(text)
    to_str = clean_text(decode_header_value(msg, "To"))
    cc_str = clean_text(decode_header_value(msg, "Cc"))
    from_str = re.sub(r"<(.+?)>", r"\1", from_str)
    from_str = re.sub(r'^"(.+?)"\s*', r"\1 ", from_str)
    to_str = re.sub(r'^"(.+?)"\s*', r"\1 ", to_str)
    to_str = re.sub(r"<(.+?)>", r"\1", to_str)
    cc_str = re.sub(r'^"(.+?)"\s*', r"\1 ", cc_str)
    cc_str = re.sub(r"<(.+?)>", r"\1", cc_str)
    set_log_ctx(count, date_str, subject or "(no subject)", from_str)
    log("VERBOSE", f"parsed images={len(buckets['images'])} pdfs={len(buckets['pdfs'])} all={len(buckets['all'])}")
    return ProcessedEmail(
        count=count, date_str=date_str, date_raw=date_raw, date_dt=date_dt,
        from_str=from_str, to_str=to_str,
        cc_str=cc_str, subject=subject or "(no subject)", body_text=text,
        original_plain=buckets["orig_plain"], original_html=buckets["orig_html"],
        images=buckets["images"], pdf_attachments=buckets["pdfs"],
        all_attachments=buckets["all"],
    )


def recipient_line_count(layout: Layout, raw: str) -> int:
    address = re.sub(r"[<>]", "", raw or "")
    m = re.search(r"^(.*?)(\S+@\S+)$", address)
    if not m:
        return 1
    display = resolve_display_name(m.group(1), m.group(2))
    wide = layout.width - layout.name_x - 10 * layout.label_x
    lines = wrap_words(display, FONT_REGULAR, layout.font_header, wide)
    last = lines[-1] if lines else ""
    last_w = sw(last, FONT_REGULAR, layout.font_header)
    fits = (layout.name_x + last_w + layout.label_x) < layout.email_x
    n = max(1, len(lines))
    return n + (0 if fits else 1)


def bar_first_baseline(box_bottom: float, box_h: float, nlines: int, line_h: float, font: str, size: float) -> float:
    ascent = pdfmetrics.getAscent(font) * size / 1000.0
    descent = abs(pdfmetrics.getDescent(font) * size / 1000.0)
    ink_h = (nlines - 1) * line_h + ascent + descent
    top_ink = box_bottom + (box_h + ink_h) / 2.0
    return top_ink - ascent + BAR_OPTICAL_NUDGE * size


def draw_title_bar(
        c: rl_canvas.Canvas,
        layout: Layout,
        title: str,
        font_size: Optional[float] = None,
) -> None:
    s = layout.width / PERL_REF_W
    size = 24.0 * s if font_size is None else font_size
    max_w = layout.width - 80.0 * s
    lines = wrap_words(title, FONT_BOLD, size, max_w) or [""]
    line_h = size * 1.25
    pad = 16.0 * s
    bar_h = max(50.0 * s, 2.0 * pad + len(lines) * line_h)
    bar_top = layout.height - 10.0 * s
    bar_bottom = bar_top - bar_h
    c.setFillColor(DARK_BLUE)
    c.rect(30.0 * s, bar_bottom, layout.width - 60.0 * s, bar_h, fill=1, stroke=0)
    y = bar_first_baseline(bar_bottom, bar_h, len(lines), line_h, FONT_BOLD, size)
    y -= TITLE_BAR_TEXT_NUDGE * s
    c.setFillColor(white)
    c.setFont(FONT_BOLD, size)
    for line in lines:
        c.drawString(40.0 * s, y, line)
        y -= line_h


def draw_placeholder_page(c: rl_canvas.Canvas, layout: Layout, title: str, label: str) -> None:
    new_page(c, layout)
    draw_title_bar(c, layout, title)
    c.setFillColor(HexColor(MUTED_GREY))
    c.setFont(FONT_BOLD, layout.font_empty)
    c.drawCentredString(layout.width / 2.0, layout.height / 2.0, label)


def draw_recipient_line(
        c: rl_canvas.Canvas, layout: Layout, raw: str, y: float, is_cc: bool = False,
) -> float:
    address = re.sub(r"[<>]", "", raw or "")
    m = re.search(r"^(.*?)(\S+@\S+)$", address)
    if not m:
        c.setFont(FONT_REGULAR, layout.font_header)
        c.setFillColor(LIGHT_GRAY)
        c.drawString(layout.name_x, y, address)
        return y - layout.line_height
    raw_name, email = m.group(1), m.group(2)
    display = resolve_display_name(raw_name, email)
    wide_width = layout.width - layout.name_x - 10 * layout.label_x
    lines = wrap_words(display, FONT_REGULAR, layout.font_header, wide_width)
    last = lines[-1] if lines else ""
    last_w = sw(last, FONT_REGULAR, layout.font_header)
    email_fits = (layout.name_x + last_w + layout.label_x) < layout.email_x
    n = len(lines) if lines else 1
    if not email_fits:
        n += 1
    yy = y
    for i in range(n):
        line_text = lines[i] if i < len(lines) else ""
        c.setFont(FONT_REGULAR, layout.font_header)
        c.setFillColor(LIGHT_GRAY)
        if i == n - 1:
            email_text = f"{email} (CC)" if is_cc else email
            if email_fits:
                c.drawString(layout.name_x, yy, line_text)
                c.setFont(FONT_ITALIC, layout.font_header)
                c.drawString(layout.email_x, yy, email_text)
            else:
                if line_text:
                    c.drawString(layout.name_x, yy, line_text)
                    yy -= layout.line_height
                c.setFont(FONT_ITALIC, layout.font_header)
                c.drawString(layout.email_x, yy, email_text)
        else:
            c.drawString(layout.name_x, yy, line_text)
            yy -= layout.line_height
    return yy - layout.line_height


def write_header(
        c: rl_canvas.Canvas, layout: Layout, email: ProcessedEmail, page_number: int,
) -> float:
    lh = layout.line_height
    s = layout.width / PERL_REF_W
    to_block = parse_email_list(email.to_str)
    cc_block = parse_email_list(email.cc_str)

    extra = recipient_line_count(layout, email.from_str) - 1
    if email.to_str.strip():
        extra += 1
        for rec in to_block:
            extra += recipient_line_count(layout, f"{rec['name']} {rec['email']}")
        extra -= 1
    for rec in cc_block:
        extra += recipient_line_count(layout, f"{rec['name']} {rec['email']}")

    blue_h = 3 * lh + extra * lh
    if blue_h < 3 * lh:
        blue_h = 3 * lh

    c.setFillColor(DARK_BLUE)
    c.rect(0, layout.height - blue_h, layout.width, blue_h, fill=1, stroke=0)

    c.setFillColor(LIGHT_GRAY)
    c.setFont(FONT_BOLD, layout.font_header)
    c.drawString(layout.label_x, layout.height - lh, f"Email #{email.count}")
    c.drawRightString(layout.width - layout.label_x, layout.height - lh, email.date_str)
    c.drawRightString(layout.width - layout.label_x, layout.height - 2 * lh, f"Page {page_number}")

    y = layout.height - 2 * lh
    c.setFont(FONT_BOLD, layout.font_header)
    c.drawString(layout.label_x, y, "From:")
    y = draw_recipient_line(c, layout, email.from_str, y, False)
    if email.to_str.strip():
        c.setFont(FONT_BOLD, layout.font_header)
        c.setFillColor(LIGHT_GRAY)
        c.drawString(layout.label_x, y, "To:")
        for rec in to_block:
            y = draw_recipient_line(c, layout, f"{rec['name']} {rec['email']}", y, False)
    if email.cc_str.strip():
        for rec in cc_block:
            y = draw_recipient_line(c, layout, f"{rec['name']} {rec['email']}", y, True)

    subj_lines = wrap_words(
        email.subject or "(no subject)", FONT_BOLD, layout.font_subject,
        layout.width - 2 * layout.label_x,
        ) or [""]
    subj_line_h = layout.font_subject * 1.25
    pad = 18.0 * s
    subj_h = max(lh + pad, 2.0 * pad + len(subj_lines) * subj_line_h)
    band_top = layout.height - blue_h
    band_bottom = band_top - subj_h

    sy = bar_first_baseline(
        band_bottom, subj_h, len(subj_lines), subj_line_h, FONT_BOLD, layout.font_subject,
    )
    c.setFillColor(black)
    c.setFont(FONT_BOLD, layout.font_subject)
    for line in subj_lines:
        c.drawString(layout.label_x, sy, line)
        sy -= subj_line_h

    c.setStrokeColor(DARK_BLUE)
    c.setLineWidth(layout.header_rule)
    c.line(0, band_bottom, layout.width, band_bottom)
    return band_bottom - 2 * lh


def new_page(c: rl_canvas.Canvas, layout: Layout) -> None:
    c.showPage()
    c.setPageSize((layout.width, layout.height))

def write_long_text(
        c: rl_canvas.Canvas, layout: Layout, email: ProcessedEmail, start_y: float,
) -> float:
    text = (email.body_text or "").strip()
    if len(text) <= 2:
        return start_y

    page_num = 1
    y = start_y
    level = 0
    bq_starts: List[float] = []
    margin = layout.margin
    s = layout.width / PERL_REF_W
    leading = layout.leading
    bottom = margin + 80 * s

    def bar_x(lv: int) -> float:
        return margin + 8 * s + (lv - 1) * 12 * s

    def stroke_level(lv: int, y0: float, y1: float) -> None:
        c.setStrokeColor(DARK_BLUE)
        c.setLineWidth(layout.bar_width)
        c.line(bar_x(lv), y0 + 6 * s, bar_x(lv), y1 + 6 * s)

    def stroke_open_to(y1: float) -> None:
        for lv in range(1, level + 1):
            if lv - 1 < len(bq_starts):
                stroke_level(lv, bq_starts[lv - 1], y1)

    def stroke_line_tick(y_line: float) -> None:
        c.setStrokeColor(DARK_BLUE)
        c.setLineWidth(layout.bar_width)
        for lv in range(1, level + 1):
            c.line(bar_x(lv), y_line + 6 * s, bar_x(lv), y_line - leading - 2 * s)

    def new_text_page() -> None:
        nonlocal y, page_num
        stroke_open_to(y)
        new_page(c, layout)
        page_num += 1
        y = write_header(c, layout, email, page_num)
        bq_starts[:] = [y] * level

    def ensure_space() -> None:
        if y < bottom:
            new_text_page()

    for line in text.split("\n"):
        ensure_space()

        if "__BLOCKQUOTE_START__" in line:
            level += 1
            bq_starts.append(y)
            continue

        if "__BLOCKQUOTE_END__" in line:
            if level > 0:
                start = bq_starts.pop() if bq_starts else y
                stroke_level(level, start, y)
                level -= 1
            continue

        if len(line) == 0:
            y -= leading * 0.7
            continue

        x = margin + level * layout.quote_indent
        width = layout.width - 2 * margin - level * layout.quote_indent - 10 * s
        stroke_line_tick(y)

        wrapped = wrap_words(line, FONT_REGULAR, layout.font_body, width)
        first = True
        for chunk in wrapped:
            if not first:
                y -= leading
                ensure_space()
            first = False
            c.setFillColor(black)
            c.setFont(FONT_REGULAR, layout.font_body)
            c.drawString(x, y, chunk)

        y -= leading

    stroke_open_to(y)
    return y

def human_size(n: int) -> str:
    if not n:
        return "0 B"
    units = ["B", "KB", "MB", "GB"]
    f = float(n)
    i = 0
    while f >= 1024 and i < len(units) - 1:
        f /= 1024
        i += 1
    return f"{f:.1f} {units[i]}"


def draw_attachment_summary(
        c: rl_canvas.Canvas, layout: Layout, email: ProcessedEmail, y: float,
) -> float:
    if not email.all_attachments:
        return y
    s = layout.width / PERL_REF_W


    y -= 150 * s
    rule_y = y - ATTACH_HEADING_RULE_NUDGE * s
    c.setStrokeColor(DARK_BLUE)
    c.setLineWidth(layout.bar_width)
    c.line(50 * s, rule_y, layout.width - 50 * s, rule_y)
    y -= 25 * s
    c.setFillColor(black)
    c.setFont(FONT_BOLD, layout.font_header)
    c.drawString(30 * s, y + 35 * s, "Anhänge")
    y -= 25 * s
    c.setFont(FONT_REGULAR, layout.font_attach)
    for i, att in enumerate(email.all_attachments, 1):
        if att.converted_size:
            status = f"Converted → PDF {human_size(att.converted_size)}"
        elif att.type == "pdf":
            status = "Eingefügt"
        elif att.type == "image":
            status = "Angefügt"
        elif att.type == "ics":
            status = "Kalender"
        elif att.type == "vcf":
            status = "Kontakt"
        elif att.type == "zip":
            status = "Archiv"
        else:
            status = "Als Hash"
        extra = ""
        if att.sha256 and att.type not in {"image", "pdf"} and not att.converted_size:
            extra = f"  [{att.sha256[:12]}]"
        line = f"{i}. {att.original_name}   ({human_size(att.size)})   [{status}]{extra}"
        c.drawString(layout.margin + 12 * s, y, line)
        y -= layout.leading
    return y


def magick_bin() -> str:
    return shutil.which("magick") or shutil.which("convert") or "magick"


def identify_size(src: Path) -> Tuple[int, int]:
    src_arg = f"{src}[0]" if src.suffix.lower() == ".gif" else str(src)
    rc, out = run_cmd(
        [magick_bin(), src_arg, "-auto-orient", "-format", "%w %h", "info:"],
    )
    w = h = 0
    if rc == 0:
        parts = out.strip().replace(",", ".").split()
        try:
            w, h = int(float(parts[0])), int(float(parts[1]))
        except Exception:
            pass
    return w, h


def identify_image(src: Path) -> Tuple[int, int, float, float]:
    src_arg = f"{src}[0]" if src.suffix.lower() == ".gif" else str(src)
    rc, out = run_cmd(
        [magick_bin(), src_arg, "-auto-orient", "-format",
         "%w %h %[resolution.x] %[resolution.y] %[units]", "info:"],
    )
    w = h = 0
    dx = dy = 0.0
    units = ""
    if rc == 0:
        parts = out.strip().replace(",", ".").split()
        try:
            w, h = int(float(parts[0])), int(float(parts[1]))
            if len(parts) >= 4:
                dx, dy = float(parts[2] or 0), float(parts[3] or 0)
            if len(parts) >= 5:
                units = "".join(parts[4:]).lower()
        except Exception:
            pass
    if "centimeter" in units or units in {"pixelspercentimeter", "ppc"}:
        dx, dy = dx * 2.54, dy * 2.54
        log("DEBUG", "dpi was pixels/cm, converted to inch")
    if dx == 0 or dy == 0:
        log("WARNING", "img_dpi_x: invalid, set to default 72", review=True)
        log("WARNING", "img_dpi_y: invalid, set to default 72", review=True)
        dx, dy = 72.0, 72.0
    elif fast_level == 0 and (dx < 50 or dy < 50):
        log("WARNING", f"img_dpi_x: {dx} probably too low, increase by factor 10", review=True)
        log("WARNING", f"img_dpi_y: {dy} probably too low, increase by factor 10", review=True)
        dx, dy = dx * 10, dy * 10
    log("DEBUG", f"iw: {w}")
    log("DEBUG", f"ih: {h}")
    log("DEBUG", f"img_dpi_x: {dx}")
    log("DEBUG", f"img_dpi_y: {dy}")
    return w, h, dx, dy


def calculate_y_value(layout: Layout, text_length: int, kind: str = "picture") -> float:
    infobox_bottom = layout.height - layout.height * 0.05
    infobox_height = layout.height - infobox_bottom
    y_buffer = 50.0 * (layout.width / PERL_REF_W) if kind == "picture" else 0.0
    if text_length > 150:
        return infobox_bottom - 2 * infobox_height - y_buffer
    if text_length > 0:
        return infobox_bottom - infobox_height - y_buffer
    return infobox_bottom - y_buffer


def png_write_args() -> List[str]:
    if finalize:
        return ["-quality", "100"]
    return ["-compress", "None", "-define", "png:compression-level=0"]


def add_images_to_pdf(
        c: rl_canvas.Canvas,
        layout: Layout,
        images_ref: List[Attachment],
        text_length: int,
        temp_dir: Path,
        use_montages_for_images: bool,
        font_path: Optional[Path] = None,
) -> int:
    images = list(images_ref)
    if not images:
        return 0

    size_x = layout.width
    size_y = layout.height
    s = size_x / PERL_REF_W
    document_dpi = 300.0 * s
    raster = float(layout.image_dpi)
    max_per_page = 9 if use_montages_for_images else 1
    top_margin = 110.0 * s
    magick = magick_bin()
    out_dir = temp_dir / "montages"
    out_dir.mkdir(parents=True, exist_ok=True)
    converted = temp_dir / "converted"
    converted.mkdir(parents=True, exist_ok=True)
    png_out = png_write_args()

    img_index = 0
    batch_i = 0
    while img_index < len(images):
        page_images = images[img_index: min(img_index + max_per_page, len(images))]
        arr_size = len(page_images)
        label = "Bild: " if arr_size == 1 else "Bilder: "
        names = ", ".join(im.original_name for im in page_images)
        title = label + names

        if fast_level >= 2:
            draw_placeholder_page(c, layout, title, "Image")
            img_index += max_per_page
            batch_i += 1
            continue

        new_page(c, layout)
        draw_title_bar(c, layout, title)
        y_value = calculate_y_value(layout, text_length, "picture")

        if arr_size == 1:
            cols, rows = 1, 1
        elif arr_size == 2:
            cols, rows = 1, 2
        elif arr_size <= 4:
            cols, rows = 2, 2
        elif arr_size <= 6:
            cols, rows = 3, 2
        else:
            cols, rows = 3, 3

        available_width = size_x - 100.0 * s
        available_height = y_value - top_margin - 60.0 * s
        if available_height < 80.0 * s:
            available_height = size_y - top_margin - 130.0 * s
        quadrant_w = int(available_width / cols)
        quadrant_h = int(available_height / rows)
        padding = int(round(12.0 * s))

        cell_files: List[str] = []
        for img in page_images:
            img_path = str(img.saved_path)
            src_arg = f"{img_path}[0]" if img_path.lower().endswith(".gif") else img_path
            iw, ih, dpi_x, dpi_y = identify_image(Path(img_path))

            if iw == 0 or ih == 0:
                log("ERROR", f"Skipping image with invalid dimensions: {q(img_path)}")
                ph = converted / "no_text_placeholder.png"
                font_arg = ["-font", str(font_path)] if font_path and font_path.exists() else []
                run_cmd([
                    magick, "-size", "1200x512", "xc:white",
                    *font_arg, "-pointsize", "90", "-fill", MUTED_GREY,
                    "-gravity", "center", "-annotate", "0", "this image is corrupted",
                    *png_out, f"png24:{ph}",
                ])
                img_path = str(ph)
                src_arg = img_path
                iw, ih, dpi_x, dpi_y = identify_image(ph)

            phys_w_in = iw / max(dpi_x, 1e-6)
            phys_h_in = ih / max(dpi_y, 1e-6)
            target_w = int(phys_w_in * document_dpi)
            target_h = int(phys_h_in * document_dpi)
            if target_w > quadrant_w or target_h > quadrant_h:
                scale = min(
                    quadrant_w / max(target_w, 1),
                    quadrant_h / max(target_h, 1),
                    )
                target_w = int(target_w * scale)
                target_h = int(target_h * scale)

            px_q_w = max(1, int(round(quadrant_w * raster / 72.0)))
            px_q_h = max(1, int(round(quadrant_h * raster / 72.0)))
            px_t_w = max(1, int(round(target_w * raster / 72.0)))
            px_t_h = max(1, int(round(target_h * raster / 72.0)))

            cell = converted / f"canvas_{img.saved_path.stem}_{batch_i}.png"
            rc, err = run_cmd([
                magick, src_arg,
                "-auto-orient",
                "-resize", f"{px_t_w}x{px_t_h}",
                "-background", "white",
                "-gravity", "center",
                "-extent", f"{px_q_w}x{px_q_h}",
                *png_out,
                f"png24:{cell}",
            ])
            if rc != 0 or not cell.exists():
                log("ERROR", f"Failed to create canvas: {err[:400]}")
                continue
            cell_files.append(str(cell))

        if not cell_files:
            img_index += max_per_page
            batch_i += 1
            continue

        pad_px = max(0, int(round(padding * raster / 72.0)))
        mont = out_dir / f"montage_{batch_i}.png"
        rc, err = run_cmd([
            magick, "montage", *cell_files,
            "-tile", f"{cols}x{rows}",
            "-geometry", f"{px_q_w}x{px_q_h}+{pad_px}+{pad_px}",
            "-background", "white",
            "-border", "0",
            "-density", str(int(raster)),
            *png_out,
            f"png24:{mont}",
        ])
        if rc != 0 or not mont.exists():
            log("ERROR", f"Failed to create montage: {err[:400]}")
            img_index += max_per_page
            batch_i += 1
            continue
        log("VERBOSE", f"Created montage file: {q(mont)}")

        mw, mh = identify_size(mont)
        if mw <= 0 or mh <= 0:
            log("ERROR", f"montage has invalid size: {q(mont)}")
            img_index += max_per_page
            batch_i += 1
            continue

        mont_w = mw * 72.0 / raster
        mont_h = mh * 72.0 / raster
        position_x = int((size_x - mont_w) / 2)
        position_y = y_value - top_margin - mont_h
        if position_y < 60.0 * s:
            position_y = 60.0 * s
        try:
            c.drawImage(
                str(mont),
                position_x,
                position_y,
                width=mont_w,
                height=mont_h,
                preserveAspectRatio=False,
                mask="auto",
            )
        except Exception as e:
            log("ERROR", f"drawImage montage failed: {e}")

        img_index += max_per_page
        batch_i += 1
    return 1


def is_pdf_file(path: Path) -> bool:
    return read_head(path, 1024).find(b"%PDF-") >= 0


def scale_pdf_to_fit(src: Path, dest: Path, w_pt: float, h_pt: float, gs: str) -> bool:
    if not is_pdf_file(src):
        log("WARNING", f"Skip GS fit, not a PDF: {q(src.name)}", review=True)
        return False
    dest.parent.mkdir(parents=True, exist_ok=True)
    log("VERBOSE", f"GS fit {q(src)} → {q(dest)}")
    cmd = [
        gs, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.7",
        "-dPDFSETTINGS=/prepress", "-dNOPAUSE", "-dBATCH", "-dSAFER",
        "-dFIXEDMEDIA", "-dPDFFitPage",
        f"-dDEVICEWIDTHPOINTS={int(round(w_pt))}",
        f"-dDEVICEHEIGHTPOINTS={int(round(h_pt))}",
        f"-sOutputFile={str(dest)}", str(src),
    ]
    rc, out = run_cmd(cmd)
    if rc != 0 or not is_pdf_file(dest):
        log("ERROR", f"Ghostscript failed to scale {q(src.name)}: {out[:500]}")
        return False
    return True


def make_pdf_bar_page(layout: Layout, filename: str) -> bytes:
    buf = io.BytesIO()
    c = new_canvas(buf, (layout.width, layout.height))
    draw_title_bar(c, layout, f"PDF: {filename}")
    c.save()
    return buf.getvalue()

CLAUSE_RE = re.compile(r"^(!?)(?:(\d+)-(\d+)|-(\d+)|(\d+)-|(\d+))$")


def parse_selector(sel: str) -> List[Tuple[str, int, int]]:
    sel = (sel or "0").strip()
    if sel in ("", "0"):
        return [("inc", 1, 0)]
    clauses: List[Tuple[str, int, int]] = []
    for raw in sel.split(";"):
        piece = raw.strip()
        if not piece:
            raise ValueError(f"empty clause in --select-emails {q(sel)}")
        m = CLAUSE_RE.match(piece)
        if not m:
            raise ValueError(f"invalid --select-emails clause {q(piece)}")
        bang, a, b, open_end_b, open_start_a, single = m.groups()
        kind = "exc" if bang else "inc"
        if a and b:
            start, end = int(a), int(b)
            if start > end:
                raise ValueError(f"range start > end in {q(piece)}")
        elif open_end_b:
            start, end = 1, int(open_end_b)
        elif open_start_a:
            start, end = int(open_start_a), 0
        else:
            start = end = int(single)
        clauses.append((kind, start, end))
    return clauses


def index_selected(i: int, n: int, clauses: List[Tuple[str, int, int]]) -> bool:
    has_inc = any(k == "inc" for k, _, _ in clauses)
    selected = not has_inc
    for kind, start, end in clauses:
        last = n if end == 0 else end
        if kind == "inc" and start <= i <= last:
            selected = True
    for kind, start, end in clauses:
        last = n if end == 0 else end
        if kind == "exc" and start <= i <= last:
            selected = False
    return selected


def msg_sort_ts(m: Message) -> float:
    try:
        return parsedate_to_datetime(m.get("Date") or "").timestamp()
    except Exception:
        return 0.0


def load_mbox_messages(path: Path) -> List[Message]:
    log("VERBOSE", f"Opening mbox with RobustMbox: {q(path)}")
    loaded: List[Message] = []
    mbox = RobustMbox(path)
    try:
        for i, msg in enumerate(mbox, 1):
            if msg.get("X-Robust-Mbox-Error"):
                log(
                    "WARNING",
                    f"unparseable message #{i} in {q(path)}: {msg.get('X-Robust-Mbox-Error')}",
                    review=True,
                )
            loaded.append(msg)
    finally:
        mbox.close()
    log("VERBOSE", f"Read {len(loaded)} messages from {q(path)}")
    return loaded


def mbox_is_fresh(src: Path, dest: Path) -> bool:
    if not dest.exists() or dest.stat().st_size < 1:
        return False
    return dest.stat().st_mtime >= src.stat().st_mtime


def unique_source_messages(
        messages: List[Message], temp_dir: Path, converter: str,
) -> Tuple[List[Message], List[Message]]:
    seen: set[Tuple[str, str]] = set()
    kept: List[Message] = []
    skipped: List[Message] = []
    for i, msg in enumerate(messages, 1):
        email = process_single_message(msg, i, temp_dir, converter)
        warn_empties(email)
        if is_empty_email(email):
            kept.append(msg)
            continue
        fp = email_fingerprint(email)
        if fp in seen:
            log("INFO", f"dedup mbox drop duplicate {q(email.subject)}")
            skipped.append(msg)
            continue
        seen.add(fp)
        kept.append(msg)
    set_log_ctx(None)
    return kept, skipped


def mbox_message_count(path: Path) -> Optional[int]:
    if not path.exists():
        return None
    n = 0
    box = RobustMbox(path)
    try:
        for _ in box:
            n += 1
    finally:
        box.close()
    return n


def log_mbox_family_counts(src: Path) -> None:
    if not debug:
        return
    orig = mbox_message_count(src)
    ded = mbox_message_count(deduplicated_path(src))
    skp = mbox_message_count(skipped_path(src))

    def fmt(n: Optional[int]) -> str:
        return "missing" if n is None else str(n)

    log(
        "DEBUG",
        f"mbox message counts source={q(src)} "
        f"original={fmt(orig)} deduplicated={fmt(ded)} skipped={fmt(skp)}",
    )


def resolve_input_mboxes(
        input_files: List[str], use_dedup: bool, temp_dir: Path, converter: str,
) -> List[Path]:
    resolved: List[Path] = []
    for raw in input_files:
        src = Path(raw)
        log_mbox_family_counts(src)
        log("VERBOSE", f"Input mbox path: {q(src)}")
        if src.name.endswith(DEDUP_POSTFIX) or src.name.endswith(SKIPPED_POSTFIX):
            log("ERROR", f"Refusing input that is a cache mbox: {q(src)}")
            sys.exit(1)
        if not src.exists():
            log("ERROR", f"mbox not found: {q(src)}")
            sys.exit(1)
        if not use_dedup:
            resolved.append(src)
            continue
        dest = deduplicated_path(src)
        skipf = skipped_path(src)
        if mbox_is_fresh(src, dest):
            log("INFO", f"Using cached deduplicated mbox: {q(dest)}")
            resolved.append(dest)
            continue
        log("INFO", f"Building deduplicated mbox from scratch: {q(dest)}")
        loaded = load_mbox_messages(src)
        loaded.sort(key=msg_sort_ts)
        unique, skipped = unique_source_messages(loaded, temp_dir, converter)
        try:
            write_mbox(dest, unique)
            write_mbox(skipf, skipped)
        except OSError as e:
            log("ERROR", f"Cannot write cache mbox: {e}")
            sys.exit(1)
        log("INFO", f"Wrote {len(unique)} keepers to {q(dest)}")
        log("INFO", f"Wrote {len(skipped)} skipped to {q(skipf)}")
        resolved.append(dest)
    return resolved


def render_one_email(
        email: ProcessedEmail, layout: Layout, page_size, tmp: Path,
        use_montage: bool, gs: str, archive: bool, archive_root: Optional[Path],
        font_path: Optional[Path] = None,
) -> Optional[Path]:
    set_log_ctx(email.count, email.date_str, email.subject, email.from_str)
    log("INFO", f"Processing email number {email.count}")

    if archive and archive_root is not None:
        try:
            archive_email(email, archive_root)
        except Exception as e:
            log("ERROR", f"archive failed: {e}")

    piece = tmp / f"email_{email.count}.pdf"
    try:
        c = new_canvas(str(piece), page_size)
        y = write_header(c, layout, email, 1)
        y = write_long_text(c, layout, email, y)
        if not email.body_text or len(email.body_text.strip()) < 1:
            c.setFillColor(HexColor(MUTED_GREY))
            c.setFont(FONT_BOLD, layout.font_empty)
            c.drawCentredString(
                layout.width / 2, y - 100 * (layout.width / PERL_REF_W),
                "this email is without text",
                )
        draw_attachment_summary(c, layout, email, y)
        add_images_to_pdf(
            c, layout, email.images, len(email.body_text or ""),
            tmp, use_montage, font_path,
        )
        if fast_level >= 3:
            for att in email.all_attachments:
                if att.type == "image":
                    continue
                if att.type in {"pdf"} or att.converted_from or ext_of(att.original_name) in CONVERTIBLE:
                    title = f"{placeholder_label(att)}: {att.original_name}"
                    draw_placeholder_page(c, layout, title, placeholder_label(att))
        c.save()
    except Exception as e:
        log("ERROR", f"ReportLab failed: {e}")
        if debug:
            traceback.print_exc()
        return None
    merged_piece = tmp / f"email_{email.count}_merged.pdf"
    try:
        writer_e = new_pdf_writer()
        for page in PdfReader(str(piece)).pages:
            writer_e.add_page(page)
        if fast_level < 3:
            for att in email.pdf_attachments:
                if att.type != "pdf" or not is_pdf_file(att.saved_path):
                    continue
                fitted = tmp / "pdf_attachments" / f"fitted_{email.count}_{att.saved_path.name}"
                if not scale_pdf_to_fit(att.saved_path, fitted, layout.width, layout.height, gs):
                    continue
                try:
                    att_reader = PdfReader(str(fitted))
                    bar = PdfReader(io.BytesIO(make_pdf_bar_page(layout, att.original_name)))
                    first = True
                    for pg in att_reader.pages:
                        if first:
                            try:
                                pg.merge_page(bar.pages[0])
                            except Exception:
                                writer_e.add_page(bar.pages[0])
                                first = False
                                writer_e.add_page(pg)
                                continue
                            first = False
                        writer_e.add_page(pg)
                except Exception as e:
                    log("WARNING", f"insert PDF {q(att.original_name)}: {e}", review=True)
        with open(merged_piece, "wb") as f:
            writer_e.write(f)
    except Exception as e:
        log("ERROR", f"pypdf merge failed: {e}")
        if debug:
            traceback.print_exc()
        return piece if piece.exists() else None
    return merged_piece


def qpdf_check(qpdf: str, path: Path, *, strict: bool = False,
               expect_pages: Optional[int] = None) -> bool:
    args = [qpdf, "--check"]
    if not strict:
        args.append("--warning-exit-0")
    args.append(str(path))
    log("INFO", f"qpdf --check {q(path)}")
    rc, out = run_cmd(args, timeout=QPDF_TIMEOUT)
    snippet = (out or "").strip()
    if snippet:
        log("INFO", f"qpdf --check: {snippet[:800]}")
    if strict:
        if rc != 0 or re.search(r"(?im)^WARNING:", out or ""):
            log("WARNING", f"qpdf --check rejected {q(path)}", review=True)
            return False
        if expect_pages is not None:
            rc2, nout = run_cmd([qpdf, "--show-npages", str(path)])
            try:
                got = int((nout or "").strip().split()[0])
            except Exception:
                got = -1
            if rc2 != 0 or got != expect_pages:
                log("WARNING", f"qpdf page count {got} != {expect_pages}", review=True)
                return False
        return True
    if rc != 0:
        log("WARNING", f"qpdf --check failed: {snippet[:800]}", review=True)
        return False
    return True


def qpdf_finalize_file(qpdf: str, src: Path, dest: Path) -> bool:
    rc, out = run_cmd(
        [qpdf, "--object-streams=generate", "--min-version=1.7",
         "--compress-streams=y", "--recompress-flate", str(src), str(dest)],
        timeout=QPDF_TIMEOUT,
    )
    if rc == 2 or not dest.exists() or dest.stat().st_size <= 0:
        log("WARNING", f"qpdf finalize failed: {(out or '')[:500]}", review=True)
        return False
    try:
        expect = len(PdfReader(str(src)).pages)
    except Exception:
        expect = None
    if not qpdf_check(qpdf, dest, strict=True, expect_pages=expect):
        log("WARNING", f"qpdf output failed strict --check, not promoting {q(dest)}", review=True)
        return False
    return True


def main() -> None:
    global debug, quiet, verbose, prefer_plain_text, email_to_name
    global fast_level, finalize
    global FONT_REGULAR, FONT_BOLD, FONT_ITALIC

    parser = argparse.ArgumentParser(description="Convert mbox archives to a master print PDF")
    parser.add_argument("--base-directory", type=Path, required=True)
    parser.add_argument("--working-directory", type=Path, required=True)
    parser.add_argument("--configuration-file", type=Path, required=True)
    parser.add_argument("--document-converter-binary", type=str, required=True)
    parser.add_argument("--select-year", type=int, default=-1)
    parser.add_argument("--select-emails", type=str, default="0")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument(
        "-f", "--fast", action="count", default=0,
        help="Repeatable: -f dpi=10; -ff image placeholders; -fff also attachment placeholders",
    )
    parser.add_argument(
        "--finalize", action="store_true",
        help="Compress PNG and qpdf object-streams + recompress to <stem>-master.pdf",
    )
    parser.add_argument(
        "--list-emails", dest="list_needle",
        nargs="?", const="", default=None,
        help="List emails as TSV; optional literal substring filter",
    )
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    args = parser.parse_args()
    if args.finalize and args.fast:
        parser.error("--finalize and --fast are mutually exclusive")

    debug, quiet, verbose = args.debug, args.quiet, args.verbose or args.debug
    fast_level = int(args.fast or 0)
    finalize = bool(args.finalize)
    base = args.base_directory.resolve()
    working = args.working_directory
    if not working.is_absolute():
        working = (Path.cwd() / working).resolve()

    config_path = args.configuration_file if args.configuration_file.is_absolute() else working / args.configuration_file
    log("VERBOSE", f"Opening configuration: {q(config_path)}")
    if not config_path.exists():
        log("ERROR", f"Configuration file not found: {q(config_path)}")
        sys.exit(1)
    config = json.loads(config_path.read_text(encoding="utf-8"))

    input_files = config.get("input_files", [])
    if isinstance(input_files, str):
        input_files = [input_files]
    if not input_files:
        log("ERROR", "missing input_files")
        sys.exit(1)

    use_dedup = bool(config.get("use_deduplicated_mbox", False))
    split_big = bool(config.get("split_big_files_on_limit", False))
    output_path = Path(config["output_path"])
    output_filename = config["output_filename"]

    if args.list_needle is not None:
        try:
            input_paths = resolve_input_mboxes_fast(input_files, use_dedup)
            all_messages: List[Message] = []
            for mbox_path in input_paths:
                all_messages.extend(load_mbox_messages(mbox_path))
            all_messages.sort(key=msg_sort_ts)
            print_email_list(all_messages, args.select_year, args.select_emails, args.list_needle)
            return
        except Exception as e:
            rundir = make_rundir(output_path)
            setup_log_files(rundir)
            log("ERROR", f"--list-emails failed: {e}")
            if debug:
                traceback.print_exc()
            link_into_root(output_path, rundir, "convert.log")
            link_into_root(output_path, rundir, "errors.log")
            close_log_files()
            sys.exit(1)

    soffice = args.document_converter_binary
    if not Path(soffice).exists() and not shutil.which(soffice):
        log("ERROR", f"LibreOffice not found: {q(soffice)}")
        sys.exit(1)
    gs = require_tool(["gs", "gswin64c", "gswin32c"], "PDF fit (Ghostscript)")
    qpdf = require_tool(["qpdf"], "PDF check/finalize")
    require_tool(["magick", "convert"], "image raster (ImageMagick)")

    dim = str(config.get("dimensions", "A0")).upper()
    if dim not in DIN_A:
        log("ERROR", f"Unsupported dimensions {q(dim)}")
        sys.exit(1)
    page_size = DIN_A[dim]

    dup_cfg = config.get("duplicate_emails") or {}
    if isinstance(dup_cfg, bool):
        dup_enabled, dup_action = dup_cfg, "warn"
    else:
        dup_enabled = bool(dup_cfg.get("enabled", False))
        dup_action = str(dup_cfg.get("action", "warn")).lower()
    if dup_action not in ("warn", "skip"):
        dup_action = "warn"
    image_dpi = int(config.get("image_dpi", config.get("document_dpi", 300)))
    if fast_level >= 1:
        image_dpi = 10
        log("INFO", f"--fast level {fast_level}: image_dpi=10, uncompressed PNG")
    prefer_plain_text = bool(config.get("prefer_plain_text", True))
    use_montage = str(config.get("use_montages_for_images", True)).lower() in ("1", "true", "yes")
    archive = str(config.get("archive_attachments", False)).lower() in ("1", "true", "yes")
    email_to_name = {k.lower(): v for k, v in (config.get("email_to_name") or {}).items()}

    layout = make_layout(page_size, image_dpi)
    log("INFO", f"Page {dim} {layout.width:.1f}×{layout.height:.1f} pt, image_dpi={image_dpi}")
    if finalize:
        log("INFO", "--finalize: compressed PNG + qpdf object-streams/recompress (no linearize)")

    try:
        pdfmetrics.registerFont(TTFont("OpenSans", str(base / "Build/assets/fonts/OpenSans-Regular.ttf")))
        pdfmetrics.registerFont(TTFont("OpenSans-Bold", str(base / "Build/assets/fonts/OpenSans-Bold.ttf")))
        pdfmetrics.registerFont(TTFont("OpenSans-Italic", str(base / "Build/assets/fonts/OpenSans-LightItalic.ttf")))
        FONT_REGULAR, FONT_BOLD, FONT_ITALIC = "OpenSans", "OpenSans-Bold", "OpenSans-Italic"
    except Exception as e:
        FONT_REGULAR, FONT_BOLD, FONT_ITALIC = "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"
        log("WARNING", f"OpenSans missing, Helvetica: {e}")

    tmp_hash = hashlib.md5("::".join(sorted(map(str, input_files))).encode()).hexdigest()
    tmp = Path(tempfile.gettempdir()) / tmp_hash / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    for folder in ("pdf_attachments", "images", "converted", "montages"):
        (tmp / folder).mkdir(parents=True, exist_ok=True)
    log("VERBOSE", f"Temp directory: {q(tmp)}")

    input_paths = resolve_input_mboxes(input_files, use_dedup, tmp, soffice)
    all_messages: List[Message] = []
    for mbox_path in input_paths:
        all_messages.extend(load_mbox_messages(mbox_path))
    all_messages.sort(key=msg_sort_ts)
    log("INFO", f"Loaded and sorted {len(all_messages)} emails chronologically")

    rundir = make_rundir(output_path)
    setup_log_files(rundir)
    archive_root = prepare_archive_root(rundir) if archive else None

    try:
        clauses = parse_selector(args.select_emails)
    except ValueError as e:
        log("ERROR", str(e))
        sys.exit(1)
    n_all = len(all_messages)
    processed: List[ProcessedEmail] = []
    for i, msg in enumerate(all_messages, 1):
        if args.select_year != -1:
            try:
                if parsedate_to_datetime(msg.get("Date") or "").year != args.select_year:
                    continue
            except Exception:
                continue
        if not index_selected(i, n_all, clauses):
            continue
        processed.append(process_single_message(msg, i, tmp, soffice))

    if dup_enabled:
        before = len(processed)
        processed = filter_duplicate_emails(processed, dup_action)
        log("INFO", f"Duplicate check ({dup_action}): {before} in, {len(processed)} kept")
    else:
        for email in processed:
            set_log_ctx(email.count, email.date_str, email.subject, email.from_str)
            warn_empties(email)
        set_log_ctx(None)

    log("INFO", f"Processing {len(processed)} emails after filtering")

    font_path = base / "Build/assets/fonts/OpenSans-Bold.ttf"

    stem = Path(output_filename).stem
    suffix = Path(output_filename).suffix or ".pdf"
    writer = new_pdf_writer()
    vol_bytes = 0
    vol_idx = 1
    pending: List[Path] = []
    failed = 0

    for email in processed:
        piece = render_one_email(
            email, layout, page_size, tmp, use_montage, gs, archive, archive_root, font_path,
        )
        # if render_one_email still takes output_path only for archive, change that
        # call to pass archive_root — see note below
        if piece is None or not piece.exists():
            failed += 1
            continue
        piece_size = piece.stat().st_size
        if split_big and vol_bytes > 0 and vol_bytes + piece_size >= SPLIT_LIMIT_BYTES:
            path = rundir / f"{stem}-part{vol_idx:02d}{suffix}"
            with open(path, "wb") as f:
                writer.write(f)
            pending.append(path)
            writer = new_pdf_writer()
            vol_bytes = 0
            vol_idx += 1
        try:
            for page in PdfReader(str(piece)).pages:
                writer.add_page(page)
            vol_bytes += piece_size
        except Exception as e:
            failed += 1
            log("ERROR", f"could not append rendered pages: {e}")

    if split_big and vol_idx > 1:
        if len(writer.pages):
            path = rundir / f"{stem}-part{vol_idx:02d}{suffix}"
            with open(path, "wb") as f:
                writer.write(f)
            pending.append(path)
    else:
        path = rundir / output_filename
        with open(path, "wb") as f:
            writer.write(f)
        pending = [path]

    if split_big and len(pending) == 1 and pending[0].name != output_filename:
        single = rundir / output_filename
        pending[0].rename(single)
        pending = [single]

    log("INFO", f"Assembled volumes={len(pending)} ({failed} emails failed)")

    masters: List[Path] = []
    if finalize:
        for assembled in pending:
            qpdf_tmp = tmp / f"finalized_{assembled.name}"
            if qpdf_finalize_file(qpdf, assembled, qpdf_tmp):
                master = assembled.with_name(f"{assembled.stem}-master{assembled.suffix}")
                shutil.copy2(qpdf_tmp, master)
                masters.append(master)
                log("INFO", f"Wrote finalized PDF: {q(master)}")
            else:
                log("WARNING", f"qpdf finalize not promoted; left {q(assembled)}", review=True)
    else:
        for assembled in pending:
            qpdf_check(qpdf, assembled)

    link_into_root(output_path, rundir, "convert.log")
    link_into_root(output_path, rundir, "errors.log")
    if archive_root is not None:
        link_into_root(output_path, rundir, "attachments")
    keep = {p.name for p in pending + masters}
    for p in pending + masters:
        link_into_root(output_path, rundir, p.name)
    cleanup_stale_part_symlinks(output_path, stem, suffix, keep)

    log("INFO", f"Please delete the temporary directory when you no longer need the intermediate files: {q(tmp)}")

    close_log_files()

if __name__ == "__main__":
    try:
        main()
    finally:
        close_log_files()
