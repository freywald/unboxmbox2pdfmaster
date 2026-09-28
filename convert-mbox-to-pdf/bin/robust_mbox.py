#!/usr/bin/env python3
"""
robust_mbox.py — mbox I/O that does not require a 7-bit From-line.

Read:
	mbox = RobustMbox(path)
	for msg in mbox:
		raw = msg.as_bytes()
	mbox.close()

Write:
	write_mbox(path, messages)

Envelope lines may be UTF-8. MIME is left to email.parser.BytesParser.
"""

from __future__ import annotations

import io
import os
import time
from email import policy
from email.generator import BytesGenerator
from email.message import EmailMessage, Message
from email.parser import BytesParser
from pathlib import Path
from typing import Iterable, Iterator, List, Optional, Tuple, Union

PathLike = Union[str, Path]
FROM_PREFIX = b"From "


def _open_rb(path: Path):
	return open(path, "rb")


def _scan_toc(fh) -> List[Tuple[int, int]]:
	"""Byte offsets (start, stop) for each message. Same idea as mailbox._generate_toc."""
	starts: List[int] = []
	stops: List[int] = []
	last_was_empty = False
	fh.seek(0)
	while True:
		line_pos = fh.tell()
		line = fh.readline()
		if line.startswith(FROM_PREFIX):
			if len(stops) < len(starts):
				stops.append(line_pos - 1 if last_was_empty and line_pos > 0 else line_pos)
			starts.append(line_pos)
			last_was_empty = False
		elif not line:
			end = line_pos
			if last_was_empty and end > 0:
				end = max(starts[-1] + 1, end - 1) if starts else end
			if len(stops) < len(starts):
				stops.append(line_pos)
			break
		elif line in (b"\n", b"\r\n", b"\r"):
			last_was_empty = True
		else:
			last_was_empty = False
	return list(zip(starts, stops))


def _decode_envelope(from_line: bytes) -> str:
	line = from_line.rstrip(b"\r\n")
	if line.startswith(FROM_PREFIX):
		line = line[len(FROM_PREFIX):]
	return line.decode("utf-8", errors="replace")


def _stub_message(raw: bytes, reason: str, envelope: str) -> Message:
	msg = EmailMessage()
	msg["Subject"] = f"(unparseable message: {reason})"
	msg["X-Robust-Mbox-Error"] = reason
	if envelope:
		msg["X-Robust-Mbox-From-Line"] = envelope[:200]
	try:
		preview = raw.decode("utf-8", errors="replace")
	except Exception:
		preview = raw[:4000].decode("latin-1", errors="replace")
	msg.set_content(preview[:20000])
	return msg


def _parse_slice(blob: bytes) -> Tuple[Message, str]:
	"""Return (message, envelope_text). First line is the From_ envelope if present."""
	if blob.startswith(b"\r\n"):
		blob = blob[2:]
	elif blob.startswith(b"\n"):
		blob = blob[1:]

	envelope = ""
	rest = blob
	nl = blob.find(b"\n")
	first = blob[: nl + 1] if nl >= 0 else blob
	if first.lstrip(b"\r").startswith(FROM_PREFIX) or first.startswith(FROM_PREFIX):
		envelope = _decode_envelope(first)
		rest = blob[len(first):] if nl >= 0 else b""

	rest = rest.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
	try:
		msg = BytesParser(policy=policy.default).parsebytes(rest)
	except Exception as e:
		return _stub_message(rest, str(e), envelope), envelope

	if envelope:
		try:
			msg.set_unixfrom("From " + envelope)
		except Exception:
			pass
	return msg, envelope


class RobustMbox:
	"""Minimal mailbox.mbox-compatible reader."""

	def __init__(self, path: PathLike, factory=None, create: bool = False):
		self.path = Path(path)
		self._factory = factory
		self._fh = None
		self._toc: List[Tuple[int, int]] = []
		if not self.path.exists():
			if create:
				self.path.touch()
			else:
				raise FileNotFoundError(self.path)
		self._fh = _open_rb(self.path)
		self._toc = _scan_toc(self._fh)

	def __iter__(self) -> Iterator[Message]:
		for key in range(len(self._toc)):
			yield self.get_message(key)

	def __len__(self) -> int:
		return len(self._toc)

	def keys(self):
		return range(len(self._toc))

	def get_bytes(self, key: int) -> bytes:
		if self._fh is None:
			raise ValueError("mbox is closed")
		start, stop = self._toc[key]
		self._fh.seek(start)
		return self._fh.read(stop - start)

	def get_message(self, key: int) -> Message:
		blob = self.get_bytes(key)
		msg, envelope = _parse_slice(blob)
		if self._factory is not None:
			try:
				return self._factory(io.BytesIO(blob))
			except Exception:
				pass
		msg._robust_envelope = envelope  # type: ignore[attr-defined]
		msg._robust_raw = blob  # type: ignore[attr-defined]
		return msg

	def close(self) -> None:
		if self._fh is not None:
			self._fh.close()
			self._fh = None

	def __enter__(self) -> "RobustMbox":
		return self

	def __exit__(self, *exc) -> None:
		self.close()


def _envelope_bytes(msg: Message) -> bytes:
	raw = getattr(msg, "_robust_raw", None)
	if isinstance(raw, bytes) and raw.startswith(FROM_PREFIX):
		nl = raw.find(b"\n")
		if nl != -1:
			return raw[:nl].rstrip(b"\r")

	unix = None
	try:
		unix = msg.get_unixfrom()
	except Exception:
		unix = None
	if unix:
		line = unix if unix.startswith("From ") else "From " + unix
		return line.encode("utf-8", errors="replace")

	env = getattr(msg, "_robust_envelope", None)
	if env:
		return (FROM_PREFIX + str(env).encode("utf-8", errors="replace"))

	stamp = time.asctime(time.gmtime())
	return f"From MAILER-DAEMON {stamp}".encode("ascii")


def _flatten_message(msg: Message) -> bytes:
	raw = getattr(msg, "_robust_raw", None)
	if isinstance(raw, bytes) and len(raw) > 5:
		nl = raw.find(b"\n")
		body = raw[nl + 1:] if raw.startswith(FROM_PREFIX) and nl >= 0 else raw
		return body.replace(b"\r\n", b"\n").replace(b"\r", b"\n")

	buf = io.BytesIO()
	BytesGenerator(buf, mangle_from_=False, maxheaderlen=0).flatten(msg)
	return buf.getvalue().replace(b"\r\n", b"\n").replace(b"\r", b"\n")


def _mangle_from(body: bytes) -> bytes:
	if not body:
		return body
	parts = body.split(b"\n")
	out = []
	for i, line in enumerate(parts):
		if line.startswith(FROM_PREFIX):
			out.append(b">" + line)
		else:
			out.append(line)
	data = b"\n".join(out)
	if not data.endswith(b"\n"):
		data += b"\n"
	return data


def write_mbox(path: PathLike, messages: Iterable[Message]) -> Path:
	"""
	Atomic write. Envelope is UTF-8. Body lines starting with 'From '
	are written as '>From ' so the file stays a valid mbox.
	"""
	dest = Path(path)
	tmp = dest.with_name(dest.name + ".tmp")
	if tmp.exists():
		tmp.unlink()
	with open(tmp, "wb") as fh:
		for msg in messages:
			fh.write(_envelope_bytes(msg))
			fh.write(b"\n")
			fh.write(_mangle_from(_flatten_message(msg)))
			if not fh.tell() or True:
				# blank line between messages (classic mbox)
				pass
		fh.flush()
		os.fsync(fh.fileno())
	tmp.replace(dest)
	return dest
