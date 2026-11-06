"""Outlook signature discovery and loading.

Classic Outlook stores each signature as files in the user's signatures
folder (``%APPDATA%\\Microsoft\\Signatures`` — the last segment is
localized on some installs, e.g. ``Signaturen``):

* ``<Name>.htm`` — the HTML version we insert
* ``<Name>_files\\`` — images referenced by the .htm (relative ``src``)

Which signature is the *default* for new messages and for
replies/forwards is stored per account in the profile registry:

  HKCU\\Software\\Microsoft\\Office\\<ver>\\Outlook\\Profiles\\<profile>\\
      9375CFF0413111d3B88A00104B2A6676\\<0000000N>
        "Account Name"             (UTF-16 binary or string)
        "New Signature"            (UTF-16 binary or string)
        "Reply-Forward Signature"  (UTF-16 binary or string)

with an older, profile-wide fallback under
``HKCU\\Software\\Microsoft\\Office\\<ver>\\Common\\MailSettings``
(``NewSignature`` / ``ReplySignature``).

Inserting the .htm ourselves (rather than relying on Outlook creating an
inspector) works the same way for new mail, replies and forwards, and
lets the caller pick a signature by name. Images are attached as hidden
inline attachments and the ``src`` rewritten to ``cid:`` references.

The parsing helpers are pure and platform-independent; only the
directory/registry lookups need Windows.
"""

from __future__ import annotations

import os
import re
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path

from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.html_compose import body_inner_html

_OFFICE_VERSIONS = ("16.0", "15.0")
_ACCOUNTS_SUBKEY = "9375CFF0413111d3B88A00104B2A6676"

_CHARSET_RE = re.compile(rb"charset\s*=\s*[\"']?([A-Za-z0-9_\-]+)", re.IGNORECASE)
# src="..." / src='...' / src=bare  (also matches VML <v:imagedata src=...>)
_SRC_RE = re.compile(
    r"""(?P<attr>\bsrc\s*=\s*)(?P<q>["']?)(?P<url>[^"'\s>]+)(?P=q)""",
    re.IGNORECASE,
)


@dataclass
class Signature:
    name: str
    html: str  # inner body markup, image src already rewritten to cid:
    images: list[tuple[str, str]] = field(default_factory=list)  # (abs path, content-id)


# --------------------------------------------------------------------- pure


def decode_reg_string(value: object) -> str:
    """Decode a signature-name registry value (REG_BINARY UTF-16 or REG_SZ)."""
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray)):
        raw = bytes(value)
        try:
            text = raw.decode("utf-16-le")
        except UnicodeDecodeError:
            text = raw.decode("latin-1")
        return text.split("\x00", 1)[0].strip()
    return str(value).split("\x00", 1)[0].strip()


def decode_htm(raw: bytes) -> str:
    """Decode a signature .htm using its declared charset.

    Outlook writes signatures as windows-1252 (or another legacy code
    page) with a ``<meta charset>``; newer builds may write UTF-8.
    """
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace")
    m = _CHARSET_RE.search(raw[:4096])
    charset = m.group(1).decode("ascii", "ignore") if m else "windows-1252"
    try:
        return raw.decode(charset, errors="replace")
    except LookupError:
        return raw.decode("windows-1252", errors="replace")


def _content_id(index: int, filename: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]", "_", filename) or "image"
    return f"sig{index}-{stem}@outlook-mcp"


def prepare_signature_html(
    html: str, base_dir: Path, name: str = ""
) -> Signature:
    """Extract the body of a signature document and inline its images.

    Relative ``src`` values that resolve to an existing file under
    ``base_dir`` are rewritten to ``cid:`` references; the returned
    ``images`` list says which files to attach under which content-id.
    Absolute URLs (``http:``, ``https:``, ``cid:``, ``data:``) are left
    alone, as are references to files that don't exist.
    """
    inner = body_inner_html(html)
    images: list[tuple[str, str]] = []
    cid_for_path: dict[str, str] = {}
    base_resolved = base_dir.resolve()

    def _rewrite(m: re.Match[str]) -> str:
        url = m.group("url")
        lowered = url.lower()
        if re.match(r"^[a-z][a-z0-9+.\-]*:", lowered) or lowered.startswith("//"):
            return m.group(0)
        rel = urllib.parse.unquote(url).replace("\\", "/")
        candidate = (base_dir / rel).resolve()
        try:
            candidate.relative_to(base_resolved)
        except ValueError:
            return m.group(0)  # escapes the signatures folder: leave it
        if not candidate.is_file():
            return m.group(0)
        key = str(candidate)
        cid = cid_for_path.get(key)
        if cid is None:
            cid = _content_id(len(images) + 1, candidate.name)
            cid_for_path[key] = cid
            images.append((key, cid))
        q = m.group("q") or '"'
        return f"{m.group('attr')}{q}cid:{cid}{q}"

    inner = _SRC_RE.sub(_rewrite, inner)
    return Signature(name=name, html=inner.strip(), images=images)


# ------------------------------------------------------------- filesystem


def signatures_dir() -> Path:
    """Locate the user's Outlook signatures folder.

    ``OUTLOOK_MCP_SIGNATURES_DIR`` overrides the lookup.
    """
    override = os.environ.get("OUTLOOK_MCP_SIGNATURES_DIR", "").strip()
    if override:
        return Path(override)
    appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    base = Path(appdata) / "Microsoft"
    localized = _registry_signatures_folder_name()
    if localized and (base / localized).is_dir():
        return base / localized
    return base / "Signatures"


def list_signatures() -> list[str]:
    d = signatures_dir()
    if not d.is_dir():
        return []
    return sorted(p.stem for p in d.glob("*.htm"))


def load_signature(name: str) -> Signature:
    """Load ``<name>.htm`` from the signatures folder (case-insensitive)."""
    d = signatures_dir()
    if not d.is_dir():
        raise OutlookError(
            f"Outlook signatures folder not found ({d}). Create a signature in "
            "Outlook (File > Options > Mail > Signatures) or set "
            "OUTLOOK_MCP_SIGNATURES_DIR."
        )
    wanted = name.strip().lower()
    for p in d.glob("*.htm"):
        if p.stem.lower() == wanted:
            return prepare_signature_html(decode_htm(p.read_bytes()), d, name=p.stem)
    available = ", ".join(list_signatures()) or "none"
    raise OutlookError(
        f"Signature '{name}' not found. Available signatures: {available}. "
        "Use outlook_list_signatures to see them."
    )


# --------------------------------------------------------------- registry


def _winreg():
    try:
        import winreg  # type: ignore[import-not-found]
    except ImportError:  # not on Windows
        return None
    return winreg


def _registry_signatures_folder_name() -> str | None:
    winreg = _winreg()
    if winreg is None:
        return None
    for ver in _OFFICE_VERSIONS:
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, rf"Software\Microsoft\Office\{ver}\Common\General"
            ) as key:
                value, _ = winreg.QueryValueEx(key, "Signatures")
                text = decode_reg_string(value)
                if text:
                    return text
        except OSError:
            continue
    return None


def _query(winreg, key, name: str) -> str:
    try:
        value, _ = winreg.QueryValueEx(key, name)
    except OSError:
        return ""
    return decode_reg_string(value)


def default_signature_name(
    kind: str, profile: str | None, account_names: list[str]
) -> str | None:
    """Return the default signature name for ``kind`` ('new' or 'reply').

    ``account_names`` are the identifiers of the sending account (display
    name and SMTP address); the first registry account whose
    "Account Name" matches wins. Falls back to the first account that has
    a signature configured, then to the profile-wide MailSettings value.
    """
    winreg = _winreg()
    if winreg is None:
        return None
    value_name = "New Signature" if kind == "new" else "Reply-Forward Signature"
    wanted = {n.strip().lower() for n in account_names if n}
    first_configured: str | None = None

    for ver in _OFFICE_VERSIONS:
        if profile:
            path = rf"Software\Microsoft\Office\{ver}\Outlook\Profiles\{profile}\{_ACCOUNTS_SUBKEY}"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as root:
                    i = 0
                    while True:
                        try:
                            sub = winreg.EnumKey(root, i)
                        except OSError:
                            break
                        i += 1
                        try:
                            with winreg.OpenKey(root, sub) as acc_key:
                                sig = _query(winreg, acc_key, value_name)
                                if not sig:
                                    continue
                                acc_name = _query(winreg, acc_key, "Account Name").lower()
                                if acc_name and acc_name in wanted:
                                    return sig
                                if first_configured is None:
                                    first_configured = sig
                        except OSError:
                            continue
            except OSError:
                pass
        if first_configured:
            return first_configured

        legacy = "NewSignature" if kind == "new" else "ReplySignature"
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, rf"Software\Microsoft\Office\{ver}\Common\MailSettings"
            ) as key:
                sig = _query(winreg, key, legacy)
                if sig:
                    return sig
        except OSError:
            pass
    return None
