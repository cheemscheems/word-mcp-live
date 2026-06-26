"""
Word OOXML document protection via <w:documentProtection> in settings.xml.

Implements the standard Word password hash algorithm (salted, iterated SHA-512)
and the w:documentProtection element that Word's Restrict Editing pane uses.

Word 2013+ default: SHA-512, 100000 iterations, 16-byte salt.
python-docx manages word/settings.xml natively, so no zip-level manipulation
is needed -- doc.settings.element provides direct access to the XML tree.
"""

import base64
import hashlib
import secrets

from docx.oxml.shared import OxmlElement, qn


# -- Edit modes supported by w:documentProtection ----------------------------
EDIT_MODES = {"readOnly", "comments", "trackedChanges", "forms"}


def _compute_word_hash(password: str, salt: bytes, spin_count: int = 100000) -> bytes:
    """Compute the Word documentProtection password hash.

    The algorithm matches Word 2013+ behaviour:
    1. UTF-16LE encode the password plus a null terminator (Word appends
       the null internally).
    2. H₀ = SHA-512(password_bytes ∥ salt)
    3. For i in 1..spinCount: Hᵢ = SHA-512(password_bytes ∥ Hᵢ₋₁)
    4. Return final H (base64-encoded by caller).

    Args:
        password:  The protection password (plain text).
        salt:      16-byte cryptographically random salt.
        spin_count: Iteration count (Word 2013+ default is 100000).

    Returns:
        Raw 64-byte SHA-512 digest.
    """
    password_bytes = password.encode("utf-16-le") + b"\x00\x00"
    h = hashlib.sha512(password_bytes + salt).digest()
    for _ in range(spin_count):
        h = hashlib.sha512(password_bytes + h).digest()
    return h


def add_restricted_editing_protection(doc, edit_mode: str, password: str,
                                      spin_count: int = 100000) -> None:
    """Add real Word restricted-editing protection to a document.

    Inserts a ``<w:documentProtection>`` element into ``word/settings.xml``
    that Word's "Restrict Editing" pane (Review → Restrict Editing) will
    honour when the document is opened.

    Args:
        doc:        python-docx ``Document`` instance (must already be open).
        edit_mode:  One of ``"readOnly"``, ``"comments"``,
                    ``"trackedChanges"``, or ``"forms"``.
        password:   The password required to change or remove protection.
        spin_count: Hash iteration count (default 100000, Word 2013+ default).

    Raises:
        ValueError: If *edit_mode* is not a recognised mode.
    """
    if edit_mode not in EDIT_MODES:
        raise ValueError(
            f"Unknown edit_mode '{edit_mode}'. "
            f"Choose from: {', '.join(sorted(EDIT_MODES))}"
        )

    salt = secrets.token_bytes(16)
    hash_bytes = _compute_word_hash(password, salt, spin_count)
    hash_b64 = base64.b64encode(hash_bytes).decode("ascii")
    salt_b64 = base64.b64encode(salt).decode("ascii")

    settings = doc.settings.element

    # Remove any existing documentProtection element first
    for child in list(settings):
        if child.tag == qn("w:documentProtection"):
            settings.remove(child)

    prot = OxmlElement("w:documentProtection")
    prot.set(qn("w:edit"), edit_mode)
    prot.set(qn("w:enforcement"), "1")
    prot.set(qn("w:cryptProviderType"), "rsaAES")
    prot.set(qn("w:cryptAlgorithmClass"), "hash")
    prot.set(qn("w:cryptAlgorithmType"), "typeAny")
    prot.set(qn("w:cryptAlgorithmSid"), "14")          # SHA-512
    prot.set(qn("w:cryptSpinCount"), str(spin_count))
    prot.set(qn("w:hash"), hash_b64)
    prot.set(qn("w:salt"), salt_b64)
    settings.append(prot)


def remove_restricted_editing_protection(doc) -> bool:
    """Remove ``<w:documentProtection>`` from a document's settings.xml.

    Args:
        doc: python-docx ``Document`` instance.

    Returns:
        ``True`` if protection was found and removed, ``False`` if the
        element was not present.
    """
    settings = doc.settings.element
    for child in list(settings):
        if child.tag == qn("w:documentProtection"):
            settings.remove(child)
            return True
    return False


def has_document_protection(doc) -> tuple:
    """Check whether the document has active restricted-editing protection.

    Args:
        doc: python-docx ``Document`` instance.

    Returns:
        ``(is_protected: bool, edit_mode: str | None)``.
    """
    settings = doc.settings.element
    for child in settings:
        if child.tag == qn("w:documentProtection"):
            mode = child.get(qn("w:edit"), "readOnly")
            return True, mode
    return False, None
