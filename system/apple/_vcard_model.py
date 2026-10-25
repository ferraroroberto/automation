"""
vCard data model and parser used by the unification tool.

Holds the contact dataclasses and ``VCardParser``, which turns a ``.vcf`` file
into ``Contact`` objects; the duplicate detection, merge and write steps live
in ``unify_vcards.py``.
"""

import re
import logging
from typing import List, Optional
from dataclasses import dataclass, field

from _vcard_fields import extract_type_label

logger = logging.getLogger(__name__)


def longest_common_substring(a: str, b: str) -> str:
    """Longest common substring of a and b, compared case-insensitively.

    The result is sliced from ``a`` (so it keeps a's casing); on ties the
    earliest-ending match in ``a`` wins. Returns "" when nothing is shared.
    """
    a_lower = a.lower()
    b_lower = b.lower()
    len1, len2 = len(a), len(b)
    dp = [[0] * (len2 + 1) for _ in range(len1 + 1)]

    longest_len = 0
    end_pos = 0

    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            if a_lower[i-1] == b_lower[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                if dp[i][j] > longest_len:
                    longest_len = dp[i][j]
                    end_pos = i
            else:
                dp[i][j] = 0

    return a[end_pos - longest_len:end_pos]


@dataclass
class ContactPhone:
    """Represents a phone number with its type and formatted value."""
    number: str = ""
    type_label: str = "CELL"
    original_line: str = ""

    def __str__(self) -> str:
        return f"{self.type_label}: {self.number}"


@dataclass
class ContactEmail:
    """Represents an email address with its type."""
    address: str
    type_label: str = "INTERNET"
    original_line: str = ""

    def __str__(self) -> str:
        return f"{self.type_label}: {self.address}"


@dataclass
class ContactAddress:
    """Represents a physical address with its type."""
    street: str = ""
    city: str = ""
    region: str = ""
    postal_code: str = ""
    country: str = ""
    type_label: str = "HOME"
    original_line: str = ""

    def __str__(self) -> str:
        parts = [self.street, self.city, self.region, self.postal_code, self.country]
        return ", ".join([p for p in parts if p])


@dataclass
class Contact:
    """Represents a complete vCard contact."""
    vcard_id: int
    full_name: str = ""
    first_name: str = ""
    last_name: str = ""
    phones: List[ContactPhone] = field(default_factory=list)
    emails: List[ContactEmail] = field(default_factory=list)
    addresses: List[ContactAddress] = field(default_factory=list)
    organization: str = ""
    title: str = ""
    notes: str = ""
    raw_vcard: str = ""
    original_lines: List[str] = field(default_factory=list)
    # Properties the parser does not model (PHOTO, BDAY, URL, CATEGORIES, X-*, ...),
    # one entry per property with folded continuation lines joined by a newline.
    extra_lines: List[str] = field(default_factory=list)

    def get_display_name(self) -> str:
        """Get the best display name for this contact."""
        if self.full_name:
            return self.full_name
        if self.first_name or self.last_name:
            return f"{self.first_name} {self.last_name}".strip()
        return f"Contact {self.vcard_id}"

    def get_all_phone_numbers(self) -> List[str]:
        """Get all phone numbers as strings."""
        return [phone.number for phone in self.phones]

    def has_shared_phone(self, other: 'Contact') -> bool:
        """Check if this contact shares any phone numbers with another."""
        phones1 = set(self.get_all_phone_numbers())
        phones2 = set(other.get_all_phone_numbers())
        return bool(phones1.intersection(phones2))


class VCardParser:
    """Parser for vCard files to extract structured contact information."""

    def __init__(self):
        """Initialize the vCard parser."""
        self.contact_counter = 0

    def parse_file(self, file_path: str) -> List[Contact]:
        """Parse a vCard file and return a list of Contact objects.

        Args:
            file_path: Path to the vCard file

        Returns:
            List of parsed Contact objects
        """
        logger.info(f"📂 Reading vCard file: {file_path}")

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            logger.error(f"❌ Failed to read file {file_path}: {e}")
            return []

        contacts = self.parse_content(content)
        logger.info(f"✅ Parsed {len(contacts)} contacts from file")
        return contacts

    def parse_content(self, content: str) -> List[Contact]:
        """Parse vCard content and extract contacts.

        Args:
            content: Raw vCard content

        Returns:
            List of parsed Contact objects
        """
        # Split content into individual vCard entries
        vcard_blocks = content.split("BEGIN:VCARD")
        contacts = []

        for block in vcard_blocks:
            if not block.strip():
                continue

            # Add BEGIN:VCARD back
            vcard_content = "BEGIN:VCARD" + block

            contact = self._parse_single_vcard(vcard_content)
            if contact:
                contacts.append(contact)

        return contacts

    def _parse_single_vcard(self, vcard_content: str) -> Optional[Contact]:
        """Parse a single vCard entry.

        Args:
            vcard_content: Single vCard entry content

        Returns:
            Parsed Contact object or None if parsing fails
        """
        self.contact_counter += 1
        contact = Contact(vcard_id=self.contact_counter)
        contact.raw_vcard = vcard_content
        contact.original_lines = vcard_content.split('\n')
        contact.extra_lines = self._collect_extra_lines(vcard_content)

        lines = vcard_content.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Parse different field types
            if line.startswith("FN:"):
                contact.full_name = line[3:].strip()
            elif line.startswith("N:"):
                self._parse_name_field(line, contact)
            elif line.startswith("TEL"):
                self._parse_phone_field(line, contact)
            elif line.startswith("EMAIL"):
                self._parse_email_field(line, contact)
            elif line.startswith("ADR"):
                self._parse_address_field(line, contact)
            elif line.startswith("ORG:"):
                contact.organization = line[4:].strip()
            elif line.startswith("TITLE:"):
                contact.title = line[6:].strip()
            elif line.startswith("NOTE:"):
                contact.notes = line[5:].strip()

        # If no full name but we have first/last, create full name
        if not contact.full_name and (contact.first_name or contact.last_name):
            contact.full_name = f"{contact.first_name} {contact.last_name}".strip()

        return contact

    # Property prefixes _parse_single_vcard models (plus the card framing); every
    # other property is carried through verbatim in Contact.extra_lines.
    _HANDLED_PREFIXES = ("BEGIN:", "END:", "VERSION", "FN:", "N:", "TEL", "EMAIL",
                         "ADR", "ORG:", "TITLE:", "NOTE:")

    def _collect_extra_lines(self, vcard_content: str) -> List[str]:
        """Collect the properties this parser does not model, folded lines included."""
        extras: List[str] = []
        in_extra = False
        for raw_line in vcard_content.split('\n'):
            line = raw_line.rstrip('\r')
            if not line.strip():
                continue
            if line[0] in ' \t':
                # Folded continuation of the previous property
                if in_extra:
                    extras[-1] += "\n" + line
                continue
            in_extra = not line.startswith(self._HANDLED_PREFIXES)
            if in_extra:
                extras.append(line)
        return extras

    def _parse_name_field(self, line: str, contact: Contact):
        """Parse the N (name) field.

        Args:
            line: N field line
            contact: Contact object to update
        """
        if ":" not in line:
            return

        name_parts = line.split(":", 1)[1].split(";")
        if len(name_parts) >= 2:
            contact.last_name = name_parts[0].strip()
            contact.first_name = name_parts[1].strip()

    def _parse_phone_field(self, line: str, contact: Contact):
        """Parse a phone field.

        Args:
            line: Phone field line
            contact: Contact object to update
        """
        # Extract phone type and number
        if ":" in line:
            prefix, number = line.split(":", 1)
            phone_number = self._normalize_phone_number(number.strip())

            phone = ContactPhone(number=phone_number, original_line=line)

            # Extract type from prefix
            type_label = extract_type_label(prefix)
            if type_label is not None:
                phone.type_label = type_label

            contact.phones.append(phone)

    def _parse_email_field(self, line: str, contact: Contact):
        """Parse an email field.

        Args:
            line: Email field line
            contact: Contact object to update
        """
        if ":" in line:
            prefix, address = line.split(":", 1)
            email_addr = address.strip()

            email = ContactEmail(address=email_addr, original_line=line)

            # Extract type from prefix (no legacy-positional fallback for
            # email, matching the original parsing - only TEL had that)
            type_label = extract_type_label(prefix, legacy_positional_fallback=False)
            if type_label is not None:
                email.type_label = type_label

            contact.emails.append(email)

    def _parse_address_field(self, line: str, contact: Contact):
        """Parse an address field.

        Args:
            line: Address field line
            contact: Contact object to update
        """
        address = ContactAddress(original_line=line)

        if ":" in line:
            prefix, addr_data = line.split(":", 1)
            addr_parts = addr_data.split(";")

            # vCard address format: post-office-box;ext-address;street-address;locality;region;postal-code;country
            if len(addr_parts) >= 7:
                address.street = addr_parts[2].strip()  # street address
                address.city = addr_parts[3].strip()    # locality
                address.region = addr_parts[4].strip()  # region
                address.postal_code = addr_parts[5].strip()  # postal code
                address.country = addr_parts[6].strip()  # country

            # Extract type from prefix (no legacy-positional fallback for
            # address, matching the original parsing - only TEL had that)
            type_label = extract_type_label(prefix, legacy_positional_fallback=False)
            if type_label is not None:
                address.type_label = type_label

        contact.addresses.append(address)

    def _normalize_phone_number(self, phone: str) -> str:
        """Normalize phone number format.

        Args:
            phone: Raw phone number

        Returns:
            Normalized phone number
        """
        phone = phone.strip()

        # Remove all non-digit characters except + and spaces
        phone = re.sub(r'[^\d\+\s]', '', phone)

        # Handle common international prefixes
        if phone.startswith("0039"):
            return "+39" + phone[4:]
        elif phone.startswith("0034"):
            return "+34" + phone[4:]
        elif phone.startswith("00"):
            return "+" + phone[2:]

        return phone
