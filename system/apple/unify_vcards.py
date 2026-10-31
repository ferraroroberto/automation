#!/usr/bin/env python3
"""
vCard Unification Tool

This module provides an interactive tool to unify duplicate contacts in vCard files.
It analyzes contacts for potential duplicates based on shared phone numbers and common name words,
then guides users through an interactive process to merge or keep contacts.
"""

import os
import re
import logging
from pathlib import Path
from typing import List, Optional, Tuple, Set
import datetime
import tkinter as tk
from tkinter import filedialog

from _vcard_fields import has_type_param
from _vcard_model import Contact, ContactAddress, VCardParser, longest_common_substring

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DuplicateDetector:
    """Detects potential duplicate contacts based on various criteria."""

    def __init__(self):
        """Initialize the duplicate detector."""
        self.processed_pairs: Set[Tuple[int, int]] = set()

    def find_duplicates(self, contacts: List[Contact]) -> List[Tuple[Contact, Contact, str]]:
        """Find potential duplicate contacts.

        Args:
            contacts: List of contacts to analyze

        Returns:
            List of tuples (contact1, contact2, reason) for potential duplicates
        """
        logger.info("🔍 Analyzing contacts for duplicates...")
        duplicates = []

        for i, contact1 in enumerate(contacts):
            for j, contact2 in enumerate(contacts):
                if i >= j:  # Avoid self-comparison and duplicate pairs
                    continue

                pair_key = tuple(sorted([contact1.vcard_id, contact2.vcard_id]))
                if pair_key in self.processed_pairs:
                    continue

                self.processed_pairs.add(pair_key)

                # Check for duplicates
                reason = self._is_duplicate(contact1, contact2)
                if reason:
                    duplicates.append((contact1, contact2, reason))

        logger.info(f"⚠️  Found {len(duplicates)} potential duplicate pairs")
        return duplicates

    def _is_duplicate(self, contact1: Contact, contact2: Contact) -> Optional[str]:
        """Check if two contacts are duplicates.

        Args:
            contact1: First contact
            contact2: Second contact

        Returns:
            Reason string if duplicate, None otherwise
        """
        # Check for shared phone numbers
        if contact1.has_shared_phone(contact2):
            return "Shared phone number"

        # Check for common words in names (case insensitive)
        name1 = contact1.get_display_name().lower()
        name2 = contact2.get_display_name().lower()

        # Split names into words and filter out non-name words
        words1 = self._filter_name_words(name1.split())
        words2 = self._filter_name_words(name2.split())

        common_words = words1.intersection(words2)

        if common_words:
            return f"Common name words: {', '.join(sorted(common_words))}"

        return None

    def _filter_name_words(self, words: List[str]) -> Set[str]:
        """Filter out non-name words from a list of words.

        Args:
            words: List of words from a name

        Returns:
            Set of filtered words that are likely to be name components
        """
        # Words to exclude (phone types, locations, etc.)
        exclude_words = {
            'cell', 'cellular', 'mobile', 'phone', 'tel', 'telephone',
            'home', 'work', 'office', 'main', 'fax', 'pager',
            'casa', 'house', 'apartment', 'apt', 'suite', 'room',
            'street', 'st', 'avenue', 'ave', 'road', 'rd', 'drive', 'dr',
            'boulevard', 'blvd', 'lane', 'ln', 'way', 'place', 'pl',
            'city', 'town', 'village', 'county', 'state', 'country',
            'email', 'mail', 'internet', 'web',
            'organization', 'org', 'company', 'corp', 'inc', 'ltd',
            'department', 'dept', 'division', 'div', 'group', 'team',
            'title', 'position', 'job', 'role',
            'contact', 'person', 'individual', 'user',
            'and', 'or', 'the', 'a', 'an', 'of', 'for', 'to', 'from', 'by', 'with', 'in', 'on', 'at'
        }

        filtered_words = set()
        for word in words:
            word = word.strip()
            # Skip very short words (likely not meaningful name components)
            if len(word) <= 2:
                continue
            # Skip excluded words
            if word.lower() in exclude_words:
                continue
            # Skip words that are clearly not names (numbers, special chars)
            if not word.replace('-', '').replace("'", '').isalpha():
                continue
            filtered_words.add(word.lower())

        return filtered_words


class ContactMerger:
    """Handles intelligent merging of contact information."""

    def __init__(self):
        """Initialize the contact merger."""
        pass

    def merge_contacts(self, contacts: List[Contact], new_name: str = "") -> Contact:
        """Merge multiple contacts into one.

        Args:
            contacts: List of contacts to merge
            new_name: Optional new name for the merged contact

        Returns:
            Merged Contact object
        """
        if not contacts:
            raise ValueError("Cannot merge empty contact list")

        # Use the first contact as base
        merged = Contact(vcard_id=min(c.vcard_id for c in contacts))
        merged.raw_vcard = ""  # Will be regenerated
        merged.original_lines = []  # Will be regenerated

        # Merge names (prefer the new name if provided)
        if new_name:
            merged.full_name = new_name
            # Try to split new name into first/last
            name_parts = new_name.split()
            if len(name_parts) >= 2:
                merged.first_name = " ".join(name_parts[:-1])
                merged.last_name = name_parts[-1]
            else:
                merged.first_name = new_name
        else:
            # Use the name from the contact with most complete info
            best_contact = self._find_best_name_contact(contacts)
            merged.full_name = best_contact.full_name
            merged.first_name = best_contact.first_name
            merged.last_name = best_contact.last_name

        # Merge phones (avoid duplicates)
        all_phones = {}
        for contact in contacts:
            for phone in contact.phones:
                # Use normalized number as key to avoid duplicates
                normalized = self._normalize_phone_for_merge(phone.number)
                if normalized not in all_phones:
                    all_phones[normalized] = phone

        merged.phones = list(all_phones.values())

        # Merge emails (avoid duplicates)
        all_emails = {}
        for contact in contacts:
            for email in contact.emails:
                email_key = email.address.lower()
                if email_key not in all_emails:
                    all_emails[email_key] = email

        merged.emails = list(all_emails.values())

        # Merge addresses (keep all unique ones)
        for contact in contacts:
            for address in contact.addresses:
                # Check if this address is already present
                if not self._address_exists(address, merged.addresses):
                    merged.addresses.append(address)

        # Merge other fields (use non-empty values)
        for contact in contacts:
            if contact.organization and not merged.organization:
                merged.organization = contact.organization
            if contact.title and not merged.title:
                merged.title = contact.title
            if contact.notes and not merged.notes:
                merged.notes = contact.notes

        # Carry unmodelled properties (PHOTO, BDAY, URL, X-*, ...) through the merge
        for contact in contacts:
            for extra in contact.extra_lines:
                if extra not in merged.extra_lines:
                    merged.extra_lines.append(extra)

        return merged

    def _find_best_name_contact(self, contacts: List[Contact]) -> Contact:
        """Find the contact with the best/most complete name.

        Args:
            contacts: List of contacts

        Returns:
            Contact with best name
        """
        if len(contacts) <= 2:
            # For 2 or fewer contacts, use the existing logic
            # Prefer contacts with both first and last name
            for contact in contacts:
                if contact.first_name and contact.last_name:
                    return contact

            # Then prefer contacts with full name
            for contact in contacts:
                if contact.full_name:
                    return contact

            # Finally, return the first contact
            return contacts[0]
        else:
            # For more than 2 contacts, try to find a common name pattern
            names = [c.get_display_name() for c in contacts if c.get_display_name()]
            if len(names) >= 2:
                # Find the most common pattern among names
                common_name = self._find_common_name_among_multiple(names)
                if common_name:
                    # Return the contact whose name matches this common pattern
                    for contact in contacts:
                        if contact.get_display_name() == common_name:
                            return contact

            # Fall back to the original logic
            return self._find_best_name_contact_legacy(contacts[:2])

    def _find_best_name_contact_legacy(self, contacts: List[Contact]) -> Contact:
        """Legacy method for finding best name contact (for <= 2 contacts).

        Args:
            contacts: List of contacts (max 2)

        Returns:
            Contact with best name
        """
        # Prefer contacts with both first and last name
        for contact in contacts:
            if contact.first_name and contact.last_name:
                return contact

        # Then prefer contacts with full name
        for contact in contacts:
            if contact.full_name:
                return contact

        # Finally, return the first contact
        return contacts[0]

    def _find_common_name_among_multiple(self, names: List[str]) -> Optional[str]:
        """Find the most common name pattern among multiple names.

        Args:
            names: List of names to analyze

        Returns:
            Common name pattern or None
        """
        if len(names) < 2:
            return names[0] if names else None

        # Find the longest common substring among all names
        common_substring = names[0]
        for name in names[1:]:
            common_substring = longest_common_substring(common_substring, name)
            if not common_substring:
                break

        if common_substring and len(common_substring) >= 3:
            return common_substring

        return None

    def _normalize_phone_for_merge(self, phone: str) -> str:
        """Normalize phone number for duplicate detection during merge.

        Args:
            phone: Phone number to normalize

        Returns:
            Normalized phone number
        """
        # Remove all non-digit characters except +
        return re.sub(r'[^\d\+]', '', phone.strip())

    def _address_exists(self, address: ContactAddress, existing_addresses: List[ContactAddress]) -> bool:
        """Check if an address already exists in the list.

        Args:
            address: Address to check
            existing_addresses: List of existing addresses

        Returns:
            True if address exists, False otherwise
        """
        for existing in existing_addresses:
            # Compare key fields
            if (address.street == existing.street and
                address.city == existing.city and
                address.postal_code == existing.postal_code):
                return True
        return False


class VCardWriter:
    """Handles writing contacts back to vCard format."""

    def __init__(self):
        """Initialize the vCard writer."""
        pass

    def write_contacts(self, contacts: List[Contact], file_path: str):
        """Write contacts to a vCard file.

        Args:
            file_path: Output file path
            contacts: List of contacts to write
        """
        logger.info(f"📝 Writing {len(contacts)} contacts to {file_path}")

        vcard_lines = []

        for contact in contacts:
            vcard_lines.extend(self._contact_to_vcard_lines(contact))
            vcard_lines.append("")  # Empty line between contacts

        # Write to file with proper line endings
        with open(file_path, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write('\n'.join(vcard_lines))

        logger.info(f"✅ Successfully wrote contacts to {file_path}")

    def _contact_to_vcard_lines(self, contact: Contact) -> List[str]:
        """Convert a Contact object to vCard format lines.

        Args:
            contact: Contact to convert

        Returns:
            List of vCard format lines
        """
        # An untouched contact keeps its original card verbatim, so fields this
        # tool does not model (PHOTO, BDAY, URL, CATEGORIES, X-*) survive the rewrite
        raw_lines = [line.rstrip('\r') for line in contact.raw_vcard.strip().split('\n')]
        if raw_lines and raw_lines[0] == "BEGIN:VCARD" and raw_lines[-1] == "END:VCARD":
            return raw_lines

        lines = [
            "BEGIN:VCARD",
            "VERSION:3.0"
        ]

        # Name fields
        if contact.full_name:
            lines.append(f"FN:{contact.full_name}")

        name_parts = [contact.last_name, contact.first_name, "", "", ""]
        lines.append(f"N:{';'.join(name_parts)}")

        # Phone numbers
        for phone in contact.phones:
            if has_type_param(phone.original_line):
                lines.append(phone.original_line)
            else:
                lines.append(f"TEL;TYPE={phone.type_label}:{phone.number}")

        # Email addresses
        for email in contact.emails:
            if has_type_param(email.original_line):
                lines.append(email.original_line)
            else:
                lines.append(f"EMAIL;TYPE={email.type_label}:{email.address}")

        # Addresses
        for address in contact.addresses:
            if address.original_line:
                lines.append(address.original_line)
            else:
                # Construct address line
                addr_parts = ["", "", address.street, address.city,
                            address.region, address.postal_code, address.country]
                lines.append(f"ADR;TYPE={address.type_label}:{';'.join(addr_parts)}")

        # Other fields
        if contact.organization:
            lines.append(f"ORG:{contact.organization}")
        if contact.title:
            lines.append(f"TITLE:{contact.title}")
        if contact.notes:
            lines.append(f"NOTE:{contact.notes}")

        # Unmodelled properties carried over from the source contacts
        for extra in contact.extra_lines:
            lines.extend(extra.split('\n'))

        lines.append("END:VCARD")
        return lines


class UnificationReport:
    """Generates reports of the unification process."""

    def __init__(self):
        """Initialize the report generator."""
        self.unified_groups = []
        self.cancelled_contacts = []
        self.kept_contacts = []

    def add_unified_group(self, original_contacts: List[Contact], merged_contact: Contact):
        """Add a unified contact group to the report.

        Args:
            original_contacts: List of original contacts that were merged
            merged_contact: The resulting merged contact
        """
        self.unified_groups.append({
            'original_contacts': original_contacts,
            'merged_contact': merged_contact
        })

    def add_cancelled_contact(self, contact: Contact):
        """Add a cancelled contact to the report.

        Args:
            contact: Contact that was cancelled
        """
        self.cancelled_contacts.append(contact)

    def add_kept_contact(self, contact: Contact):
        """Add a kept contact to the report.

        Args:
            contact: Contact that was kept as-is
        """
        self.kept_contacts.append(contact)

    def generate_report(self, output_file: str):
        """Generate a comprehensive report.

        Args:
            output_file: Path to write the report
        """
        logger.info(f"📊 Generating unification report: {output_file}")

        report_lines = [
            "# vCard Unification Report",
            f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "## Summary",
            f"- Contacts unified: {len(self.unified_groups)} groups",
            f"- Contacts cancelled: {len(self.cancelled_contacts)}",
            f"- Contacts kept: {len(self.kept_contacts)}",
            "",
        ]

        # Unified contacts section
        if self.unified_groups:
            report_lines.extend([
                "## Unified Contacts",
                "",
            ])

            for i, group in enumerate(self.unified_groups, 1):
                merged = group['merged_contact']
                originals = group['original_contacts']

                report_lines.extend([
                    f"### Group {i}: {merged.get_display_name()}",
                    f"**Merged from {len(originals)} contacts:**",
                ])

                for orig in originals:
                    phones = ", ".join([str(p) for p in orig.phones])
                    report_lines.append(f"- {orig.get_display_name()} (Phones: {phones})")

                report_lines.extend([
                    "",
                    f"**Final contact:** {merged.get_display_name()}",
                    f"- Phones: {len(merged.phones)}",
                    f"- Emails: {len(merged.emails)}",
                    f"- Addresses: {len(merged.addresses)}",
                    "",
                ])

        # Cancelled contacts section
        if self.cancelled_contacts:
            report_lines.extend([
                "## Cancelled Contacts",
                "",
            ])

            for contact in self.cancelled_contacts:
                phones = ", ".join([str(p) for p in contact.phones])
                report_lines.extend([
                    f"- {contact.get_display_name()} (Phones: {phones})",
                ])

            report_lines.append("")

        # Kept contacts section
        if self.kept_contacts:
            report_lines.extend([
                "## Kept Contacts",
                "",
                f"Total: {len(self.kept_contacts)} contacts kept as-is",
                "",
            ])

        # Write report
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))

        logger.info(f"✅ Report generated: {output_file}")


class InteractiveUnifier:
    """Interactive interface for contact unification."""

    def __init__(self):
        """Initialize the interactive unifier."""
        self.parser = VCardParser()
        self.detector = DuplicateDetector()
        self.merger = ContactMerger()
        self.writer = VCardWriter()
        self.report = UnificationReport()

    def unify_file(self, input_file: str, output_file: str = "", report_file: str = ""):
        """Main unification process.

        Args:
            input_file: Path to input vCard file
            output_file: Path to output unified vCard file (auto-generated if empty)
            report_file: Path to report file (auto-generated if empty)
        """
        logger.info("🚀 Starting vCard unification process")

        # Generate output paths if not provided
        if not output_file:
            input_path = Path(input_file)
            output_file = str(input_path.parent / f"{input_path.stem}_unified{input_path.suffix}")

        if not report_file:
            input_path = Path(input_file)
            report_file = str(input_path.parent / f"{input_path.stem}_unification_report.md")

        # Parse contacts
        contacts = self.parser.parse_file(input_file)
        if not contacts:
            logger.error("❌ No contacts found in input file")
            return

        logger.info(f"📊 Processing {len(contacts)} contacts")

        # Interactive unification process
        final_contacts = []
        processed_ids = set()

        # One pass: merged contacts are never fed back into the scan, so a second
        # scan could not find a pair the first one did not already report.
        logger.info("🔄 Finding duplicates...")
        duplicates = self.detector.find_duplicates(contacts)

        if not duplicates:
            logger.info("✅ No duplicates found")
        else:
            logger.info(f"Found {len(duplicates)} potential duplicate pairs to review")

        for i, (contact1, contact2, reason) in enumerate(duplicates, 1):
            # Skip if either contact already processed (a contact can sit in several pairs)
            if contact1.vcard_id in processed_ids or contact2.vcard_id in processed_ids:
                logger.debug(f"⏭️  Skipping pair ({contact1.vcard_id}, {contact2.vcard_id}) - already processed")
                continue

            self._process_duplicate_pair(i, len(duplicates), contact1, contact2, reason,
                                         final_contacts, processed_ids)

        # Add remaining unprocessed contacts
        for contact in contacts:
            if contact.vcard_id not in processed_ids:
                final_contacts.append(contact)
                self.report.add_kept_contact(contact)

        # Write unified file
        self.writer.write_contacts(final_contacts, output_file)

        # Generate report
        self.report.generate_report(report_file)

        logger.info("✅ Unification process completed!")
        logger.info(f"📄 Unified file: {output_file}")
        logger.info(f"📊 Report: {report_file}")

    def _process_duplicate_pair(self, index: int, total: int, contact1: Contact,
                              contact2: Contact, reason: str, final_contacts: List[Contact],
                              processed_ids: Set[int]) -> bool:
        """Process a single duplicate pair interactively.

        Args:
            index: Current pair index
            total: Total number of pairs
            contact1: First contact
            contact2: Second contact
            reason: Reason they are considered duplicates
            final_contacts: List to add final contacts to
            processed_ids: Set of already processed contact IDs

        Returns:
            True if contacts were processed (merged, kept, or cancelled), False if skipped
        """
        print(f"\n{'='*60}")
        print(f"🔍 DUPLICATE PAIR {index}/{total}")
        print(f"{'='*60}")
        print(f"Reason: {reason}")
        print()

        # Display contact details
        self._display_contact_details(contact1, "Contact 1")
        print()
        self._display_contact_details(contact2, "Contact 2")
        print()

        # Ask user for decision
        while True:
            print("Options:")
            print("1. Merge these contacts [DEFAULT]")
            print("2. Keep both contacts (no merge)")
            print("3. Cancel Contact 1 (keep only Contact 2)")
            print("4. Cancel Contact 2 (keep only Contact 1)")
            print("5. Skip this pair")

            try:
                choice = input("Choose option (1-5) [press Enter for 1]: ").strip()

                # Default to option 1 if empty
                if choice == "":
                    choice = "1"

                if choice == "1":
                    self._handle_merge(contact1, contact2, final_contacts, processed_ids)
                    return True
                elif choice == "2":
                    final_contacts.append(contact1)
                    final_contacts.append(contact2)
                    self.report.add_kept_contact(contact1)
                    self.report.add_kept_contact(contact2)
                    processed_ids.update([contact1.vcard_id, contact2.vcard_id])
                    return True
                elif choice == "3":
                    final_contacts.append(contact2)
                    self.report.add_cancelled_contact(contact1)
                    self.report.add_kept_contact(contact2)
                    processed_ids.update([contact1.vcard_id, contact2.vcard_id])
                    return True
                elif choice == "4":
                    final_contacts.append(contact1)
                    self.report.add_kept_contact(contact1)
                    self.report.add_cancelled_contact(contact2)
                    processed_ids.update([contact1.vcard_id, contact2.vcard_id])
                    return True
                elif choice == "5":
                    # Don't add either contact yet, let them be processed later if found in other pairs
                    return False
                else:
                    print("❌ Invalid choice. Please enter 1-5.")
            except KeyboardInterrupt:
                print("\n⏹️  Process interrupted by user")
                raise
            except Exception as e:
                print(f"❌ Error: {e}")
                return False

    def _display_contact_details(self, contact: Contact, label: str):
        """Display contact details in a formatted way.

        Args:
            contact: Contact to display
            label: Label for the contact (e.g., "Contact 1")
        """
        print(f"{label}: {contact.get_display_name()}")
        print(f"  ID: {contact.vcard_id}")

        if contact.phones:
            print(f"  📞 Phones: {len(contact.phones)}")
            for phone in contact.phones:
                print(f"    - {phone}")

        if contact.emails:
            print(f"  📧 Emails: {len(contact.emails)}")
            for email in contact.emails:
                print(f"    - {email}")

        if contact.addresses:
            print(f"  📍 Addresses: {len(contact.addresses)}")
            for addr in contact.addresses:
                print(f"    - {addr}")

        if contact.organization:
            print(f"  🏢 Organization: {contact.organization}")

    def _handle_merge(self, contact1: Contact, contact2: Contact,
                     final_contacts: List[Contact], processed_ids: Set[int]):
        """Handle the merge process for two contacts.

        Args:
            contact1: First contact
            contact2: Second contact
            final_contacts: List to add merged contact to
            processed_ids: Set of processed contact IDs
        """
        # Suggest a merged name based on common parts
        suggested_name = self._suggest_merged_name(contact1, contact2)

        print(f"\n💡 Suggested merged name (common part): '{suggested_name}'")

        # Ask if user wants to change the name
        while True:
            name_choice = input("Use suggested name? (y/n) [press Enter for y]: ").strip().lower()
            if name_choice in ['y', 'yes', '']:
                final_name = suggested_name
                break
            elif name_choice in ['n', 'no']:
                final_name = input("Enter new name: ").strip()
                if not final_name:
                    final_name = suggested_name
                break
            else:
                print("Please enter 'y' or 'n'")

        # Merge contacts
        merged_contact = self.merger.merge_contacts([contact1, contact2], final_name)
        final_contacts.append(merged_contact)

        # Record in report
        self.report.add_unified_group([contact1, contact2], merged_contact)

        # Mark as processed
        processed_ids.update([contact1.vcard_id, contact2.vcard_id])

        print(f"✅ Contacts merged as: {merged_contact.get_display_name()}")

    def _suggest_merged_name(self, contact1: Contact, contact2: Contact) -> str:
        """Suggest a name for the merged contact based on common parts.

        Args:
            contact1: First contact
            contact2: Second contact

        Returns:
            Suggested merged name based on common parts
        """
        name1 = contact1.get_display_name().strip()
        name2 = contact2.get_display_name().strip()

        # If one name is contained in the other, use the shorter/common one
        if name1.lower() in name2.lower():
            return name1  # name1 is shorter/common part
        elif name2.lower() in name1.lower():
            return name2  # name2 is shorter/common part

        # Find the longest common substring
        common_substring = self._find_longest_common_substring(name1, name2)

        if common_substring and len(common_substring) >= 3:
            return common_substring.strip()

        # If no good common substring, find common prefix
        common_prefix = self._find_common_prefix(name1, name2)
        if common_prefix and len(common_prefix) >= 3:
            return common_prefix.strip()

        # If no common parts found, prefer the longer name
        if len(name1) >= len(name2):
            return name1
        else:
            return name2

    def _find_longest_common_substring(self, str1: str, str2: str) -> str:
        """Find the longest common substring between two strings.

        Args:
            str1: First string
            str2: Second string

        Returns:
            Longest common substring
        """
        str1_lower = str1.lower()
        str2_lower = str2.lower()

        # Check if one string is entirely contained in the other
        if str1_lower in str2_lower:
            return str1
        elif str2_lower in str1_lower:
            return str2

        return longest_common_substring(str1, str2)

    def _find_common_prefix(self, str1: str, str2: str) -> str:
        """Find the common prefix between two strings.

        Args:
            str1: First string
            str2: Second string

        Returns:
            Common prefix
        """
        min_len = min(len(str1), len(str2))
        for i in range(min_len):
            if str1[i].lower() != str2[i].lower():
                return str1[:i]
        return str1[:min_len]


def select_input_file():
    """Show a file dialog to select the input vCard file.

    Returns:
        Path to the selected file, or None if cancelled
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Configure the file dialog
    file_path = filedialog.askopenfilename(
        title="Select vCard file to unify",
        filetypes=[
            ("vCard files", "*.vcf"),
            ("All files", "*.*")
        ],
        initialdir=os.getcwd()
    )

    root.destroy()
    return file_path if file_path else None


def main():
    """Main function to run the vCard unification tool."""
    import sys

    # Get input file from command line or file dialog
    if len(sys.argv) < 2:
        logger.info("📂 No input file specified, opening file selector...")
        input_file = select_input_file()
        if not input_file:
            logger.error("❌ No file selected. Exiting.")
            sys.exit(1)
    else:
        input_file = sys.argv[1]

    # Get output files from command line arguments
    output_file = sys.argv[2] if len(sys.argv) > 2 else ""
    report_file = sys.argv[3] if len(sys.argv) > 3 else ""

    # Validate input file exists
    if not os.path.exists(input_file):
        logger.error("❌ Input file not found: %s", input_file)
        sys.exit(1)

    try:
        unifier = InteractiveUnifier()
        unifier.unify_file(input_file, output_file, report_file)
    except KeyboardInterrupt:
        logger.info("⏹️  Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.error("❌ Application failed: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
