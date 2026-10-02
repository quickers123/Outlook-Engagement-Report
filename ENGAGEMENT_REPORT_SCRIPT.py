#!/usr/bin/env python3
r"""
Standalone V10.3 Outlook OST external-engagement reporter.

Workflow:
1. Archives prior External_Contacts_Reports_* folders into ./Archive.
2. Discovers mailbox stores/inboxes and infers their SMTP addresses from sampled message headers.
3. Prompts for one mailbox or a comma-separated combination, then asks whether to include external, internal, or all email addresses.
4. Parses the OST once for the selected scope and deduplicates messages.
5. Creates one timestamped report workbook with:
   - External Contacts
   - native Excel PivotTable based only on External Contacts
   - Email Interactions, appended after the pivot stage

The interactions sheet includes ExternalDomain and readable plain-text bodies.
Long interaction content is never passed through Excel COM.

Dependencies:
  python -m pip install libpff-python-windows openpyxl pywin32
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email import policy
from email.parser import Parser
from email.utils import getaddresses
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable

INTERNAL_DOMAINS_DEFAULT = {"tewhatuora.govt.nz"}
EXCEL_CELL_LIMIT = 32767
EMAIL_RE = re.compile(r"(?i)(?<![\w.+-])([a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,})(?![\w.-])")
AUTOMATED_LOCAL_PARTS = {"mailer-daemon", "postmaster", "no-reply", "noreply", "donotreply", "do-not-reply", "bounce", "notifications", "notification"}
BLOCK_TAGS = {"address", "article", "aside", "blockquote", "div", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "main", "nav", "p", "pre", "section", "table", "tr"}
IGNORE_TAGS = {"head", "style", "script", "title", "svg", "xml", "noscript"}
INBOX_NAMES = {"inbox", "boîte de réception", "posteingang"}
SENT_NAMES = {"sent items", "sent", "éléments envoyés", "gesendete elemente"}


def get_value(obj: Any, names: Iterable[str], default: Any = "") -> Any:
    for name in names:
        try:
            value = getattr(obj, name)
            if callable(value):
                value = value()
            if value is not None:
                return value
        except Exception:
            continue
    return default


def normalise_address(value: str) -> str:
    value = (value or "").strip().strip("<>").lower()
    return value[5:] if value.startswith("smtp:") else value


def normalise_mailbox_list(value: Any) -> list[str]:
    if value is None:
        return []
    parts = [str(v) for v in value] if isinstance(value, (set, list, tuple)) else re.split(r"\s*;\s*", str(value))
    return sorted({normalise_address(v) for v in parts if normalise_address(v) and normalise_address(v) != "0"})


def mailbox_string(value: Any) -> str:
    return "; ".join(normalise_mailbox_list(value))


def addresses_from_text(value: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    seen: set[str] = set()
    for name, addr in getaddresses([value or ""]):
        addr = normalise_address(addr)
        if "@" in addr and addr not in seen:
            seen.add(addr)
            found.append(((name or "").strip(), addr))
    for addr in EMAIL_RE.findall(value or ""):
        addr = normalise_address(addr)
        if addr not in seen:
            seen.add(addr)
            found.append(("", addr))
    return found


def domain_of(address: str) -> str:
    return address.rsplit("@", 1)[1].lower() if "@" in address else ""


def is_internal(address: str, domains: set[str]) -> bool:
    domain = domain_of(address)
    return any(domain == d or domain.endswith("." + d) for d in domains)


def is_automated(address: str) -> bool:
    return address.split("@", 1)[0].lower() in AUTOMATED_LOCAL_PARTS if "@" in address else False


def iso_datetime(value: Any) -> str:
    if not value:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ") if value.tzinfo is None else value.astimezone(timezone.utc).isoformat()
    return str(value)


def clean_cell(value: Any, limit: int = EXCEL_CELL_LIMIT) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="replace")
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(value))
    if len(text) > limit:
        suffix = "\n[Content truncated at Excel's 32,767-character cell limit]"
        text = text[: limit - len(suffix)] + suffix
    return text


def transport_headers(message: Any) -> str:
    value = get_value(message, ["transport_headers", "get_transport_headers"], "")
    return value.decode("utf-8", errors="replace") if isinstance(value, bytes) else str(value or "")


def parse_headers(raw: str):
    try:
        return Parser(policy=policy.default).parsestr(raw or "")
    except Exception:
        return Parser().parsestr(raw or "")


class VisibleTextParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.ignore_depth = 0
        self.list_depth = 0

    def newline(self):
        if self.parts and not self.parts[-1].endswith("\n"):
            self.parts.append("\n")

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in IGNORE_TAGS:
            self.ignore_depth += 1
            return
        if self.ignore_depth:
            return
        if tag in BLOCK_TAGS or tag == "br":
            self.newline()
        elif tag in {"ul", "ol"}:
            self.list_depth += 1
            self.newline()
        elif tag == "li":
            self.newline()
            self.parts.append("  " * max(self.list_depth - 1, 0) + "• ")
        elif tag in {"td", "th"} and self.parts and not self.parts[-1].endswith(("\n", "\t")):
            self.parts.append("\t")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in IGNORE_TAGS:
            if self.ignore_depth:
                self.ignore_depth -= 1
            return
        if self.ignore_depth:
            return
        if tag in BLOCK_TAGS or tag in {"li", "td", "th"}:
            self.newline()
        elif tag in {"ul", "ol"}:
            self.list_depth = max(self.list_depth - 1, 0)
            self.newline()

    def handle_data(self, data):
        if not self.ignore_depth and data:
            self.parts.append(data)

    def handle_comment(self, data):
        return


def decode_body(value: Any) -> str:
    if value is None:
        return ""
    if not isinstance(value, bytes):
        return str(value)
    prefix = value[:4096].decode("ascii", errors="ignore")
    match = re.search(r"charset\s*=\s*[\"']?([A-Za-z0-9._-]+)", prefix, re.I)
    encodings = ([match.group(1)] if match else []) + ["utf-8", "windows-1252", "latin-1"]
    for encoding in encodings:
        try:
            return value.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            pass
    return value.decode("utf-8", errors="replace")


def looks_like_html(text: str) -> bool:
    sample = html.unescape(text[:4000]).lstrip().lower()
    return bool(re.search(r"<(?:!doctype\s+html|html|head|body|div|p|span|table|br)\b", sample))


def html_to_plain_text(value: Any) -> str:
    source = html.unescape(decode_body(value))
    parser = VisibleTextParser()
    try:
        parser.feed(source)
        parser.close()
        text = "".join(parser.parts)
    except Exception:
        text = re.sub(r"(?is)<(?:style|script|head)\b.*?</(?:style|script|head)>", "", source)
        text = re.sub(r"(?i)<br\s*/?>|</p\s*>|</div\s*>|</tr\s*>", "\n", text)
        text = html.unescape(re.sub(r"(?s)<[^>]+>", "", text))
    text = text.replace("\u00a0", " ").replace("\u200b", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def message_body(message: Any) -> tuple[str, str]:
    plain = decode_body(get_value(message, ["plain_text_body", "get_plain_text_body", "body", "get_body"], ""))
    html_body = decode_body(get_value(message, ["html_body", "get_html_body"], ""))
    if plain and not looks_like_html(plain):
        return clean_cell(plain), "Plain text"
    source = plain if plain else html_body
    if source:
        text = html_to_plain_text(source)
        if text:
            return clean_cell(text), "HTML converted to plain text"
    return "", "Unavailable"


def decode_attachment_name(value: Any) -> str:
    """Decode and normalise a possible attachment filename."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        for encoding in ("utf-16-le", "utf-8", "windows-1252", "latin-1"):
            try:
                candidate = value.decode(encoding).rstrip("\x00").strip()
                if candidate:
                    printable = sum(character.isprintable() for character in candidate)
                    if printable / max(len(candidate), 1) > 0.85:
                        return candidate
            except Exception:
                continue
        return value.decode("utf-8", errors="replace").rstrip("\x00").strip()
    return str(value).rstrip("\x00").strip()


def entry_property_id(entry: Any) -> int | None:
    """Return the low 16-bit MAPI property identifier where available."""
    value = get_value(
        entry,
        [
            "entry_type", "get_entry_type",
            "type", "get_type",
            "identifier", "get_identifier",
        ],
        None,
    )
    if value is None:
        return None
    try:
        return int(value) & 0xFFFF
    except (TypeError, ValueError):
        return None


def entry_text(entry: Any) -> str:
    """Read a record-entry value using the available pypff API surface."""
    for names in (
        ["data_as_string", "get_data_as_string"],
        ["value", "get_value"],
        ["data", "get_data"],
    ):
        value = get_value(entry, names, None)
        name = decode_attachment_name(value)
        if name:
            return name
    return ""


def attachment_name_from_records(attachment: Any) -> str:
    """Read standard MAPI attachment filename fields from record sets."""
    # PR_ATTACH_LONG_FILENAME, PR_ATTACH_FILENAME, PR_DISPLAY_NAME,
    # PR_ATTACH_CONTENT_ID. Long filename is deliberately preferred.
    priority = {0x3707: 0, 0x3704: 1, 0x3001: 2, 0x3712: 3}
    matches: list[tuple[int, str]] = []
    record_set_count = int(get_value(
        attachment,
        ["number_of_record_sets", "get_number_of_record_sets"],
        0,
    ) or 0)

    for set_index in range(record_set_count):
        try:
            record_set = attachment.get_record_set(set_index)
        except Exception:
            continue
        entry_count = int(get_value(
            record_set,
            ["number_of_entries", "get_number_of_entries"],
            0,
        ) or 0)
        for entry_index in range(entry_count):
            try:
                entry = record_set.get_entry(entry_index)
                property_id = entry_property_id(entry)
                if property_id not in priority:
                    continue
                name = entry_text(entry)
                if name:
                    matches.append((priority[property_id], name))
            except Exception:
                continue

    if not matches:
        return ""
    matches.sort(key=lambda item: (item[0], -len(item[1])))
    return matches[0][1]


def attachment_names(message: Any) -> tuple[str, int]:
    """Return actual attachment filenames where exposed by the OST."""
    count = int(get_value(
        message,
        ["number_of_attachments", "get_number_of_attachments"],
        0,
    ) or 0)
    names: list[str] = []

    for index in range(count):
        try:
            attachment = message.get_attachment(index)
            name = ""

            # pypff's documented attachment API exposes get_name(). Other
            # builds expose filename-like properties, so test all variants.
            for names_to_try in (
                ["name", "get_name"],
                ["long_filename", "get_long_filename"],
                ["filename", "get_filename"],
                ["display_name", "get_display_name"],
            ):
                candidate = get_value(attachment, names_to_try, None)
                name = decode_attachment_name(candidate)
                if name:
                    break

            if not name:
                name = attachment_name_from_records(attachment)

            if name:
                names.append(clean_cell(name, 500))
            else:
                size = int(get_value(
                    attachment,
                    ["size", "get_size"],
                    0,
                ) or 0)
                if size:
                    names.append(
                        f"Attachment {index + 1} "
                        f"(filename unavailable; {size:,} bytes)"
                    )
                else:
                    names.append(
                        f"Attachment {index + 1} (filename unavailable)"
                    )
        except Exception as exc:
            names.append(
                f"Attachment {index + 1} (unreadable: {type(exc).__name__})"
            )

    return "; ".join(names), count



def folder_name(folder: Any) -> str:
    return str(get_value(folder, ["name", "get_name"], "(unnamed)"))


def subfolder_count(folder: Any) -> int:
    return int(get_value(folder, ["number_of_sub_folders", "get_number_of_sub_folders"], 0) or 0)


def message_count(folder: Any) -> int:
    return int(get_value(folder, ["number_of_sub_messages", "get_number_of_sub_messages"], 0) or 0)


def folder_children(folder: Any) -> list[Any]:
    children = []
    for index in range(subfolder_count(folder)):
        try:
            children.append(folder.get_sub_folder(index))
        except Exception:
            continue
    return children


def collect_candidate_addresses(folder: Any, internal_domains: set[str], limit: int = 600) -> dict[str, int]:
    """Sample message headers under one mailbox root and count internal addresses."""
    counts: dict[str, int] = {}
    visited = 0

    def walk(current: Any, depth: int):
        nonlocal visited
        if visited >= limit or depth > 7:
            return
        for index in range(message_count(current)):
            if visited >= limit:
                return
            visited += 1
            try:
                message = current.get_sub_message(index)
                headers = parse_headers(transport_headers(message))
                values = [
                    str(headers.get("From") or ""),
                    str(headers.get("To") or ""),
                    str(headers.get("Cc") or ""),
                    str(headers.get("Bcc") or ""),
                    str(get_value(message, ["sender_email_address"], "") or ""),
                ]
                for _, address in addresses_from_text("; ".join(values)):
                    if is_internal(address, internal_domains):
                        counts[address] = counts.get(address, 0) + 1
            except Exception:
                continue
        for child in folder_children(current):
            walk(child, depth + 1)
            if visited >= limit:
                return

    walk(folder, 0)
    return counts


def infer_mailbox_address(folder: Any, name: str, path: str, internal_domains: set[str]) -> str:
    """Infer the mailbox SMTP address from folder identity plus sampled headers."""
    direct = EMAIL_RE.search(name) or EMAIL_RE.search(path)
    if direct:
        return normalise_address(direct.group(1))

    counts = collect_candidate_addresses(folder, internal_domains)
    if not counts:
        return ""

    # Strongly favour an address whose local part matches the mailbox display
    # name, e.g. HDIP_Engagement -> hdip_engagement@tewhatuora.govt.nz.
    normalised_name = re.sub(r"[^a-z0-9]+", "", name.lower())
    scored = []
    for address, count in counts.items():
        local = address.split("@", 1)[0]
        normalised_local = re.sub(r"[^a-z0-9]+", "", local.lower())
        name_match = 1 if normalised_name and normalised_name in normalised_local else 0
        scored.append((name_match, count, address))
    scored.sort(key=lambda item: (-item[0], -item[1], item[2]))
    return scored[0][2]


def mailbox_label(name: str, path: str) -> str:
    match = EMAIL_RE.search(name) or EMAIL_RE.search(path)
    if match:
        return normalise_address(match.group(1))
    cleaned = re.sub(r"^(mailbox\s*[-:]\s*|top of information store\s*[-:]\s*)", "", name, flags=re.I).strip()
    return cleaned or name


def discover_mailboxes(root: Any, internal_domains: set[str]) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []

    def walk(folder: Any, parts: list[str], depth: int):
        name = folder_name(folder)
        path_parts = parts + [name]
        path = "/".join(path_parts)
        children: list[tuple[Any, str]] = []
        for index in range(subfolder_count(folder)):
            try:
                child = folder.get_sub_folder(index)
                children.append((child, folder_name(child)))
            except Exception:
                continue
        child_names = {child_name.strip().lower() for _, child_name in children}
        if child_names & INBOX_NAMES:
            inferred = infer_mailbox_address(folder, name, path, internal_domains)
            display = inferred or mailbox_label(name, path)
            candidates.append({
                "label": display,
                "email": inferred,
                "folder_name": name,
                "path": path,
            })
        if depth < 5:
            for child, _ in children:
                walk(child, path_parts, depth + 1)

    walk(root, [], 0)
    unique: dict[str, dict[str, str]] = {}
    for item in candidates:
        key = item.get("email", "").lower() or item["path"].lower()
        existing = unique.get(key)
        if existing is None:
            unique[key] = item
        else:
            # Prefer a human-readable mailbox folder over the generic
            # IPM_SUBTREE container when both resolve to the same address.
            existing_generic = existing.get("folder_name", "").upper() == "IPM_SUBTREE"
            current_generic = item.get("folder_name", "").upper() == "IPM_SUBTREE"
            if existing_generic and not current_generic:
                unique[key] = item
    result = list(unique.values())
    result.sort(key=lambda item: (item["label"].lower(), item["path"].lower()))
    return result


def choose_mailboxes(mailboxes: list[dict[str, str]]) -> list[dict[str, str]]:
    if not mailboxes:
        raise RuntimeError("No mailbox folders with an Inbox were detected in the OST.")
    print("\nMailboxes/inboxes detected in the OST:")
    for index, item in enumerate(mailboxes, start=1):
        print(f"  {index}. {item['label']}  [{item['path']}]")

    while True:
        mode = input("\nReport scope: enter 1 for one inbox, or 2 for a combination: ").strip().lower()
        if mode in {"1", "one", "single", "s"}:
            while True:
                raw = input(f"Select one inbox number (1-{len(mailboxes)}): ").strip()
                if raw.isdigit() and 1 <= int(raw) <= len(mailboxes):
                    return [mailboxes[int(raw) - 1]]
                print("Please enter one valid inbox number.")
        if mode in {"2", "combination", "combined", "c", "multiple", "multi"}:
            while True:
                raw = input("Enter inbox numbers separated by commas, for example 1, 3: ").strip()
                tokens = [token.strip() for token in raw.split(",") if token.strip()]
                if tokens and all(token.isdigit() for token in tokens):
                    numbers: list[int] = []
                    for token in tokens:
                        number = int(token)
                        if number not in numbers:
                            numbers.append(number)
                    if len(numbers) >= 2 and all(1 <= number <= len(mailboxes) for number in numbers):
                        return [mailboxes[number - 1] for number in numbers]
                print("Enter at least two valid, comma-separated inbox numbers. Spaces are allowed.")
        print("Please enter 1 or 2.")


def choose_address_scope() -> str:
    """Ask whether the report should include external, internal, or all contacts."""
    print("\nEmail address scope:")
    print("  1. External addresses only (not @tewhatuora.govt.nz)")
    print("  2. Internal addresses only (@tewhatuora.govt.nz)")
    print("  3. All addresses (internal and external)")
    aliases = {
        "1": "external", "external": "external", "e": "external",
        "2": "internal", "internal": "internal", "i": "internal",
        "3": "all", "all": "all", "a": "all", "both": "all",
    }
    while True:
        raw = input("Select address scope (1, 2, or 3): ").strip().lower()
        if raw in aliases:
            return aliases[raw]
        print("Please enter 1, 2, or 3.")


def address_allowed(address: str, scope: str, internal_domains: set[str], own_addresses: set[str]) -> bool:
    address = normalise_address(address)
    if not address or address in own_addresses:
        return False
    internal = is_internal(address, internal_domains)
    if scope == "external":
        return not internal
    if scope == "internal":
        return internal
    return True


def mailbox_for_path(folder_path: str, selected: list[dict[str, str]]) -> str | None:
    path_lower = folder_path.lower()
    matches = [item for item in selected if path_lower.startswith(item["path"].lower())]
    if not matches:
        return None
    return max(matches, key=lambda item: len(item["path"]))["label"]


def extract_message(message: Any, folder_path: str, source_mailbox: str, internal_domains: set[str], exclude_automated: bool, address_scope: str, own_addresses: set[str]):
    headers = parse_headers(transport_headers(message))
    subject = clean_cell(headers.get("Subject") or get_value(message, ["subject", "get_subject"], ""))
    message_id = clean_cell(headers.get("Message-ID") or headers.get("Message-Id") or "")
    from_value = clean_cell(headers.get("From") or get_value(message, ["sender_email_address", "sender_name"], ""))
    to_value = clean_cell(headers.get("To") or get_value(message, ["display_to", "get_display_to"], ""))
    cc_value = clean_cell(headers.get("Cc") or get_value(message, ["display_cc", "get_display_cc"], ""))
    bcc_value = clean_cell(headers.get("Bcc") or get_value(message, ["display_bcc", "get_display_bcc"], ""))
    from_pairs = addresses_from_text(from_value)
    to_pairs = addresses_from_text(to_value)
    cc_pairs = addresses_from_text(cc_value)
    bcc_pairs = addresses_from_text(bcc_value)
    contacts: dict[str, str] = {}
    for name, address in from_pairs + to_pairs + cc_pairs + bcc_pairs:
        if address_allowed(address, address_scope, internal_domains, own_addresses):
            if not (exclude_automated and is_automated(address)):
                contacts.setdefault(address, name)
    if not contacts:
        return None
    from_address = from_pairs[0][1] if from_pairs else ""
    if from_address and not is_internal(from_address, internal_domains):
        direction = "Inbound"
    elif from_address and any(not is_internal(address, internal_domains) for _, address in to_pairs + cc_pairs + bcc_pairs):
        direction = "Outbound"
    elif from_address and is_internal(from_address, internal_domains):
        direction = "Internal"
    else:
        direction = "Unresolved"
    value = get_value(message, ["client_submit_time", "delivery_time", "creation_time"], "") or headers.get("Date", "")
    date_text = iso_datetime(value)
    recipients = sorted(address for _, address in to_pairs + cc_pairs + bcc_pairs)
    key = message_id.lower() if message_id else "fallback:" + hashlib.sha256("|".join([from_address, ";".join(recipients), date_text, subject.lower()]).encode("utf-8", errors="replace")).hexdigest()
    body, body_format = message_body(message)
    attachment_text, attachment_count = attachment_names(message)
    contact_domains = sorted({domain_of(address) for address in contacts})
    contact_types = sorted({"Internal" if is_internal(address, internal_domains) else "External" for address in contacts})
    return {
        "DuplicateKey": key,
        "SourceMailboxes": {source_mailbox},
        "Date": date_text,
        "Direction": direction,
        "Contacts": contacts,
        "ContactDomain": "; ".join(contact_domains),
        "ContactType": "; ".join(contact_types),
        "FolderPath": folder_path,
        "From": from_value,
        "To": to_value,
        "Cc": cc_value,
        "Bcc": bcc_value,
        "Subject": subject,
        "MessageContent": body,
        "ContentFormat": body_format,
        "AttachmentNames": attachment_text,
        "AttachmentCount": attachment_count,
        "InternetMessageId": message_id,
    }


def traverse_selected(folder: Any, path_parts: list[str], selected: list[dict[str, str]], log: dict[str, Any], internal_domains: set[str], exclude_automated: bool, address_scope: str, own_addresses: set[str]):
    current = "/".join(path_parts + [folder_name(folder)])
    log["folders_seen"] += 1
    source_mailbox = mailbox_for_path(current, selected)
    if source_mailbox:
        for index in range(message_count(folder)):
            log["items_seen"] += 1
            try:
                row = extract_message(folder.get_sub_message(index), current, source_mailbox, internal_domains, exclude_automated, address_scope, own_addresses)
                if row:
                    yield row
            except Exception as exc:
                log["message_errors"] += 1
                if len(log["error_examples"]) < 100:
                    log["error_examples"].append({"folder": current, "index": index, "error": repr(exc)})
    for index in range(subfolder_count(folder)):
        try:
            yield from traverse_selected(folder.get_sub_folder(index), path_parts + [folder_name(folder)], selected, log, internal_domains, exclude_automated, address_scope, own_addresses)
        except Exception as exc:
            log["folder_errors"] += 1
            if len(log["error_examples"]) < 100:
                log["error_examples"].append({"folder": current, "subfolder_index": index, "error": repr(exc)})


def build_contacts(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    contacts: dict[str, dict[str, Any]] = {}
    for message in messages:
        for address in message["Contacts"]:
            item = contacts.setdefault(address, {
                "ContactEmail": address,
                "ContactDomain": domain_of(address),
                "ContactType": "Internal" if is_internal(address, INTERNAL_DOMAINS_DEFAULT) else "External",
                "FirstEngagement": message["Date"],
                "LastEngagement": message["Date"],
                "InboundMessageCount": 0,
                "OutboundMessageCount": 0,
                "UnresolvedMessageCount": 0,
                "TotalUniqueMessages": 0,
                "MailboxesEngagedThrough": set(),
            })
            dates = [date for date in (item["FirstEngagement"], message["Date"]) if date]
            item["FirstEngagement"] = min(dates) if dates else ""
            item["LastEngagement"] = max(item["LastEngagement"], message["Date"]) if message["Date"] else item["LastEngagement"]
            item["TotalUniqueMessages"] += 1
            item["MailboxesEngagedThrough"].update(message["SourceMailboxes"])
            if message["Direction"] == "Inbound":
                item["InboundMessageCount"] += 1
            elif message["Direction"] == "Outbound":
                item["OutboundMessageCount"] += 1
            else:
                item["UnresolvedMessageCount"] += 1
    rows = []
    for item in contacts.values():
        item["MailboxesEngagedThrough"] = mailbox_string(item["MailboxesEngagedThrough"])
        rows.append(item)
    return sorted(rows, key=lambda item: (-item["TotalUniqueMessages"], item["ContactDomain"].lower(), item["ContactEmail"].lower()))


def excel_datetime(value: str):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return value


def set_literal_text(cell: Any, value: Any, limit: int = EXCEL_CELL_LIMIT) -> None:
    """Write extracted content as literal text, never as an Excel formula."""
    cell.value = clean_cell(value, limit)
    cell.data_type = "s"


def validate_xlsx_package(output: Path) -> dict[str, Any]:
    """Validate the XLSX ZIP/XML package and reject formulas in data sheets."""
    worksheet_formulas: dict[str, int] = {}
    with zipfile.ZipFile(output, "r") as package:
        bad_member = package.testzip()
        if bad_member:
            raise RuntimeError(f"XLSX ZIP validation failed at {bad_member}")
        worksheet_names = sorted(
            name for name in package.namelist()
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        )
        for name in worksheet_names:
            root = ET.fromstring(package.read(name))
            formula_count = sum(1 for element in root.iter() if element.tag.endswith("}f"))
            worksheet_formulas[name] = formula_count
    # Sheet 1 is External Contacts and sheet 3 is Email Interactions in this workflow.
    for name in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml"):
        if worksheet_formulas.get(name, 0):
            raise RuntimeError(f"Unexpected formula records detected in {name}")
    return {
        "status": "passed",
        "zip_integrity": "passed",
        "worksheet_xml_parsed": len(worksheet_formulas),
        "worksheet_formula_counts": worksheet_formulas,
    }


def create_contacts_workbook(output: Path, contacts: list[dict[str, Any]], include_mailbox: bool):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "External Contacts"
    headers = ["ContactDomain", "ContactEmail", "ContactType", "FirstEngagement", "LastEngagement", "InboundMessageCount", "OutboundMessageCount", "UnresolvedMessageCount", "TotalUniqueMessages"]
    if include_mailbox:
        headers.append("MailboxesEngagedThrough")
    sheet.append(headers)
    for row_number, item in enumerate(contacts, start=2):
        set_literal_text(sheet.cell(row_number, 1), item["ContactDomain"])
        set_literal_text(sheet.cell(row_number, 2), item["ContactEmail"])
        set_literal_text(sheet.cell(row_number, 3), item["ContactType"])
        sheet.cell(row_number, 4, excel_datetime(item["FirstEngagement"]))
        sheet.cell(row_number, 5, excel_datetime(item["LastEngagement"]))
        sheet.cell(row_number, 6, item["InboundMessageCount"])
        sheet.cell(row_number, 7, item["OutboundMessageCount"])
        sheet.cell(row_number, 8, item["UnresolvedMessageCount"])
        sheet.cell(row_number, 9, item["TotalUniqueMessages"])
        if include_mailbox:
            set_literal_text(sheet.cell(row_number, 10), item["MailboxesEngagedThrough"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    if sheet.max_row > 1:
        last = "J" if include_mailbox else "I"
        table = Table(displayName="ExternalContacts", ref=f"A1:{last}{sheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)
    widths = {"A": 30, "B": 42, "C": 14, "D": 22, "E": 22, "F": 21, "G": 22, "H": 23, "I": 21, "J": 66}
    for column, width in widths.items():
        if column != "J" or include_mailbox:
            sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        row[3].number_format = "dd mmm yyyy hh:mm"
        row[4].number_format = "dd mmm yyyy hh:mm"
    workbook.save(output)
    workbook.close()


def add_native_pivot(output: Path, include_mailbox: bool, log: dict[str, Any]) -> bool:
    import pythoncom
    import win32com.client as win32
    xlDatabase = 1
    xlRowField = 1
    xlPageField = 3
    xlSum = -4157
    xlAscending = 1
    xlDescending = 2
    xlTabularRow = 1
    xlRepeatLabels = 2
    version = 6
    pythoncom.CoInitialize()
    excel = workbook = None
    last_error = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        for attempt in range(1, 4):
            try:
                workbook = excel.Workbooks.Open(str(output.resolve()))
                break
            except Exception as exc:
                last_error = exc
                time.sleep(attempt)
        if workbook is None:
            raise last_error or RuntimeError("Excel could not open the workbook")
        pivot_sheet = workbook.Worksheets.Add(After=workbook.Worksheets.Item(workbook.Worksheets.Count))
        pivot_sheet.Name = "Pivot"
        cache = workbook.PivotCaches().Create(SourceType=xlDatabase, SourceData="ExternalContacts", Version=version)
        pivot = cache.CreatePivotTable(TableDestination="'Pivot'!R3C1", TableName="ExternalContactsPivot", DefaultVersion=version)
        domain = pivot.PivotFields("ContactDomain")
        domain.Orientation = xlRowField
        domain.Position = 1
        domain.Subtotals = (False,) * 12
        email = pivot.PivotFields("ContactEmail")
        email.Orientation = xlRowField
        email.Position = 2
        email.Subtotals = (False,) * 12
        if include_mailbox:
            mailbox = pivot.PivotFields("MailboxesEngagedThrough")
            mailbox.Orientation = xlPageField
            mailbox.Position = 1
        inbound = pivot.AddDataField(pivot.PivotFields("InboundMessageCount"), "Sum of InboundMessageCount", xlSum)
        outbound = pivot.AddDataField(pivot.PivotFields("OutboundMessageCount"), "Sum of OutboundMessageCount", xlSum)
        total = pivot.AddDataField(pivot.PivotFields("TotalUniqueMessages"), "Sum of TotalUniqueMessages", xlSum)
        for field in (inbound, outbound, total):
            field.NumberFormat = "#,##0"
        pivot.RowAxisLayout(xlTabularRow)
        pivot.RepeatAllLabels(xlRepeatLabels)
        pivot.TableStyle2 = "PivotStyleMedium2"
        pivot.ShowTableStyleRowStripes = True
        domain.AutoSort(xlAscending, "ContactDomain")
        email.AutoSort(xlDescending, total.Name)
        pivot.RefreshTable()
        cache.RefreshOnFileOpen = False
        cache.EnableRefresh = True
        pivot.SaveData = True
        pivot_sheet.Columns.AutoFit()
        workbook.Save()
        log["pivot"] = {"status": "created"}
        return True
    except Exception as exc:
        log["pivot"] = {"status": "not created", "error": repr(exc)}
        print(f"WARNING: Native PivotTable could not be created. Data workbook retained. {exc}", file=sys.stderr)
        return False
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def append_interactions_sheet(output: Path, messages: list[dict[str, Any]]):
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    workbook = load_workbook(output, keep_links=True)
    sheet = workbook.create_sheet("Email Interactions")
    headers = ["DateTime", "ContactDomain", "ContactType", "Direction", "From", "To", "Cc", "Bcc", "Subject", "MessageContent", "ContentFormat", "AttachmentNames", "AttachmentCount", "InternetMessageId", "Mailbox", "FolderPath"]
    sheet.append(headers)
    for row_number, message in enumerate(messages, start=2):
        sheet.cell(row_number, 1, excel_datetime(message["Date"]))
        literal_values = {
            2: message["ContactDomain"], 3: message["ContactType"],
            4: message["Direction"], 5: message["From"], 6: message["To"],
            7: message["Cc"], 8: message["Bcc"], 9: message["Subject"],
            10: message["MessageContent"], 11: message["ContentFormat"],
            12: message["AttachmentNames"], 14: message["InternetMessageId"],
            15: mailbox_string(message["SourceMailboxes"]), 16: message["FolderPath"],
        }
        for column, value in literal_values.items():
            set_literal_text(sheet.cell(row_number, column), value)
        sheet.cell(row_number, 13, message["AttachmentCount"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    if sheet.max_row > 1:
        table = Table(displayName="EmailInteractions", ref=f"A1:P{sheet.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sheet.add_table(table)
    widths = {"A": 21, "B": 30, "C": 14, "D": 12, "E": 38, "F": 38, "G": 38, "H": 38, "I": 55, "J": 100, "K": 26, "L": 55, "M": 15, "N": 55, "O": 45, "P": 70}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        row[0].number_format = "dd mmm yyyy hh:mm"
        for cell in row:
            cell.alignment = Alignment(vertical="top")
        for index in [4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15]:
            row[index].alignment = Alignment(vertical="top", wrap_text=True)
    workbook.save(output)
    workbook.close()


def archive_previous_reports(base: Path) -> list[dict[str, str]]:
    archive = base / "Archive"
    archive.mkdir(exist_ok=True)
    moved: list[dict[str, str]] = []
    for folder in sorted(base.glob("External_Contacts_Reports_*")):
        if not folder.is_dir() or folder.parent == archive:
            continue
        destination = archive / folder.name
        counter = 2
        while destination.exists():
            destination = archive / f"{folder.name}_{counter}"
            counter += 1
        shutil.move(str(folder), str(destination))
        moved.append({"from": str(folder), "to": str(destination)})
    return moved


def safe_scope_name(selected: list[dict[str, str]]) -> str:
    if len(selected) == 1:
        text = selected[0]["label"]
    else:
        text = "Combined_" + "_".join(item["label"] for item in selected)
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("_.-")
    return text[:100] or "Selected_Mailboxes"


def format_bytes(n: int) -> str:
    v=float(n)
    for u in ("B","KB","MB","GB","TB"):
        if v<1024 or u=="TB": return f"{v:,.1f} {u}" if u!="B" else f"{int(v):,} B"
        v/=1024

def running_outlook_processes() -> list[str]:
    if sys.platform != "win32": return []
    found=[]
    for image in ("OUTLOOK.EXE","OLK.EXE"):
        try:
            r=subprocess.run(["tasklist","/FI",f"IMAGENAME eq {image}","/NH"],capture_output=True,text=True,check=False,creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
            if image in r.stdout.upper(): found.append(image)
        except Exception: pass
    return found

def discover_profile_ost_files(root: Path=Path("C:/Users")) -> list[dict[str,Any]]:
    found=[]
    if not root.is_dir(): return found
    for profile in sorted(root.iterdir(),key=lambda x:x.name.lower()):
        folder=profile/"AppData"/"Local"/"Microsoft"/"Outlook"
        try: files=list(folder.glob("*.ost")) if folder.is_dir() else []
        except (PermissionError,OSError): continue
        for path in files:
            try:
                st=path.stat(); found.append({"profile":profile.name,"path":path.resolve(),"size":st.st_size,"modified":datetime.fromtimestamp(st.st_mtime)})
            except (PermissionError,OSError): pass
    return sorted(found,key=lambda x:(-x["size"],x["profile"].lower(),x["path"].name.lower()))

def choose_source_ost(items: list[dict[str,Any]]) -> Path|None:
    if not items: print("No accessible OST files found under C:\\Users."); return None
    largest=items[0]["size"]
    print("\nOutlook email files found (largest first):")
    for i,x in enumerate(items,1):
        rec="  [RECOMMENDED: largest file]" if x["size"]==largest else ""
        print(f"  {i}. Profile: {x['profile']}\n     File: {x['path'].name}\n     Size: {format_bytes(x['size'])}\n     Modified: {x['modified']:%d %b %Y %H:%M}{rec}")
    while True:
        raw=input("Select an OST file [1 recommended], or C to cancel: ").strip()
        if not raw: return items[0]["path"]
        if raw.lower() in {"c","cancel"}: return None
        if raw.isdigit() and 1<=int(raw)<=len(items): return items[int(raw)-1]["path"]
        print("Enter a listed number, Enter for the recommendation, or C.")

def wait_for_outlook() -> bool:
    while (procs:=running_outlook_processes()):
        print(f"\nOutlook is running ({', '.join(procs)}). Close it before copying the OST.")
        if input("Press R to recheck, or C to cancel refresh: ").strip().lower() in {"c","cancel"}: return False
    return True

def refresh_working_ost(src: Path, base: Path) -> tuple[Path,dict[str,Any]]:
    src=src.resolve(); st=src.stat(); folder=base/"Source_OST"; folder.mkdir(exist_ok=True)
    dst=folder/src.name; partial=dst.with_name(dst.name+".partial"); previous=dst.with_name(dst.name+".previous")
    if src==dst.resolve(): return src,{"status":"already working copy","working_copy":str(src)}
    required=st.st_size+1024**3; free=shutil.disk_usage(folder).free
    if free<required: raise OSError(f"Insufficient disk space: need {format_bytes(required)}, available {format_bytes(free)}")
    partial.unlink(missing_ok=True); method="Windows CopyFile2"
    print(f"\nCopying {format_bytes(st.st_size)} from:\n  {src}\nTo working copy:\n  {dst}")
    try: shutil.copy2(src,partial)
    except PermissionError as e1:
        partial.unlink(missing_ok=True); method="read-only streamed copy"; print("File lock detected; trying read-only streamed copy...")
        try:
            with src.open("rb",buffering=0) as fi, partial.open("xb",buffering=0) as fo:
                while chunk:=fi.read(8*1024*1024): fo.write(chunk)
            shutil.copystat(src,partial)
        except (PermissionError,OSError) as e2:
            partial.unlink(missing_ok=True)
            raise RuntimeError(f"OST remains locked. Close Outlook, wait for shutdown, and retry. Copy errors: {e1}; {e2}") from e2
    copied=partial.stat()
    if copied.st_size!=st.st_size: partial.unlink(missing_ok=True); raise OSError("Copy validation failed: file sizes differ")
    previous.unlink(missing_ok=True)
    if dst.exists(): dst.replace(previous)
    try: partial.replace(dst)
    except Exception:
        if previous.exists() and not dst.exists(): previous.replace(dst)
        raise
    previous.unlink(missing_ok=True); print("OST working copy updated; source was not modified.")
    return dst,{"status":"copied and size validated","source":str(src),"working_copy":str(dst),"source_size":st.st_size,"copied_size":copied.st_size,"copy_method":method,"copy_completed_utc":datetime.now(timezone.utc).isoformat()}

def prepare_source_ost(explicit: str|None, base: Path, skip: bool=False) -> tuple[Path,dict[str,Any]]:
    refresh=False
    if not skip:
        while True:
            a=input("Do you want to update the source email file? [Y/N]: ").strip().lower()
            if a in {"y","yes"}: refresh=True; break
            if a in {"n","no",""}: break
            print("Please enter Y or N.")
    if refresh:
        selected=choose_source_ost(discover_profile_ost_files())
        if selected and wait_for_outlook(): return refresh_working_ost(selected,base)
        print("Refresh cancelled; using an existing working copy.")
    if explicit:
        path=Path(explicit).expanduser()
        if not path.is_file(): raise FileNotFoundError(path)
        return path.resolve(),{"status":"explicit path used","working_copy":str(path.resolve())}
    folder=base/"Source_OST"; files=list(folder.glob("*.ost")) if folder.is_dir() else []
    if not files: files=list(base.glob("*.ost"))
    if not files: raise FileNotFoundError("No working OST found in Source_OST or beside the script")
    files.sort(key=lambda x:x.stat().st_size,reverse=True)
    return files[0].resolve(),{"status":"existing working copy used","working_copy":str(files[0].resolve()),"selection_reason":"largest available working OST"}


def main() -> int:
    parser = argparse.ArgumentParser(description="Interactive OST external-engagement report")
    parser.add_argument("--ost")
    parser.add_argument("--output")
    parser.add_argument("--internal-domain", action="append", default=[])
    parser.add_argument("--include-automated", action="store_true")
    parser.add_argument("--skip-source-prompt", action="store_true")
    args = parser.parse_args()
    try:
        import pypff
    except ImportError:
        print("Install into this interpreter: python -m pip install libpff-python-windows", file=sys.stderr)
        return 2
    try:
        import openpyxl  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        print("Install into this interpreter: python -m pip install openpyxl pywin32", file=sys.stderr)
        return 2

    script_folder = Path(__file__).resolve().parent
    archived = archive_previous_reports(script_folder)
    if archived:
        print(f"Archived {len(archived)} previous report folder(s) to: {script_folder / 'Archive'}")

    try:
        ost, ost_refresh = prepare_source_ost(args.ost, script_folder, args.skip_source_prompt)
    except (OSError, RuntimeError) as exc:
        print(f"\nSOURCE OST ERROR: {exc}", file=sys.stderr)
        print("The source OST was not changed and no partial copy was retained.", file=sys.stderr)
        return 3
    pff = pypff.file()
    try:
        pff.open(str(ost))
        root = pff.get_root_folder()
        internal_domains = {
            normalise_address(value).lstrip("@")
            for value in (args.internal_domain or INTERNAL_DOMAINS_DEFAULT)
        }
        print("Inspecting mailbox headers to identify SMTP addresses...")
        detected = discover_mailboxes(root, internal_domains)
        selected = choose_mailboxes(detected)
        address_scope = choose_address_scope()
        own_addresses = {
            normalise_address(item.get("email", ""))
            for item in selected
            if item.get("email")
        }
        print("\nSelected report scope:")
        for item in selected:
            print(f"  - {item['label']}")
        print(f"Address scope: {address_scope}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        parent = Path(args.output).resolve() if args.output else script_folder
        final_folder = parent / f"External_Contacts_Reports_{timestamp}"
        staging_root = Path(tempfile.mkdtemp(prefix="ost_external_report_"))
        staging_folder = staging_root / final_folder.name
        staging_folder.mkdir(parents=True)
        scope_name = safe_scope_name(selected)
        workbook_path = staging_folder / f"External_Contacts_{scope_name}_{timestamp}.xlsx"
        log: dict[str, Any] = {
            "started_utc": datetime.now(timezone.utc).isoformat(),
            "source": str(ost),
            "ost_refresh": ost_refresh,
            "detected_mailboxes": detected,
            "selected_mailboxes": selected,
            "address_scope": address_scope,
            "excluded_selected_mailbox_addresses": sorted(own_addresses),
            "archived_previous_reports": archived,
            "folders_seen": 0,
            "items_seen": 0,
            "message_errors": 0,
            "folder_errors": 0,
            "duplicates_removed": 0,
            "error_examples": [],
        }
        by_key: dict[str, dict[str, Any]] = {}
        for message in traverse_selected(root, [], selected, log, internal_domains, not args.include_automated, address_scope, own_addresses):
            key = message["DuplicateKey"]
            if key in by_key:
                log["duplicates_removed"] += 1
                existing = by_key[key]
                existing["SourceMailboxes"].update(message["SourceMailboxes"])
                existing["FolderPath"] = "; ".join(sorted(set(existing["FolderPath"].split("; ") + [message["FolderPath"]])))
                existing_domains = {value.strip() for value in existing["ContactDomain"].split(";") if value.strip()}
                existing_domains.update(value.strip() for value in message["ContactDomain"].split(";") if value.strip())
                existing["ContactDomain"] = "; ".join(sorted(existing_domains))
                existing_types = {value.strip() for value in existing["ContactType"].split(";") if value.strip()}
                existing_types.update(value.strip() for value in message["ContactType"].split(";") if value.strip())
                existing["ContactType"] = "; ".join(sorted(existing_types))
                if not existing["MessageContent"] and message["MessageContent"]:
                    existing["MessageContent"] = message["MessageContent"]
                    existing["ContentFormat"] = message["ContentFormat"]
                if not existing["AttachmentNames"] and message["AttachmentNames"]:
                    existing["AttachmentNames"] = message["AttachmentNames"]
                    existing["AttachmentCount"] = message["AttachmentCount"]
            else:
                by_key[key] = message
        messages = sorted(by_key.values(), key=lambda item: (item["Date"], item["InternetMessageId"], item["Subject"]), reverse=True)
        contacts = build_contacts(messages)
        include_mailbox = len(selected) > 1
        create_contacts_workbook(workbook_path, contacts, include_mailbox)
        pivot_created = add_native_pivot(workbook_path, include_mailbox, log)
        append_interactions_sheet(workbook_path, messages)
        workbook_validation = validate_xlsx_package(workbook_path)
        log.update({
            "workbook_validation": workbook_validation,
            "completed_utc": datetime.now(timezone.utc).isoformat(),
            "external_contacts": len(contacts),
            "interaction_rows": len(messages),
            "pivot_created_before_interactions": pivot_created,
        })
        (staging_folder / "processing_log.json").write_text(json.dumps(log, indent=2), encoding="utf-8")
        if final_folder.exists():
            raise FileExistsError(final_folder)
        shutil.move(str(staging_folder), str(final_folder))
        shutil.rmtree(staging_root, ignore_errors=True)
    finally:
        try:
            pff.close()
        except Exception:
            pass

    print(f"\nDone: {final_folder}")
    print(f"External contacts: {len(contacts):,}")
    print(f"Interaction rows: {len(messages):,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
