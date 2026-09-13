"""Compose an email draft via Outlook (Windows COM automation), with
attachments -- used by the per-client "Prepare Email" buttons in the
Create (EXCEL+Biglietti) page. Always opens a review-before-send compose
window (.Display()); this never sends an email automatically.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import subprocess
import tkinter as tk
from tkinter import ttk


EMAIL_RE = re.compile(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", re.IGNORECASE)


def extract_email_addresses(*values: str) -> list[str]:
    """Return unique email addresses from one or more recipient fields."""
    found: list[str] = []
    seen: set[str] = set()
    for value in values:
        for address in EMAIL_RE.findall(value or ""):
            address = address.strip()
            key = address.casefold()
            if key not in seen:
                seen.add(key)
                found.append(address)
    return found


class EmailAutocomplete(ttk.Combobox):
    """Editable recipient field with case-insensitive prefix suggestions."""

    def __init__(self, master, values: list[str] | tuple[str, ...] = (), **kwargs):
        self._email_values = list(values)
        self._typed_prefix = ""
        super().__init__(master, values=self._email_values, **kwargs)
        self.bind("<KeyRelease>", self._filter_suggestions, add="+")
        self.bind("<<ComboboxSelected>>", self._use_suggestion, add="+")

    def set_email_values(self, values: list[str] | tuple[str, ...]) -> None:
        self._email_values = list(dict.fromkeys(values))
        self["values"] = self._email_values

    def _filter_suggestions(self, _event=None) -> None:
        # Suggest against the part currently being typed after the last ';'.
        current = self.get()
        separator = current.rfind(";")
        self._typed_prefix = current[:separator + 1] + " " if separator >= 0 else ""
        token = current[separator + 1:].strip().casefold()
        matches = [v for v in self._email_values if v.casefold().startswith(token)] if token else self._email_values
        self["values"] = matches
        if matches and token:
            self.after_idle(self.event_generate, "<Down>")

    def _use_suggestion(self, _event=None) -> None:
        selected = self.get().strip()
        self.set(f"{self._typed_prefix}{selected}")


def _outlook_application():
    """Return the running Outlook instance, or start one if necessary."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("This feature needs Microsoft Outlook installed.") from exc
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            return win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"Could not open Outlook: {exc}") from exc


def get_outlook_accounts() -> list[tuple[str, str]]:
    """List connected Outlook accounts as ``(smtp_address, display_name)``."""
    outlook = _outlook_application()
    accounts: list[tuple[str, str]] = []
    for account in outlook.Session.Accounts:
        address = str(getattr(account, "SmtpAddress", "") or "").strip()
        name = str(getattr(account, "DisplayName", "") or address).strip()
        if address and not any(address.casefold() == item[0].casefold() for item in accounts):
            accounts.append((address, name))
    return accounts


def open_outlook_account_setup() -> None:
    """Open Outlook's account/profile setup so the user can add an account."""
    try:
        # The Mail control-panel applet opens even when Outlook is already
        # running.  Launching ``outlook.exe /manageprofiles`` first can only
        # activate the existing Outlook window and show no add-account UI.
        subprocess.Popen(["rundll32.exe", "shell32.dll,Control_RunDLL", "mlcfg32.cpl"])
    except OSError:
        try:
            subprocess.Popen(["outlook.exe", "/manageprofiles"])
        except OSError as exc:
            raise RuntimeError(
                "Could not open Outlook account setup. Open Outlook manually, "
                "then choose File > Add Account."
            ) from exc


@dataclass
class EmailTemplate:
    to: str = ""
    cc: str = ""
    subject: str = ""
    body: str = ""

    def is_blank(self) -> bool:
        return not (self.to.strip() or self.cc.strip() or self.subject.strip() or self.body.strip())

    def to_dict(self) -> dict:
        return {"to": self.to, "cc": self.cc, "subject": self.subject, "body": self.body}

    @classmethod
    def from_dict(cls, data: dict | None) -> "EmailTemplate":
        data = data or {}
        return cls(
            to=str(data.get("to", "")), cc=str(data.get("cc", "")),
            subject=str(data.get("subject", "")), body=str(data.get("body", "")),
        )


def open_outlook_email(template: EmailTemplate, attachments: list[Path], sender_email: str = "") -> None:
    """Open a new Outlook compose window pre-filled from *template*, with
    every existing path in *attachments* attached, and hand control back to
    the user (Outlook's .Display(), never .Send()) so they can review
    before sending. Raises RuntimeError with a clear, user-facing message
    if Outlook automation isn't available -- never lets the raw COM
    exception surface.
    """
    # Prefer attaching to an already-running, already-signed-in Outlook
    # (the normal case: the person already has Outlook open on their PC).
    # win32com.client.Dispatch() can instead spin up a brand-new Outlook
    # process, and if that process has no default profile configured yet
    # it walks straight into Outlook's first-run "Welcome"/"Add an Email
    # Account" wizard instead of just composing the message. GetActiveObject
    # attaches to the running instance's existing, already-logged-in
    # session, so it never triggers that wizard.
    try:
        outlook = _outlook_application()
    except RuntimeError as exc:
        raise RuntimeError(f"Could not open Outlook to prepare the email: {exc}") from exc

    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = template.to
        mail.CC = template.cc
        mail.Subject = template.subject
        mail.Body = template.body
        if sender_email.strip():
            selected = next(
                (account for account in outlook.Session.Accounts
                 if str(getattr(account, "SmtpAddress", "")).strip().casefold() == sender_email.strip().casefold()),
                None,
            )
            if selected is None:
                raise RuntimeError(f"The Outlook account '{sender_email}' is not connected on this computer.")
            mail.SendUsingAccount = selected
        for path in attachments:
            if path and Path(path).is_file():
                mail.Attachments.Add(str(Path(path).resolve()))
        mail.Display()
    except RuntimeError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            f"Could not open Outlook to prepare the email: {exc}\n\n"
            "If Outlook itself shows an account-setup wizard, Outlook has no "
            "mail profile configured yet on this PC -- open Outlook once on "
            "its own and finish adding your email account there; after that "
            "this button will attach to it directly."
        ) from exc
