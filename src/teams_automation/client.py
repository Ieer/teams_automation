"""High-level controls for Microsoft Teams automation."""

from __future__ import annotations

import os
import re
import time
import ctypes
from contextlib import contextmanager
from io import BytesIO
from typing import Dict, Iterable, Iterator, Optional, Sequence, Tuple

import uiautomation as auto
from PIL import Image
import win32clipboard as clipboard
import win32con

kernel32 = ctypes.windll.kernel32
kernel32.GlobalAlloc.argtypes = (ctypes.c_uint, ctypes.c_size_t)
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = (ctypes.c_void_p,)
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = (ctypes.c_void_p,)
kernel32.GlobalUnlock.restype = ctypes.c_bool
kernel32.GlobalFree.argtypes = (ctypes.c_void_p,)
kernel32.GlobalFree.restype = ctypes.c_void_p

CFSTR_PREFERREDDROPEFFECT = clipboard.RegisterClipboardFormat("Preferred DropEffect")
DROPEFFECT_COPY = 1

__all__ = ["TeamsClient", "TeamsAutomationError"]


_PATTERN_LAST_MSG = re.compile(r"Last message.*", re.IGNORECASE)
_PATTERN_LEADING = re.compile(r"^(?:Group|Chat)[:\-\u2013\u2014]?\s*", re.IGNORECASE)
_PATTERN_STATUS = re.compile(r"\b(?:Available|Busy|Away|Offline)\b", re.IGNORECASE)
_PATTERN_TIME = re.compile(r"\b\d{1,2}:\d{2}\s*(?:AM|PM)?\b.*$", re.IGNORECASE)
_PATTERN_WHITESPACE = re.compile(r"[\s\u00A0]+")


class _DROPFILES(ctypes.Structure):
    _fields_ = [
        ("pFiles", ctypes.c_uint32),
        ("x", ctypes.c_int32),
        ("y", ctypes.c_int32),
        ("fNC", ctypes.c_int32),
        ("fWide", ctypes.c_int32),
    ]


class TeamsAutomationError(RuntimeError):
    """Raised when a required Teams UI element cannot be resolved."""


class TeamsClient:
    """A thin wrapper around uiautomation primitives for Microsoft Teams."""

    _DEFAULT_WINDOW_KEYWORD = "Microsoft Teams"
    _DEFAULT_WINDOW_CLASS = "TeamsWebView"
    _DEFAULT_SECTION_NAMES = ("Favorites", "Chats")
    _MESSAGE_FIELD_NAMES = ("Type a message", "Type your message")
    _SEND_BUTTON_NAMES = ("Send (Ctrl+Enter)", "Send")
    _IMAGE_EXTENSIONS = {
        ".png",
        ".jpg",
        ".jpeg",
        ".bmp",
        ".gif",
        ".tif",
        ".tiff",
        ".webp",
    }
    _CLIPBOARD_OPEN_RETRIES = 5
    _CLIPBOARD_RETRY_DELAY = 0.1

    def __init__(
        self,
        *,
        window_keyword: str | Iterable[str] = _DEFAULT_WINDOW_KEYWORD,
        window_class_keyword: str = _DEFAULT_WINDOW_CLASS,
        activation_delay: float = 3.0,
        search_timeout: float = 5.0,
        section_preference: Iterable[str] | None = None,
        aliases: Optional[Dict[str, str]] = None,
        minimize_after_send: bool = False,
    ) -> None:
        self.window_keyword = self._ensure_iterable(window_keyword)
        self.window_class_keyword = window_class_keyword
        self.activation_delay = activation_delay
        self.search_timeout = search_timeout
        self.section_preference = tuple(section_preference or self._DEFAULT_SECTION_NAMES)
        self.aliases = aliases or {"Teams Chatbot Bot": "Columbus Teams Chatbot"}
        self._alias_lookup = {key.lower(): value for key, value in self.aliases.items()}
        self.minimize_after_send = minimize_after_send
        self._window: Optional[auto.Control] = None

    @staticmethod
    def _ensure_iterable(value: str | Iterable[str]) -> Tuple[str, ...]:
        if isinstance(value, str):
            return (value,)
        return tuple(value)

    def connect(self) -> auto.Control:
        """Resolve and activate the Teams window."""
        if self._window and self._window.Exists(0, 0):
            return self._window
        root = auto.GetRootControl()
        target_window = None
        for candidate in root.GetChildren():
            if not any(keyword in candidate.Name for keyword in self.window_keyword):
                continue
            if self.window_class_keyword and self.window_class_keyword not in candidate.ClassName:
                continue
            candidate.SetActive()  # type: ignore[attr-defined]
            target_window = auto.WindowControl(searchDepth=1, Name=candidate.Name)
            break
        if target_window is None or not target_window.Exists(0, 0):
            raise TeamsAutomationError("Unable to locate an active Microsoft Teams window.")
        if self.activation_delay:
            time.sleep(self.activation_delay)
        self._window = target_window
        return target_window

    @property
    def window(self) -> auto.Control:
        """Cached window control access."""
        if self._window is None or not self._window.Exists(0, 0):
            return self.connect()
        return self._window

    def send_text(
        self,
        message: str,
        *,
        chat_name: str,
        section_name: Optional[str] = None,
        close_filter: bool = True,
        wait_after_send: float = 3.0,
    ) -> None:
        """Send a plain text message to a Teams chat."""
        self.send_message(
            message,
            chat_name=chat_name,
            section_name=section_name,
            close_filter=close_filter,
            wait_after_send=wait_after_send,
        )

    def send_message(
        self,
        message: str,
        *,
        chat_name: str,
        section_name: Optional[str] = None,
        image_path: Optional[str] = None,
        close_filter: bool = True,
        wait_after_send: float = 3.0,
    ) -> None:
        """Send a text (and optional image) to a Teams chat."""
        window = self.window
        self._activate_chat(window, chat_name, section_name, close_filter)
        self._enter_message(window, message)
        if image_path:
            self._append_image(window, image_path)
        self._trigger_send(window)
        if wait_after_send:
            time.sleep(wait_after_send)
        if self.minimize_after_send:
            auto.SendKeys("{win}d")

    def send_files(
        self,
        filepaths: Sequence[str],
        *,
        chat_name: str,
        section_name: Optional[str] = None,
        caption: Optional[str] = None,
        embed_images: bool = True,
        close_filter: bool = True,
        wait_after_send: float = 3.0,
    ) -> None:
        """Send one or more files to a Teams chat."""
        window = self.window
        self._activate_chat(window, chat_name, section_name, close_filter)
        resolved_paths = self._validate_files(filepaths)
        message_field = self._focus_message_field(window)
        message_field.SendKeys("{Ctrl}a{Del}")
        if caption:
            message_field.SendKeys(caption)
        image_paths: list[str] = []
        file_paths: list[str] = []
        for path in resolved_paths:
            if embed_images and self._is_image(path):
                image_paths.append(path)
            else:
                file_paths.append(path)
        if image_paths:
            for image_path in image_paths:
                self._append_image(window, image_path)
        if file_paths:
            self._load_files_to_clipboard(file_paths)
            message_field.SetFocus()
            message_field.SendKeys("{Ctrl}v")
            time.sleep(2.0)
        self._trigger_send(window)
        if wait_after_send:
            time.sleep(wait_after_send)
        if self.minimize_after_send:
            auto.SendKeys("{win}d")

    def _open_chat_hub(self, window: auto.Control) -> None:
        chat_button = window.Control(searchDepth=30, Name="Chat (Ctrl+2)")
        if not chat_button.Exists(0, 0):
            raise TeamsAutomationError("Chat hub button not found.")
        chat_button.Click()
        window.SetActive()  # type: ignore[attr-defined]

    def _apply_filter(self, window: auto.Control) -> None:
        filter_button = window.ButtonControl(searchDepth=30, Name="Show filter text box (Ctrl+Shift+F)")
        if not filter_button.Exists(0, 0):
            raise TeamsAutomationError("Filter text box trigger not found.")
        filter_button.Click()

    def _filter_search(self, window: auto.Control, term: str) -> None:
        search_field = window.EditControl(searchDepth=30, Name="Filter by name or group name")
        if not search_field.Exists(0, 0):
            raise TeamsAutomationError("Filter search field not found.")
        search_field.SetFocus()
        search_field.SendKeys("{Ctrl}a{Del}")
        search_field.SendKeys(term)
        time.sleep(self.search_timeout)

    def _find_chat_entry(
        self,
        window: auto.Control,
        display_name: str,
        section_name: Optional[str],
    ) -> auto.Control:
        normalized_target = self._normalize_name(display_name)
        sections = self._collect_sections(window)
        ordered_sections: Tuple[str, ...]
        if section_name:
            if section_name not in sections:
                available = ", ".join(sorted(sections))
                raise TeamsAutomationError(
                    f"Section '{section_name}' not available. Found: {available}"
                )
            ordered_sections = (section_name,) + tuple(
                key for key in sections if key != section_name
            )
        else:
            ordered_sections = tuple(
                key for key in self.section_preference if key in sections
            ) + tuple(key for key in sections if key not in self.section_preference)

        inspected: list[str] = []
        for current_section in ordered_sections:
            candidates = sections.get(current_section, ())
            for candidate in candidates:
                friendly_name = self._normalize_name(candidate.Name)
                if self._names_match(friendly_name, normalized_target):
                    return candidate
                inspected.append(f"{current_section}:{friendly_name}")
        inspected_text = ", ".join(inspected) or "<none>"
        raise TeamsAutomationError(
            f"Chat '{display_name}' not located. Inspected entries: {inspected_text}"
        )

    def _collect_sections(self, window: auto.Control) -> Dict[str, Tuple[auto.Control, ...]]:
        container = window.Control(searchDepth=30, Name="Filter active")
        if not container.Exists(0, 0):
            raise TeamsAutomationError("Active filter list not found.")
        sections: Dict[str, Tuple[auto.Control, ...]] = {}
        for section in container.GetChildren():
            if not section.Name:
                continue
            entries = []
            for group in section.GetChildren():
                for entry in group.GetChildren():
                    if not entry.Name:
                        continue
                    entries.append(entry)
            if entries:
                sections[section.Name] = tuple(entries)
        return sections

    def _close_filter(self, window: auto.Control) -> None:
        close_button = window.ButtonControl(searchDepth=30, Name="Close filter text box")
        if close_button.Exists(0, 0):
            close_button.Click()
            time.sleep(0.5)

    def _focus_message_field(self, window: auto.Control) -> auto.Control:
        for field_name in self._MESSAGE_FIELD_NAMES:
            candidate = window.EditControl(searchDepth=30, Name=field_name)
            if candidate.Exists(0, 0):
                candidate.SetFocus()
                return candidate
        raise TeamsAutomationError("Message input field not found.")

    def _enter_message(self, window: auto.Control, message: str) -> auto.Control:
        message_field = self._focus_message_field(window)
        message_field.SendKeys("{Ctrl}a{Del}")
        message_field.SendKeys(message)
        return message_field

    def _append_image(self, window: auto.Control, image_path: str) -> None:
        resolved_path = os.path.realpath(image_path)
        if not os.path.exists(resolved_path):
            raise TeamsAutomationError(f"Image path not found: {resolved_path}")
        self._load_image_to_clipboard(resolved_path)
        message_field = self._focus_message_field(window)
        message_field.SetFocus()
        message_field.SendKeys("{Ctrl}v")
        time.sleep(2.0)

    def _trigger_send(self, window: auto.Control) -> None:
        send_button = None
        for button_name in self._SEND_BUTTON_NAMES:
            candidate = window.ButtonControl(searchDepth=30, Name=button_name)
            if candidate.Exists(0, 0):
                send_button = candidate
                break
        if send_button is None:
            raise TeamsAutomationError("Send button not found.")
        send_button.Click()

    def _load_image_to_clipboard(self, path: str) -> None:
        with Image.open(path) as img:
            with BytesIO() as output:
                img.convert("RGB").save(output, format="BMP")
                bmp_data = output.getvalue()[14:]
        with self._clipboard_session():
            clipboard.EmptyClipboard()
            clipboard.SetClipboardData(win32con.CF_DIB, bmp_data)

    def _load_files_to_clipboard(self, paths: Sequence[str]) -> None:
        if not paths:
            raise TeamsAutomationError("No file paths provided.")
        file_str = "\0".join(paths) + "\0\0"
        file_bytes = file_str.encode("utf-16le")
        total_size = ctypes.sizeof(_DROPFILES) + len(file_bytes)
        handle = kernel32.GlobalAlloc(win32con.GHND, total_size)
        if not handle:
            raise TeamsAutomationError("Failed to allocate clipboard memory for files.")
        ptr = kernel32.GlobalLock(handle)
        ptr_value = int(ptr) if ptr else 0
        if not ptr_value:
            kernel32.GlobalFree(handle)
            raise TeamsAutomationError("Failed to lock clipboard memory for files.")
        try:
            drop = _DROPFILES.from_address(ptr_value)
            drop.pFiles = ctypes.sizeof(_DROPFILES)
            drop.x = drop.y = 0
            drop.fNC = 0
            drop.fWide = 1
            ctypes.memmove(ptr_value + ctypes.sizeof(_DROPFILES), file_bytes, len(file_bytes))
        finally:
            kernel32.GlobalUnlock(handle)
        drop_effect_handle = kernel32.GlobalAlloc(
            win32con.GHND, ctypes.sizeof(ctypes.c_uint32)
        )
        if not drop_effect_handle:
            kernel32.GlobalFree(handle)
            raise TeamsAutomationError("Failed to allocate clipboard memory for drop effect.")
        effect_ptr = kernel32.GlobalLock(drop_effect_handle)
        effect_ptr_value = int(effect_ptr) if effect_ptr else 0
        if not effect_ptr_value:
            kernel32.GlobalFree(drop_effect_handle)
            kernel32.GlobalFree(handle)
            raise TeamsAutomationError("Failed to lock clipboard memory for drop effect.")
        try:
            ctypes.c_uint32.from_address(effect_ptr_value).value = DROPEFFECT_COPY
        finally:
            kernel32.GlobalUnlock(drop_effect_handle)
        drop_handle_owned = True
        effect_handle_owned = True
        try:
            with self._clipboard_session():
                clipboard.EmptyClipboard()
                clipboard.SetClipboardData(win32con.CF_HDROP, int(handle))
                drop_handle_owned = False
                clipboard.SetClipboardData(CFSTR_PREFERREDDROPEFFECT, int(drop_effect_handle))
                effect_handle_owned = False
        finally:
            if drop_handle_owned:
                kernel32.GlobalFree(handle)
            if effect_handle_owned:
                kernel32.GlobalFree(drop_effect_handle)

    def _normalize_name(self, raw_name: str) -> str:
        name = raw_name or ""
        name = _PATTERN_LAST_MSG.sub("", name)
        name = _PATTERN_LEADING.sub("", name)
        name = _PATTERN_TIME.sub("", name)
        name = _PATTERN_STATUS.sub("", name)
        name = _PATTERN_WHITESPACE.sub(" ", name).strip()
        lower_name = name.lower()
        for alias_key, alias_value in self._alias_lookup.items():
            if lower_name == alias_key or alias_key in lower_name:
                return alias_value
        return name

    def _names_match(self, candidate: str, target: str) -> bool:
        if not candidate or not target:
            return False
        cand = candidate.lower()
        tgt = target.lower()
        if cand == tgt:
            return True
        if tgt in cand or cand in tgt:
            return True
        return False

    @classmethod
    def _is_image(cls, path: str) -> bool:
        _, ext = os.path.splitext(path)
        return ext.lower() in cls._IMAGE_EXTENSIONS

    def _activate_chat(
        self,
        window: auto.Control,
        chat_name: str,
        section_name: Optional[str],
        close_filter: bool,
    ) -> None:
        self._open_chat_hub(window)
        self._apply_filter(window)
        self._filter_search(window, chat_name)
        target_control = self._find_chat_entry(window, chat_name, section_name)
        target_control.Click()
        time.sleep(1.0)
        if close_filter:
            self._close_filter(window)

    def _validate_files(self, filepaths: Sequence[str]) -> Tuple[str, ...]:
        resolved: list[str] = []
        for path in filepaths:
            resolved_path = os.path.realpath(path)
            if not os.path.exists(resolved_path):
                raise TeamsAutomationError(f"File path not found: {resolved_path}")
            if not os.path.isfile(resolved_path):
                raise TeamsAutomationError(f"Path is not a file: {resolved_path}")
            resolved.append(resolved_path)
        return tuple(resolved)

    @contextmanager
    def _clipboard_session(self) -> Iterator[None]:
        last_error: Optional[Exception] = None
        for _ in range(self._CLIPBOARD_OPEN_RETRIES):
            try:
                clipboard.OpenClipboard()
                break
            except OSError as exc:
                last_error = exc
                time.sleep(self._CLIPBOARD_RETRY_DELAY)
        else:
            raise TeamsAutomationError("Unable to access clipboard; it may be locked by another process.") from last_error
        try:
            yield
        finally:
            clipboard.CloseClipboard()
