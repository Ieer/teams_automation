"""Automation helpers for Microsoft Teams."""

from importlib.metadata import PackageNotFoundError, version as _importlib_version

from .client import TeamsClient

try:
	__version__ = _importlib_version(__name__)
except PackageNotFoundError:  # pragma: no cover - fallback only during local dev
	__version__ = "0.1.0"

__all__ = ["TeamsClient", "__version__"]
