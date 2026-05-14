"""services package.

This file intentionally keeps package initialization lightweight.
Do not import heavy modules here, so GitHub Actions / PyInstaller / app startup
will not trigger teacher strategy or KPattern pipeline during package import.
"""

__all__ = []
