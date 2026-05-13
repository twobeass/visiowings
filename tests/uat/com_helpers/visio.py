"""Visio COM context manager — guarantees cleanup, including zombie kill."""

from __future__ import annotations

from .process import kill_zombies


class VBOMAccessDeniedError(RuntimeError):
    """Raised when Trust Center 'AccessVBOM' is not effective.

    Bootstrap detects this and prompts the user to start Visio once manually
    to acknowledge the trust settings; retrying after that usually succeeds.
    """


class VisioContext:
    """Headless Visio.Application instance with auto-cleanup.

    Usage::

        with VisioContext() as app:
            doc = app.Documents.Add("")
            ...
    """

    def __init__(self, visible: bool = False, alert_response: int = 7, shared: bool = False):
        self.visible = visible
        self.alert_response = alert_response
        self.shared = shared
        self.app = None

    def __enter__(self):
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom.CoInitialize()
        self._co_inited = True
        if self.shared:
            self.app = win32com.client.Dispatch("Visio.Application")
        else:
            self.app = win32com.client.DispatchEx("Visio.Application")
        try:
            self.app.Visible = self.visible
        except Exception:
            pass
        try:
            self.app.AlertResponse = self.alert_response
        except Exception:
            pass
        return self.app

    def __exit__(self, exc_type, exc, tb):
        """Forceful cleanup.

        We deliberately do NOT wait on ``app.Quit()`` — Visio occasionally
        blocks on a hidden "save changes?" dialog despite ``AlertResponse=7``
        and that hangs the test under pytest-timeout. Instead: try to close
        documents, fire-and-forget Quit, then terminate the process.
        """
        import threading

        import pythoncom  # type: ignore

        def _attempt_graceful():
            try:
                if self.app is not None:
                    try:
                        for doc in list(self.app.Documents):
                            try:
                                doc.Saved = True
                            except Exception:
                                pass
                            try:
                                doc.Close()
                            except Exception:
                                pass
                    except Exception:
                        pass
                    try:
                        self.app.Quit()
                    except Exception:
                        pass
            except Exception:
                pass

        t = threading.Thread(target=_attempt_graceful, daemon=True)
        t.start()
        t.join(timeout=5.0)

        try:
            kill_zombies(["visio.exe"])
        finally:
            if getattr(self, "_co_inited", False):
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        return False
