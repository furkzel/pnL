"""
xlauto.py - Excel COM automation helpers for the Aktarim Paneli project.

Used by build.py, export.py and test_e2e.py. Windows only (pywin32).

Key pieces
----------
ExcelSession   : starts a private, invisible Excel instance (never touches
                 Excel windows the user already has open) and guarantees it
                 is closed / killed afterwards.
DialogWatcher  : background thread that finds modal dialogs (MsgBox,
                 InputBox, VBA compile errors) belonging to OUR Excel process,
                 records their text and answers them according to a policy.
                 Without it any MsgBox in the VBA code would block automation.
inject_sources : replaces the VBA project's modules with the files in src/.
compile_project: runs VBE "Compile VBAProject" and reports compile errors.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import pythoncom
import win32api
import win32con
import win32gui
import win32process
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"

# Standard / class modules that belong to this project (or to v33) and are
# replaced on every build.
LEGACY_MODULES = {"Module1", "Module2", "Module3"}
PROJECT_PREFIXES = ("M0",)
PROJECT_CLASSES = {"Sozluk"}
# v32/v33 Mac test harness found in old workbooks ("SADECE TEST ICIN ...
# Teslim edilen koda GIRMEZ"). They call v33 procedures that no longer
# exist, so the project would not compile with them. Removed on build and
# reported; pass keep=... to inject_sources to keep one.
LEGACY_TEST_MODULES = {"TestOrtam", "OzetTest", "Yukleyici", "Resim2"}

VBEXT_STD = 1
VBEXT_CLASS = 2
VBEXT_DOCUMENT = 100
XL_OPENXML_MACRO = 52


# ---------------------------------------------------------------------------
#  Dialog watcher
# ---------------------------------------------------------------------------
@dataclass
class DialogEvent:
    title: str
    text: str
    buttons: list
    answer: str
    t: float = field(default_factory=time.time)


Policy = Callable[[str, str, list], "str | tuple[str, str] | None"]


def default_policy(title: str, text: str, buttons: list):
    """Yes if available, otherwise OK. A VBA run-time error dialog is
    answered with End (Continue is disabled there and would hang)."""
    if "&End" in buttons and "&Debug" in buttons:
        return "&End"
    for want in ("&Yes", "Yes", "&Evet", "Evet", "OK", "Tamam"):
        if want in buttons:
            return want
    return buttons[0] if buttons else None


def _window_text(hwnd) -> str:
    """Full window text (win32gui.GetWindowText truncates long MsgBox texts)."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        n = user32.SendMessageW(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.SendMessageW(hwnd, win32con.WM_GETTEXT, n + 1, buf)
        return buf.value
    except Exception:
        try:
            return win32gui.GetWindowText(hwnd)
        except Exception:
            return ""


def _children(hwnd):
    out = []

    def cb(h, _):
        out.append(h)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception:
        pass
    return out


class DialogWatcher(threading.Thread):
    def __init__(self, pid: int, policy: Policy = default_policy, interval: float = 0.15):
        super().__init__(daemon=True)
        self.pid = pid
        self.policy = policy
        self.interval = interval
        self.events: list[DialogEvent] = []
        self._stop = threading.Event()
        self._handled: dict[int, float] = {}

    def stop(self):
        self._stop.set()

    def run(self):
        while not self._stop.is_set():
            try:
                self._scan()
            except Exception:
                pass
            time.sleep(self.interval)

    def _scan(self):
        found = []

        def cb(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32gui.GetClassName(hwnd) != "#32770":
                    return True
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == self.pid:
                    found.append(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(cb, None)
        for hwnd in found:
            last = self._handled.get(hwnd)
            if last and time.time() - last < 1.0:
                continue
            self._answer(hwnd)

    def _answer(self, hwnd):
        title = _window_text(hwnd)
        texts, buttons, edits = [], {}, []
        for ch in _children(hwnd):
            cls = win32gui.GetClassName(ch)
            txt = _window_text(ch)
            if cls == "Button" and txt:
                buttons[txt] = ch
            elif cls == "Static" and txt:
                texts.append(txt)
            elif cls == "Edit":
                edits.append(ch)
        text = "\n".join(texts)
        decision = self.policy(title, text, list(buttons))
        edit_value = None
        if isinstance(decision, tuple):
            edit_value, decision = decision
        if decision is None or decision not in buttons:
            decision = default_policy(title, text, list(buttons))
        if edit_value is not None and edits:
            win32gui.SendMessage(edits[0], win32con.WM_SETTEXT, 0, edit_value)
        self.events.append(DialogEvent(title, text, list(buttons), f"{decision}" + (f" [{edit_value}]" if edit_value is not None else "")))
        self._handled[hwnd] = time.time()
        if decision in buttons:
            win32gui.PostMessage(buttons[decision], win32con.BM_CLICK, 0, 0)

    # convenience -----------------------------------------------------------
    def texts(self) -> list[str]:
        return [e.text for e in self.events]

    def dump(self, start: int = 0) -> str:
        lines = []
        for e in self.events[start:]:
            body = e.text.replace("\r", "")
            lines.append(f"--- [{e.title}] -> {e.answer}\n{body}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
#  Excel session
# ---------------------------------------------------------------------------
class ExcelSession:
    """Private Excel instance + dialog watcher. Use as a context manager."""

    def __init__(self, visible: bool = False, policy: Policy = default_policy):
        self.visible = visible
        self.policy = policy
        self.xl = None
        self.pid = None
        self.watcher: Optional[DialogWatcher] = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self.xl = win32com.client.DispatchEx("Excel.Application")
        self.xl.Visible = self.visible
        self.xl.DisplayAlerts = False
        try:
            self.xl.AutomationSecurity = 1  # msoAutomationSecurityLow: macros enabled
        except Exception:
            pass
        _, self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)
        self.watcher = DialogWatcher(self.pid, self.policy)
        self.watcher.start()
        return self

    def set_policy(self, policy: Policy):
        self.watcher.policy = policy

    def run(self, wb, macro: str, *args, timeout: float = 600):
        """Application.Run with a hard timeout (kills Excel if it hangs)."""
        done = threading.Event()
        timer = threading.Timer(timeout, self._kill_if_running, args=(done,))
        timer.start()
        try:
            return self.xl.Run(f"'{wb.Name}'!{macro}", *args)
        finally:
            done.set()
            timer.cancel()

    def _kill_if_running(self, done: threading.Event):
        if not done.is_set():
            self.kill()

    def kill(self):
        if self.pid:
            try:
                h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, self.pid)
                win32api.TerminateProcess(h, 1)
                win32api.CloseHandle(h)
            except Exception:
                pass

    def __exit__(self, *exc):
        try:
            if self.xl is not None:
                for wb in list(self.xl.Workbooks):
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                self.xl.Quit()
        except Exception:
            pass
        time.sleep(0.5)
        if self.watcher:
            self.watcher.stop()
        self.xl = None
        # make sure the process is gone
        try:
            h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, self.pid)
            win32api.CloseHandle(h)
            time.sleep(1.0)
            self.kill()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        return False


# ---------------------------------------------------------------------------
#  Source handling
# ---------------------------------------------------------------------------
def check_ascii(src: Path = SRC) -> list[str]:
    """All VBA sources must be pure printable ASCII (+ CR/LF/TAB)."""
    problems = []
    for f in sorted(src.glob("*.*")):
        if f.suffix.lower() not in (".bas", ".cls"):
            continue
        data = f.read_bytes()
        for i, line in enumerate(data.split(b"\n"), 1):
            for b in line:
                if b in (9, 13):
                    continue
                if b < 32 or b > 126:
                    problems.append(f"{f.name}:{i}: non-ASCII byte 0x{b:02x}")
                    break
    return problems


def _crlf_copy(path: Path, folder: Path) -> Path:
    text = path.read_text(encoding="ascii")
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    out = folder / path.name
    out.write_bytes(text.encode("ascii"))
    return out


def _document_code(path: Path) -> str:
    """Code for a document module: everything from 'Option Explicit' on,
    keeping the explanatory header comments."""
    text = path.read_text(encoding="ascii").replace("\r\n", "\n")
    return text.replace("\n", "\r\n")


def _is_project_component(name: str) -> bool:
    return (name in LEGACY_MODULES or name in PROJECT_CLASSES
            or any(name.startswith(p) and "_" in name for p in PROJECT_PREFIXES))


def ensure_panel_sheet(wb):
    for ws in wb.Worksheets:
        if ws.Name == "Panel":
            return ws
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Panel"
    return ws


def inject_sources(wb, src: Path = SRC, keep: tuple = ()) -> list[str]:
    """Replace project modules with src/. Returns a log list."""
    log = []
    proj = wb.VBProject
    comps = proj.VBComponents
    for comp in list(comps):
        if comp.Type in (VBEXT_STD, VBEXT_CLASS) and _is_project_component(comp.Name):
            log.append(f"removed {comp.Name}")
            comps.Remove(comp)
        elif comp.Type == VBEXT_STD and comp.Name in LEGACY_TEST_MODULES and comp.Name not in keep:
            log.append(f"removed {comp.Name} (v32/v33 test harness, depends on removed v33 code)")
            comps.Remove(comp)

    tmp = Path(tempfile.mkdtemp(prefix="apb_src_"))
    try:
        files = [src / "Sozluk.cls"] + sorted(src.glob("M*.bas"))
        for f in files:
            comps.Import(str(_crlf_copy(f, tmp)))
            log.append(f"imported {f.name}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    ensure_panel_sheet(wb)
    panel_comp = document_component(wb, "Panel")
    _replace_code(panel_comp.CodeModule, _document_code(src / "Panel.cls"))
    log.append(f"Panel sheet code -> {panel_comp.Name}")
    wb_comp = document_component(wb, None)
    _replace_code(wb_comp.CodeModule, _document_code(src / "ThisWorkbook.cls"))
    log.append(f"ThisWorkbook code -> {wb_comp.Name}")
    return log


def document_component(wb, sheet_name):
    """Document module of a sheet (by its tab name) or of the workbook
    (sheet_name=None). Works even when CodeName is still empty, which
    happens for sheets added while the VBE has not been refreshed."""
    comps = wb.VBProject.VBComponents
    for comp in comps:
        if comp.Type != VBEXT_DOCUMENT:
            continue
        try:
            props = {p.Name: p for p in comp.Properties}
        except Exception:
            continue
        if sheet_name is None:
            if "FullName" in props and "Worksheets" in props:
                return comp
        elif "Name" in props and "UsedRange" in props:
            if str(props["Name"].Value) == sheet_name:
                return comp
    raise LookupError(f"document module not found: {sheet_name or 'ThisWorkbook'}")


def _replace_code(cm, code: str):
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code)


def compile_project(session: ExcelSession, wb) -> tuple[bool, str]:
    """Debug > Compile VBAProject. Returns (ok, dialog text)."""
    start = len(session.watcher.events)
    vbe = session.xl.VBE
    # make the target project active so the compile command applies to it
    vbe.ActiveVBProject = wb.VBProject
    ctl = vbe.CommandBars.FindControl(Type=1, Id=578)
    if ctl is None:
        return False, "Compile command not found"
    if not ctl.Enabled:
        # already compiled state (unlikely right after import) - touch a module
        pass
    ctl.Execute()
    time.sleep(1.0)
    msgs = session.watcher.dump(start)
    ok = (not ctl.Enabled) and "error" not in msgs.lower()
    try:
        vbe.MainWindow.Visible = False
    except Exception:
        pass
    return ok, msgs


def export_sources(wb, dest: Path = SRC) -> list[str]:
    """Write the workbook's project modules back to src/ (LF line endings)."""
    log = []
    tmp = Path(tempfile.mkdtemp(prefix="apb_exp_"))
    try:
        for comp in wb.VBProject.VBComponents:
            if comp.Type == VBEXT_STD and _is_project_component(comp.Name) and comp.Name not in LEGACY_MODULES:
                p = tmp / f"{comp.Name}.bas"
            elif comp.Type == VBEXT_CLASS and comp.Name in PROJECT_CLASSES:
                p = tmp / f"{comp.Name}.cls"
            else:
                continue
            comp.Export(str(p))
            text = p.read_bytes().decode("ascii", errors="replace").replace("\r\n", "\n")
            (dest / p.name).write_text(text, encoding="ascii", newline="\n")
            log.append(f"exported {p.name}")
        panel = None
        for ws in wb.Worksheets:
            if ws.Name == "Panel":
                panel = ws
        for target, comp_name in (("Panel.cls", panel.CodeName if panel else None), ("ThisWorkbook.cls", wb.CodeName)):
            if not comp_name:
                continue
            cm = wb.VBProject.VBComponents(comp_name).CodeModule
            code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            (dest / target).write_text(code.replace("\r\n", "\n") + "\n", encoding="ascii", newline="\n")
            log.append(f"exported {target}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return log
