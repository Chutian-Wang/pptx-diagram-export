from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


class BackendError(RuntimeError):
    pass


def _powerpoint_windows_export(input_path: Path, output_pdf: Path) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise BackendError("pywin32 is not installed") from exc

    # PowerPoint is a single-instance COM server, so DispatchEx attaches to an
    # instance the user already has open. Only quit it if we launched it.
    was_running = _powerpoint_windows_is_running()
    # Open a private copy so a deck the user already has open in PowerPoint
    # is never reused, exported with unsaved edits, or closed by us.
    copy_dir = Path(tempfile.mkdtemp(prefix="pptx-diagram-export-"))
    source_copy = copy_dir / f"source{input_path.suffix}"
    app = None
    presentation = None
    try:
        shutil.copy2(input_path, source_copy)
        app = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = app.Presentations.Open(
            str(source_copy), ReadOnly=True, Untitled=False, WithWindow=False
        )
        # SaveAs is supported by more Office/pywin32 combinations than
        # ExportAsFixedFormat, whose optional ExternalExporter COM parameter
        # fails marshaling on some current PowerPoint installations.
        presentation.SaveAs(str(output_pdf), 32)  # ppSaveAsPDF
    except Exception as exc:
        raise BackendError(f"PowerPoint export failed: {exc}") from exc
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None and not was_running:
            try:
                # Leave PowerPoint alone if the user opened a presentation
                # in the instance we launched while the export was running.
                if app.Presentations.Count == 0:
                    app.Quit()
            except Exception:
                pass
        shutil.rmtree(copy_dir, ignore_errors=True)


def _powerpoint_windows_is_running() -> bool:
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
            check=False,
            capture_output=True,
            text=True,
            timeout=30,
        )
    except (OSError, subprocess.SubprocessError):
        # If we cannot tell, assume it was running so we never close a
        # PowerPoint session the user is working in.
        return True
    return "powerpnt.exe" in result.stdout.lower()


def find_powerpoint_macos() -> Path | None:
    if sys.platform != "darwin":
        return None
    candidates = [
        Path("/Applications/Microsoft PowerPoint.app"),
        Path.home() / "Applications/Microsoft PowerPoint.app",
    ]
    return next((path for path in candidates if path.is_dir()), None)


def _powerpoint_macos_staging_dir() -> Path:
    return (
        Path.home()
        / "Library/Containers/com.microsoft.Powerpoint/Data/Documents"
        / "pptx-diagram-export"
    )


def _powerpoint_macos_export(input_path: Path, output_pdf: Path) -> None:
    if find_powerpoint_macos() is None:
        raise BackendError("Microsoft PowerPoint was not found")

    # Sandboxed PowerPoint cannot save into Python's /var/folders temporary
    # directory and may report the opaque error -9074. Stage the PDF in
    # PowerPoint's own container, then move it into the renderer workspace.
    # The input is staged there too: opening a private copy means a deck the
    # user already has open is never reused, exported with unsaved edits, or
    # closed by the script below.
    staging_dir = _powerpoint_macos_staging_dir()
    staging_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix="rendered-", suffix=".pdf", dir=staging_dir, delete=False
    ) as staging_file:
        staged_pdf = Path(staging_file.name)
    with tempfile.NamedTemporaryFile(
        prefix="source-", suffix=input_path.suffix, dir=staging_dir, delete=False
    ) as staging_file:
        staged_input = Path(staging_file.name)
    try:
        shutil.copy2(input_path, staged_input)
    except OSError as exc:
        staged_input.unlink(missing_ok=True)
        staged_pdf.unlink(missing_ok=True)
        raise BackendError(f"could not stage presentation: {exc}") from exc

    script = r'''
on run argv
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set openedByThisScript to false
    set wasRunning to application "Microsoft PowerPoint" is running

    try
        with timeout of 600 seconds
            tell application "Microsoft PowerPoint" to open (POSIX file inputPath)
            set openedByThisScript to true
            tell application "Microsoft PowerPoint" to save active presentation in (POSIX file outputPath) as save as PDF
            tell application "Microsoft PowerPoint" to close active presentation saving no
            set openedByThisScript to false
        end timeout
    on error errorMessage number errorNumber
        if openedByThisScript then
            try
                tell application "Microsoft PowerPoint" to close active presentation saving no
            end try
        end if
        my quitIfLaunched(wasRunning)
        error errorMessage number errorNumber
    end try
    my quitIfLaunched(wasRunning)
end run

on quitIfLaunched(wasRunning)
    if wasRunning then return
    try
        tell application "Microsoft PowerPoint"
            -- Leave PowerPoint alone if the user opened a presentation
            -- while the export was running.
            if (count of presentations) is 0 then quit saving no
        end tell
    end try
end quitIfLaunched
'''
    try:
        result = subprocess.run(
            ["/usr/bin/osascript", "-e", script, str(staged_input), str(staged_pdf)],
            check=False,
            capture_output=True,
            text=True,
            timeout=620,
        )
    except subprocess.TimeoutExpired as exc:
        staged_pdf.unlink(missing_ok=True)
        raise BackendError("PowerPoint export timed out") from exc
    finally:
        staged_input.unlink(missing_ok=True)

    if (
        result.returncode != 0
        or not staged_pdf.is_file()
        or staged_pdf.stat().st_size == 0
    ):
        detail = (result.stderr or result.stdout).strip()
        staged_pdf.unlink(missing_ok=True)
        raise BackendError(
            f"PowerPoint export failed (exit {result.returncode}): {detail}"
        )
    shutil.move(staged_pdf, output_pdf)


def _powerpoint_export(input_path: Path, output_pdf: Path) -> None:
    if sys.platform == "win32":
        _powerpoint_windows_export(input_path, output_pdf)
    elif sys.platform == "darwin":
        _powerpoint_macos_export(input_path, output_pdf)
    else:
        raise BackendError(
            "Microsoft PowerPoint automation is only available on Windows and macOS"
        )


def _libreoffice_candidates() -> list[Path]:
    candidates: list[Path] = []
    for executable in ("soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            candidates.append(Path(found))

    if sys.platform == "darwin":
        candidates.append(
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        )
    elif sys.platform == "win32":
        for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
            root = os.environ.get(variable)
            if root:
                candidates.append(Path(root) / "LibreOffice/program/soffice.exe")
    return candidates


def find_libreoffice() -> Path | None:
    return next((path for path in _libreoffice_candidates() if path.is_file()), None)


def _libreoffice_export(
    input_path: Path, output_pdf: Path, profile_dir: Path
) -> None:
    executable = find_libreoffice()
    if executable is None:
        raise BackendError(
            "LibreOffice was not found; install it or add 'soffice' to PATH"
        )

    profile_dir.mkdir(parents=True, exist_ok=True)
    output_dir = output_pdf.parent
    expected = output_dir / f"{input_path.stem}.pdf"
    command = [
        str(executable),
        "--headless",
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        "--convert-to",
        "pdf:impress_pdf_Export",
        "--outdir",
        str(output_dir),
        str(input_path),
    ]
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=180,
    )
    if result.returncode != 0 or not expected.is_file():
        detail = (result.stderr or result.stdout).strip()
        raise BackendError(
            f"LibreOffice export failed (exit {result.returncode}): {detail}"
        )
    if expected != output_pdf:
        expected.replace(output_pdf)


def render_pdf(
    input_path: Path,
    output_pdf: Path,
    profile_dir: Path,
    requested_backend: str,
) -> str:
    errors: list[str] = []

    if requested_backend in {"auto", "powerpoint"}:
        if sys.platform in {"win32", "darwin"}:
            try:
                _powerpoint_export(input_path, output_pdf)
                return "powerpoint"
            except BackendError as exc:
                errors.append(str(exc))
        elif requested_backend == "powerpoint":
            errors.append(
                "Microsoft PowerPoint automation is only available on Windows and macOS"
            )

    if requested_backend in {"auto", "libreoffice"}:
        try:
            _libreoffice_export(input_path, output_pdf, profile_dir)
            return "libreoffice"
        except BackendError as exc:
            errors.append(str(exc))

    raise BackendError("; ".join(errors) or "no rendering backend is available")
