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

    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = app.Presentations.Open(
            str(input_path), ReadOnly=True, Untitled=False, WithWindow=False
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
        if app is not None:
            app.Quit()


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
    staging_dir = _powerpoint_macos_staging_dir()
    staging_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        prefix="rendered-", suffix=".pdf", dir=staging_dir, delete=False
    ) as staging_file:
        staged_pdf = Path(staging_file.name)

    script = r'''
on run argv
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set openedByThisScript to false

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
        error errorMessage number errorNumber
    end try
end run
'''
    try:
        result = subprocess.run(
            ["/usr/bin/osascript", "-e", script, str(input_path), str(staged_pdf)],
            check=False,
            capture_output=True,
            text=True,
            timeout=620,
        )
    except subprocess.TimeoutExpired as exc:
        staged_pdf.unlink(missing_ok=True)
        raise BackendError("PowerPoint export timed out") from exc

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
