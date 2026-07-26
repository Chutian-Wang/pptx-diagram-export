from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


class BackendError(RuntimeError):
    pass


def _powerpoint_export(input_path: Path, output_pdf: Path) -> None:
    if sys.platform != "win32":
        raise BackendError("Microsoft PowerPoint automation is only available on Windows")
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
        if sys.platform == "win32":
            try:
                _powerpoint_export(input_path, output_pdf)
                return "powerpoint"
            except BackendError as exc:
                errors.append(str(exc))
        elif requested_backend == "powerpoint":
            errors.append("Microsoft PowerPoint automation is only available on Windows")

    if requested_backend in {"auto", "libreoffice"}:
        try:
            _libreoffice_export(input_path, output_pdf, profile_dir)
            return "libreoffice"
        except BackendError as exc:
            errors.append(str(exc))

    raise BackendError("; ".join(errors) or "no rendering backend is available")
