from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

import pymupdf

from .backends import BackendError, render_pdf
from .bounds import Box, RenderedGeometry, inspect_rendered_pdf, parse_margin, parse_pages


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-diagram-export",
        description="Export tightly cropped PowerPoint or PDF pages to PDF or PNG.",
    )
    parser.add_argument("input", type=Path, help="input .pptx or .pdf file")
    parser.add_argument(
        "--pages",
        default="all",
        metavar="SELECTION",
        help="slides to export, e.g. 1,3-5 (default: all)",
    )
    parser.add_argument(
        "--format",
        choices=("pdf", "png"),
        default="pdf",
        help="output format (default: pdf)",
    )
    parser.add_argument(
        "--margin",
        default="0",
        metavar="SIZE",
        help="margin around the diagram; pt by default, or use mm/cm/in (default: 0)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        metavar="DIRECTORY",
        help="output directory, created if needed (default: <input>-export)",
    )
    parser.add_argument(
        "--dpi",
        type=float,
        default=300.0,
        help="PNG resolution (default: 300)",
    )
    parser.add_argument(
        "--backend",
        choices=("auto", "powerpoint", "libreoffice"),
        default="auto",
        help="presentation renderer (default: auto)",
    )
    parser.add_argument(
        "--force", action="store_true", help="overwrite existing output files"
    )
    return parser


def _crop_box(
    geometry: RenderedGeometry,
    page: pymupdf.Page,
    slide_number: int,
    margin: float,
) -> Box:
    box = geometry.slide_boxes[slide_number - 1].expanded(margin).clamped(page.rect)
    if box.width <= 0 or box.height <= 0:
        raise ValueError(f"slide {slide_number} has an empty crop area")
    return box


def _output_targets(
    directory: Path,
    pages: list[int],
    slide_count: int,
    output_format: str,
) -> list[Path]:
    digits = max(2, len(str(slide_count)))
    return [
        directory / f"slide-{slide_number:0{digits}d}.{output_format}"
        for slide_number in pages
    ]


def _prepare_output_directory(
    directory: Path, targets: list[Path], force: bool
) -> None:
    if directory.exists() and not directory.is_dir():
        raise ValueError(f"output path is not a directory: {directory}")
    existing = [target for target in targets if target.exists()]
    if existing and not force:
        raise ValueError(
            f"output already exists: {existing[0]} (use --force to replace it)"
        )
    directory.mkdir(parents=True, exist_ok=True)


def _export_pdf(
    rendered_pdf: Path,
    targets: list[Path],
    pages: list[int],
    geometry: RenderedGeometry,
    margin: float,
) -> list[Path]:
    source = pymupdf.open(rendered_pdf)
    try:
        if source.page_count != geometry.slide_count:
            raise ValueError(
                "renderer returned "
                f"{source.page_count} pages for a {geometry.slide_count}-slide deck"
            )
        for slide_number, target in zip(pages, targets):
            result = pymupdf.open()
            try:
                source_page = source[slide_number - 1]
                box = _crop_box(geometry, source_page, slide_number, margin)
                clip = pymupdf.Rect(box.left, box.top, box.right, box.bottom)
                target_page = result.new_page(width=clip.width, height=clip.height)
                target_page.show_pdf_page(
                    target_page.rect,
                    source,
                    slide_number - 1,
                    clip=clip,
                    keep_proportion=False,
                )
                result.save(target, garbage=4, deflate=True)
            finally:
                result.close()
    finally:
        source.close()
    return targets


def _export_png(
    rendered_pdf: Path,
    targets: list[Path],
    pages: list[int],
    geometry: RenderedGeometry,
    margin: float,
    dpi: float,
) -> list[Path]:
    if dpi <= 0:
        raise ValueError("DPI must be greater than zero")

    source = pymupdf.open(rendered_pdf)
    try:
        if source.page_count != geometry.slide_count:
            raise ValueError(
                "renderer returned "
                f"{source.page_count} pages for a {geometry.slide_count}-slide deck"
            )
        matrix = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
        for slide_number, target in zip(pages, targets):
            page = source[slide_number - 1]
            box = _crop_box(geometry, page, slide_number, margin)
            clip = pymupdf.Rect(box.left, box.top, box.right, box.bottom)
            pixmap = page.get_pixmap(
                matrix=matrix,
                clip=clip,
                colorspace=pymupdf.csRGB,
                alpha=False,
                annots=True,
            )
            pixmap.save(target)
    finally:
        source.close()
    return targets


def run(args: argparse.Namespace) -> tuple[list[Path], str]:
    input_path = args.input.expanduser().resolve()
    if not input_path.is_file():
        raise ValueError(f"input file does not exist: {input_path}")
    input_suffix = input_path.suffix.lower()
    if input_suffix not in {".pptx", ".pdf"}:
        raise ValueError("input file must have a .pptx or .pdf extension")

    margin = parse_margin(args.margin)

    with tempfile.TemporaryDirectory(prefix="pptx-diagram-export-") as temp:
        temp_dir = Path(temp)
        if input_suffix == ".pdf":
            if args.backend != "auto":
                raise ValueError("--backend cannot be used with a PDF input")
            rendered_pdf = input_path
            backend = "pdf"
        else:
            rendered_pdf = temp_dir / "rendered.pdf"
            backend = render_pdf(
                input_path, rendered_pdf, temp_dir / "lo-profile", args.backend
            )
        rendered = pymupdf.open(rendered_pdf)
        try:
            geometry = inspect_rendered_pdf(rendered)
        finally:
            rendered.close()
        pages = parse_pages(args.pages, geometry.slide_count)
        output_directory = (
            args.output.expanduser().resolve()
            if args.output
            else input_path.with_name(f"{input_path.stem}-export")
        )
        targets = _output_targets(
            output_directory, pages, geometry.slide_count, args.format
        )
        _prepare_output_directory(output_directory, targets, args.force)

        if args.format == "pdf":
            created = _export_pdf(
                rendered_pdf,
                targets,
                pages,
                geometry,
                margin,
            )
        else:
            created = _export_png(
                rendered_pdf,
                targets,
                pages,
                geometry,
                margin,
                args.dpi,
            )
    return created, backend


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        created, backend = run(args)
    except (BackendError, OSError, ValueError, pymupdf.FileDataError) as exc:
        parser.error(str(exc))
    print(f"Renderer: {backend}")
    for path in created:
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
