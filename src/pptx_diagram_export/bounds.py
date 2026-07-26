from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

import pymupdf
from PIL import Image, ImageChops


@dataclass(frozen=True)
class Box:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top

    def expanded(self, amount: float) -> "Box":
        return Box(
            self.left - amount,
            self.top - amount,
            self.right + amount,
            self.bottom + amount,
        )

    def clamped(self, bounds: pymupdf.Rect) -> "Box":
        return Box(
            max(bounds.x0, self.left),
            max(bounds.y0, self.top),
            min(bounds.x1, self.right),
            min(bounds.y1, self.bottom),
        )


@dataclass(frozen=True)
class RenderedGeometry:
    slide_boxes: tuple[Box, ...]

    @property
    def slide_count(self) -> int:
        return len(self.slide_boxes)


def _visual_content_box(page: pymupdf.Page, analysis_dpi: float) -> Box:
    scale = analysis_dpi / 72.0
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        colorspace=pymupdf.csRGB,
        alpha=False,
        annots=True,
    )
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    corners = (
        image.getpixel((0, 0)),
        image.getpixel((image.width - 1, 0)),
        image.getpixel((0, image.height - 1)),
        image.getpixel((image.width - 1, image.height - 1)),
    )
    background = Counter(corners).most_common(1)[0][0]
    flat_background = Image.new("RGB", image.size, background)
    difference = ImageChops.difference(image, flat_background).convert("L")
    # Ignore one-level color noise from renderer rounding, but retain faint
    # antialiasing and shadows as visible content.
    mask = difference.point(lambda value: 255 if value > 1 else 0)
    pixel_box = mask.getbbox()
    if pixel_box is None:
        raise ValueError("slide has no pixels distinct from its background")

    left, top, right, bottom = pixel_box
    return Box(
        page.rect.x0 + left / scale,
        page.rect.y0 + top / scale,
        min(page.rect.x1, page.rect.x0 + right / scale),
        min(page.rect.y1, page.rect.y0 + bottom / scale),
    )


def inspect_rendered_pdf(
    document: pymupdf.Document, analysis_dpi: float = 144.0
) -> RenderedGeometry:
    boxes: list[Box] = []
    for slide_number, page in enumerate(document, start=1):
        try:
            boxes.append(_visual_content_box(page, analysis_dpi))
        except ValueError as exc:
            raise ValueError(
                f"slide {slide_number} has no visible diagram content"
            ) from exc
    return RenderedGeometry(tuple(boxes))


_PAGE_PART = re.compile(r"^(?P<first>\d+)(?:-(?P<last>\d+))?$")


def parse_pages(spec: str, slide_count: int) -> list[int]:
    if spec.strip().lower() == "all":
        return list(range(1, slide_count + 1))

    selected: set[int] = set()
    for raw_part in spec.split(","):
        part = raw_part.strip()
        match = _PAGE_PART.fullmatch(part)
        if not match:
            raise ValueError(
                f"invalid page selection {part!r}; use e.g. 1,3-5 or all"
            )
        first = int(match.group("first"))
        last = int(match.group("last") or first)
        if first > last:
            raise ValueError(f"invalid descending page range {part!r}")
        if first < 1 or last > slide_count:
            raise ValueError(
                f"page selection {part!r} is outside this deck's 1-{slide_count} range"
            )
        selected.update(range(first, last + 1))
    if not selected:
        raise ValueError("page selection is empty")
    return sorted(selected)


_MARGIN = re.compile(
    r"^\s*(?P<number>(?:\d+(?:\.\d*)?|\.\d+))\s*(?P<unit>pt|px|mm|cm|in)?\s*$",
    re.IGNORECASE,
)


def parse_margin(value: str) -> float:
    match = _MARGIN.fullmatch(value)
    if not match:
        raise ValueError("margin must be non-negative, e.g. 0, 6pt, 2mm, or 0.1in")
    number = float(match.group("number"))
    unit = (match.group("unit") or "pt").lower()
    factors = {
        "pt": 1.0,
        "px": 72.0 / 96.0,
        "mm": 72.0 / 25.4,
        "cm": 72.0 / 2.54,
        "in": 72.0,
    }
    return number * factors[unit]
