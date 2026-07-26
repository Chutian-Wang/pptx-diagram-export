import tempfile
from argparse import Namespace
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

import pymupdf

from pptx_diagram_export.backends import _powerpoint_macos_export
from pptx_diagram_export.bounds import inspect_rendered_pdf, parse_margin, parse_pages
from pptx_diagram_export.cli import _output_targets, _prepare_output_directory, run


class ParsePagesTests(TestCase):
    def test_all(self):
        self.assertEqual(parse_pages("all", 4), [1, 2, 3, 4])

    def test_ranges_are_sorted_and_deduplicated(self):
        self.assertEqual(parse_pages("5,1,3-5", 5), [1, 3, 4, 5])

    def test_rejects_out_of_range(self):
        with self.assertRaises(ValueError):
            parse_pages("1,6", 5)

    def test_rejects_descending_range(self):
        with self.assertRaises(ValueError):
            parse_pages("5-3", 5)


class ParseMarginTests(TestCase):
    def test_points_are_default(self):
        self.assertEqual(parse_margin("6"), 6.0)

    def test_metric_unit(self):
        self.assertAlmostEqual(parse_margin("25.4mm"), 72.0)

    def test_inches(self):
        self.assertEqual(parse_margin("0.5in"), 36.0)

    def test_rejects_negative_margin(self):
        with self.assertRaises(ValueError):
            parse_margin("-1")


class RenderedBoundsTests(TestCase):
    def test_uniform_background_is_trimmed(self):
        document = pymupdf.open()
        page = document.new_page(width=200, height=100)
        page.draw_rect(
            pymupdf.Rect(40, 20, 160, 80),
            color=(0, 0, 0),
            fill=(0.2, 0.4, 0.8),
            width=2,
        )
        geometry = inspect_rendered_pdf(document)
        box = geometry.slide_boxes[0]
        self.assertAlmostEqual(box.left, 39, delta=0.6)
        self.assertAlmostEqual(box.top, 19, delta=0.6)
        self.assertAlmostEqual(box.right, 161, delta=0.6)
        self.assertAlmostEqual(box.bottom, 81, delta=0.6)
        document.close()


class OutputTargetsTests(TestCase):
    def test_each_slide_gets_its_own_pdf(self):
        self.assertEqual(
            _output_targets(Path("out"), [1, 3, 12], 12, "pdf"),
            [
                Path("out/slide-01.pdf"),
                Path("out/slide-03.pdf"),
                Path("out/slide-12.pdf"),
            ],
        )

    def test_output_directory_is_created(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp) / "new" / "pdf"
            targets = _output_targets(directory, [2], 10, "pdf")
            _prepare_output_directory(directory, targets, force=False)
            self.assertTrue(directory.is_dir())


class PdfInputTests(TestCase):
    def _args(self, input_path: Path, output: Path, backend: str = "auto"):
        return Namespace(
            input=input_path,
            pages="all",
            format="pdf",
            margin="0",
            output=output,
            dpi=300.0,
            backend=backend,
            force=False,
        )

    def test_pdf_input_bypasses_presentation_renderer(self):
        with tempfile.TemporaryDirectory() as temp:
            temp_path = Path(temp)
            input_path = temp_path / "diagrams.pdf"
            document = pymupdf.open()
            page = document.new_page(width=200, height=100)
            page.draw_rect(
                pymupdf.Rect(40, 20, 160, 80),
                color=(0, 0, 0),
                fill=(0.2, 0.4, 0.8),
            )
            document.save(input_path)
            document.close()

            created, backend = run(
                self._args(input_path, temp_path / "exported")
            )

            self.assertEqual(backend, "pdf")
            self.assertEqual(len(created), 1)
            self.assertTrue(created[0].is_file())

    def test_pdf_input_rejects_explicit_backend(self):
        with tempfile.TemporaryDirectory() as temp:
            temp_path = Path(temp)
            input_path = temp_path / "diagrams.pdf"
            document = pymupdf.open()
            document.new_page()
            document.save(input_path)
            document.close()

            with self.assertRaisesRegex(
                ValueError, "--backend cannot be used with a PDF input"
            ):
                run(
                    self._args(
                        input_path, temp_path / "exported", backend="powerpoint"
                    )
                )


class MacPowerPointBackendTests(TestCase):
    def test_passes_paths_as_osascript_arguments(self):
        with tempfile.TemporaryDirectory() as temp:
            temp_path = Path(temp)
            input_path = temp_path / "input deck.pptx"
            output_path = temp_path / "rendered deck.pdf"

            def fake_run(command, **kwargs):
                staged_path = Path(command[-1])
                self.assertTrue(staged_path.is_file())
                self.assertEqual(command[0:2], ["/usr/bin/osascript", "-e"])
                self.assertEqual(command[-2], str(input_path))
                self.assertEqual(staged_path.parent, temp_path / "staging")
                self.assertNotIn(str(input_path), command[2])
                self.assertIn("set openedByThisScript to true", command[2])
                self.assertIn("save active presentation", command[2])
                staged_path.write_bytes(b"%PDF")
                return Namespace(returncode=0, stdout="", stderr="")

            with (
                patch(
                    "pptx_diagram_export.backends.find_powerpoint_macos",
                    return_value=Path("/Applications/Microsoft PowerPoint.app"),
                ),
                patch(
                    "pptx_diagram_export.backends._powerpoint_macos_staging_dir",
                    return_value=temp_path / "staging",
                ),
                patch(
                    "pptx_diagram_export.backends.subprocess.run",
                    side_effect=fake_run,
                ),
            ):
                _powerpoint_macos_export(input_path, output_path)

            self.assertTrue(output_path.is_file())
