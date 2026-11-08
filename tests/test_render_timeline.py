import filecmp
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import render_timeline as timeline


def make_doc(path: Path, headers, rows, heading="EVENTS"):
    doc = Document()
    if heading:
        doc.add_heading(heading, 1)
    table = doc.add_table(rows=1, cols=len(headers))
    for i, value in enumerate(headers):
        table.rows[0].cells[i].text = value
    for values in rows:
        cells = table.add_row().cells
        for i, value in enumerate(values):
            cells[i].text = str(value)
    doc.save(path)


class TimelineTests(unittest.TestCase):
    # --- Existing baseline tests (preserved) ---

    def test_aliases_and_multiple_date_formats(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["Event", "When", "Finish", "Theme"], [
                ["First", "27/07/2026", "", "A"],
                ["Second", "1 August 2026", "5 August 2026", "B"],
            ])
            result = timeline.read_docx(path)
            self.assertEqual(2, len(result.events))
            self.assertEqual("2026-07-27", result.events[0]["start"])
            self.assertEqual("2026-08-05", result.events[1]["end"])
            self.assertFalse([x for x in result.issues if x.level == "error"])

    def test_invalid_rows_are_reported_and_skipped(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["Date", "Title"], [
                ["not-a-date", "Bad"],
                ["2026-07-27", "Good"],
                ["2026-07-28", ""],
            ])
            result = timeline.read_docx(path)
            self.assertEqual(["Good"], [x["title"] for x in result.events])
            self.assertEqual(2, len([x for x in result.issues if x.level == "error"]))

    def test_duplicate_ids_are_made_unique(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["ID", "Start", "Title"], [
                ["same", "2026-01-01", "One"],
                ["same", "2026-01-02", "Two"],
            ])
            result = timeline.read_docx(path)
            self.assertEqual(["same", "same-2"], [x["id"] for x in result.events])
            self.assertTrue(any("Duplicate id" in x.message for x in result.issues))

    def test_unsafe_link_is_removed(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["Date", "Title", "Link"], [
                ["2026-01-01", "One", "javascript:alert(1)"],
            ])
            result = timeline.read_docx(path)
            self.assertEqual("", result.events[0]["link"])
            self.assertTrue(any("link" in x.message for x in result.issues))

    # --- Section 2: Table classification & multi-table regression tests ---

    def test_consecutive_valid_events_tables(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()
            doc.add_heading("EVENTS", 1)

            t1 = doc.add_table(rows=2, cols=2)
            t1.rows[0].cells[0].text = "Start"
            t1.rows[0].cells[1].text = "Title"
            t1.rows[1].cells[0].text = "2026-01-01"
            t1.rows[1].cells[1].text = "Phase 1 Event"

            t2 = doc.add_table(rows=2, cols=2)
            t2.rows[0].cells[0].text = "Date"
            t2.rows[0].cells[1].text = "Label"
            t2.rows[1].cells[0].text = "2026-02-01"
            t2.rows[1].cells[1].text = "Phase 2 Event"

            doc.save(path)
            result = timeline.read_docx(path)
            self.assertEqual(2, len(result.events))
            self.assertEqual(["Phase 1 Event", "Phase 2 Event"], [x["title"] for x in result.events])
            self.assertFalse([x for x in result.issues if x.level == "error"])

    def test_sequential_unrelated_table_does_not_fail_conversion(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()
            doc.add_heading("EVENTS", 1)

            t1 = doc.add_table(rows=2, cols=2)
            t1.rows[0].cells[0].text = "Start"
            t1.rows[0].cells[1].text = "Title"
            t1.rows[1].cells[0].text = "2026-01-01"
            t1.rows[1].cells[1].text = "Event 1"

            # Directly adjacent table with unrelated schema
            t2 = doc.add_table(rows=2, cols=2)
            t2.rows[0].cells[0].text = "Sign-off"
            t2.rows[0].cells[1].text = "Notes"
            t2.rows[1].cells[0].text = "Reviewer A"
            t2.rows[1].cells[1].text = "Approved"

            doc.save(path)
            result = timeline.read_docx(path)
            self.assertEqual(1, len(result.events))
            self.assertEqual("Event 1", result.events[0]["title"])
            self.assertFalse([x for x in result.issues if x.level == "error"])

    def test_explanatory_paragraphs_preserve_settings_and_events(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()
            doc.add_heading("SETTINGS", 1)
            doc.add_paragraph("Configure the global timeline options below before proceeding.")

            st = doc.add_table(rows=2, cols=2)
            st.rows[0].cells[0].text = "Option"
            st.rows[0].cells[1].text = "Value"
            st.rows[1].cells[0].text = "Title"
            st.rows[1].cells[1].text = "Custom Overview"

            doc.add_heading("EVENTS", 1)
            doc.add_paragraph("The milestones below represent major delivery checkpoints.")

            et = doc.add_table(rows=2, cols=2)
            et.rows[0].cells[0].text = "When"
            et.rows[0].cells[1].text = "Event"
            et.rows[1].cells[0].text = "2026-03-01"
            et.rows[1].cells[1].text = "Checkpoint 1"

            doc.save(path)
            result = timeline.read_docx(path)
            self.assertEqual("Custom Overview", result.settings.get("title"))
            self.assertEqual(1, len(result.events))
            self.assertEqual("Checkpoint 1", result.events[0]["title"])
            self.assertFalse([x for x in result.issues if x.level == "error"])

    def test_settings_schema_recognition_aliases(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()

            # No preceding heading, relying entirely on Option | Setting Value schema
            st = doc.add_table(rows=2, cols=2)
            st.rows[0].cells[0].text = "Option"
            st.rows[0].cells[1].text = "Setting Value"
            st.rows[1].cells[0].text = "Title"
            st.rows[1].cells[1].text = "Schema Only Title"

            doc.add_heading("EVENTS", 1)
            et = doc.add_table(rows=2, cols=2)
            et.rows[0].cells[0].text = "Start"
            et.rows[0].cells[1].text = "Title"
            et.rows[1].cells[0].text = "2026-04-01"
            et.rows[1].cells[1].text = "Alpha"

            doc.save(path)
            result = timeline.read_docx(path)
            self.assertEqual("Schema Only Title", result.settings.get("title"))
            self.assertEqual(1, len(result.events))

    def test_contradictory_heading_prefers_unambiguous_schema(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()
            # Contradictory heading SETTINGS above an EVENTS table
            doc.add_heading("SETTINGS", 1)
            et = doc.add_table(rows=2, cols=2)
            et.rows[0].cells[0].text = "Start Date"
            et.rows[0].cells[1].text = "Event Title"
            et.rows[1].cells[0].text = "2026-05-01"
            et.rows[1].cells[1].text = "Beta"

            doc.save(path)
            result = timeline.read_docx(path)
            self.assertEqual(1, len(result.events))
            self.assertEqual("Beta", result.events[0]["title"])

    # --- Section 3: Chronological comparison tests ---

    def test_chronological_comparisons_and_semantics(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["Start", "End", "Title"], [
                # Same-day date-only (valid)
                ["2026-05-01", "2026-05-01", "Same Day Date"],
                # Mixed date-only end on same day as timed start (calendar-day semantics: valid)
                ["2026-05-02T10:00", "2026-05-02", "Timed Start Same Day End"],
                # Mixed date-only start and timed end on same day (valid)
                ["2026-05-03", "2026-05-03T18:00", "Date Start Timed End"],
                # Timed events on same day with chronological order (valid)
                ["2026-05-04 09:00", "2026-05-04 17:00", "Timed Range"],
                # Leap year boundary (valid)
                ["2024-02-29", "2024-03-01", "Leap Boundary"],
                # Equivalent instant across timezones (valid)
                ["2026-05-05T10:00:00+00:00", "2026-05-05T11:00:00+01:00", "TZ Instant Equal"],
                # Timed events on same day inverted (invalid: error)
                ["2026-05-06 17:00", "2026-05-06 09:00", "Inverted Timed"],
                # Date-only inverted (invalid: error)
                ["2026-05-08", "2026-05-07", "Inverted Date"],
                # Timezone aware inverted (invalid: error)
                ["2026-05-09T10:00:00+00:00", "2026-05-09T10:30:00+01:00", "Inverted TZ Instant"],
            ])
            result = timeline.read_docx(path)
            valid_titles = [x["title"] for x in result.events]
            self.assertIn("Same Day Date", valid_titles)
            self.assertIn("Timed Start Same Day End", valid_titles)
            self.assertIn("Date Start Timed End", valid_titles)
            self.assertIn("Timed Range", valid_titles)
            self.assertIn("Leap Boundary", valid_titles)
            self.assertIn("TZ Instant Equal", valid_titles)

            self.assertNotIn("Inverted Timed", valid_titles)
            self.assertNotIn("Inverted Date", valid_titles)
            self.assertNotIn("Inverted TZ Instant", valid_titles)

            errors = [x for x in result.issues if x.level == "error"]
            self.assertEqual(3, len(errors))
            self.assertTrue(all("End date is before start date" in e.message for e in errors))

    # --- Section 6: Auto-generated event IDs tests ---

    def test_auto_id_collision_does_not_warn_falsely(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["ID", "Start", "Title"], [
                ["event-2", "2026-01-01", "Explicit Two"],
                ["", "2026-01-02", "Blank One"],
                ["", "2026-01-03", "Blank Two"],
            ])
            result = timeline.read_docx(path)
            self.assertEqual(["event-2", "event-1", "event-3"], [x["id"] for x in result.events])
            # No duplicate warnings should be emitted because the user did not supply a duplicate ID
            duplicate_warnings = [x for x in result.issues if "Duplicate id" in x.message]
            self.assertEqual(0, len(duplicate_warnings))

    # --- Section 4: Ungrouped and mixed-group timeline tests ---

    def test_mixed_group_timeline_uses_fallback_group(self):
        res = timeline.ParseResult(
            events=[
                {"id": "1", "start": "2026-01-01", "title": "Grouped", "group": "Track A", "category": "", "color": "", "link": "", "description": ""},
                {"id": "2", "start": "2026-01-02", "title": "Ungrouped", "group": "", "category": "", "color": "", "link": "", "description": ""},
            ]
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        self.assertIn('"Track A"', html_out)
        self.assertIn('"General"', html_out)
        # Ensure event 2 received a group assignment
        self.assertIn('"id": "2"', html_out)
        self.assertIn('"group": 2', html_out)

    def test_fully_ungrouped_timeline_has_no_groups(self):
        res = timeline.ParseResult(
            events=[
                {"id": "1", "start": "2026-01-01", "title": "A", "group": "", "category": "", "color": "", "link": "", "description": ""},
                {"id": "2", "start": "2026-01-02", "title": "B", "group": "", "category": "", "color": "", "link": "", "description": ""},
            ]
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        self.assertIn("const allItems=[", html_out)
        self.assertIn("groupData=[]", html_out)

    # --- Section 7: Zoom settings tests ---

    def test_zoom_settings_application(self):
        res = timeline.ParseResult(
            settings={"zoom_min_days": "7", "zoom_max_days": "365"},
            events=[{"id": "1", "start": "2026-01-01", "title": "A", "group": "", "category": "", "color": "", "link": "", "description": ""}],
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        self.assertIn("zoomMin:604800000", html_out)
        self.assertIn("zoomMax:31536000000", html_out)

    def test_invalid_zoom_settings_fallback_to_defaults(self):
        res = timeline.ParseResult(
            settings={"zoom_min_days": "invalid", "zoom_max_days": "-10"},
            events=[{"id": "1", "start": "2026-01-01", "title": "A", "group": "", "category": "", "color": "", "link": "", "description": ""}],
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        self.assertIn("zoomMin:86400000", html_out)
        self.assertIn("zoomMax:315576000000", html_out)
        self.assertTrue(any("zoom min days" in x.message for x in res.issues))
        self.assertTrue(any("zoom max days" in x.message for x in res.issues))

    def test_zoom_min_exceeding_max_fallback(self):
        res = timeline.ParseResult(
            settings={"zoom_min_days": "100", "zoom_max_days": "50"},
            events=[{"id": "1", "start": "2026-01-01", "title": "A", "group": "", "category": "", "color": "", "link": "", "description": ""}],
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        self.assertIn("zoomMin:86400000", html_out)
        self.assertIn("zoomMax:315576000000", html_out)
        self.assertTrue(any("zoom_min_days cannot exceed zoom_max_days" in x.message for x in res.issues))

    def test_zoom_validation_during_read_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            doc = Document()
            doc.add_heading("SETTINGS", 1)
            st = doc.add_table(rows=3, cols=2)
            st.rows[0].cells[0].text = "Setting"
            st.rows[0].cells[1].text = "Value"
            st.rows[1].cells[0].text = "Zoom Min Days"
            st.rows[1].cells[1].text = "-5"
            st.rows[2].cells[0].text = "Zoom Max Days"
            st.rows[2].cells[1].text = "abc"

            doc.add_heading("EVENTS", 1)
            et = doc.add_table(rows=2, cols=2)
            et.rows[0].cells[0].text = "Start"
            et.rows[0].cells[1].text = "Title"
            et.rows[1].cells[0].text = "2026-01-01"
            et.rows[1].cells[1].text = "Milestone"

            doc.save(path)
            res = timeline.read_docx(path)
            warnings = [x.message for x in res.issues if x.level == "warning"]
            self.assertTrue(any("Invalid zoom min days" in w for w in warnings))
            self.assertTrue(any("Invalid zoom max days" in w for w in warnings))

    # --- Security, escaping, and validation tests ---

    def test_html_escaping_and_script_breakout_prevention(self):
        payload_title = 'Milestone </script><script>alert("xss")</script> & "Quotes"'
        payload_desc = "Line 1\nLine 2 & <tags>"
        res = timeline.ParseResult(
            events=[{
                "id": "e1", "start": "2026-01-01", "title": payload_title,
                "description": payload_desc, "category": "Cat <1>",
                "group": "Track <1>", "color": "#123456", "link": "https://example.com/test",
            }]
        )
        html_out = timeline.build_html(res, ROOT, use_cdn=True)
        # Verify closing HTML tags in JSON data are forward-slash escaped
        self.assertIn("<\\/strong>", html_out)
        self.assertIn("<\\/p>", html_out)
        # Inside the script tag data block, unescaped </ must not exist
        data_script = html_out.split("<script>\nconst allItems=")[1].split("</script></body></html>")[0]
        self.assertNotIn("</script>", data_script)
        # HTML entities properly converted in tooltips and titles
        self.assertIn("&amp; &quot;Quotes&quot;", html_out)
        self.assertIn("Line 1<br>Line 2 &amp; &lt;tags&gt;", html_out)

    def test_color_and_link_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, ["Start", "Title", "Colour", "URL"], [
                ["2026-01-01", "Good", "rgb(255, 0, 0)", "https://example.com"],
                ["2026-01-02", "Bad CSS", "red; background: url(x)", "http://test.org"],
                ["2026-01-03", "Bad Link", "#fff", "data:text/html,<script>"],
            ])
            result = timeline.read_docx(path)
            self.assertEqual("rgb(255, 0, 0)", result.events[0]["color"])
            self.assertEqual("https://example.com", result.events[0]["link"])

            # Disallowed CSS/link should be cleaned
            self.assertEqual("", result.events[1]["color"])
            self.assertEqual("", result.events[2]["link"])

            warnings = [x.message for x in result.issues if x.level == "warning"]
            self.assertTrue(any("Invalid colour" in w for w in warnings))
            self.assertTrue(any("Unsafe or invalid link" in w for w in warnings))

    # --- CLI execution and exit codes ---

    def test_cli_execution_clean_success(self):
        fixture = ROOT / "examples" / "timeline_example.docx"
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "out"
            res = subprocess.run(
                [sys.executable, str(ROOT / "render_timeline.py"), str(fixture), "-o", str(out_dir)],
                capture_output=True,
                text=True,
            )
            self.assertEqual(0, res.returncode)
            self.assertTrue((out_dir / "index.html").is_file())
            self.assertTrue((out_dir / "data.json").is_file())
            self.assertTrue((out_dir / "qa_report.md").is_file())

    def test_cli_strict_exit_code_on_warnings(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            out_dir = Path(tmp) / "out"
            # Row has invalid end date (emits warning)
            make_doc(path, ["Start", "End", "Title"], [["2026-01-01", "TBD", "Milestone"]])

            # Without --strict: returns 0
            res = subprocess.run(
                [sys.executable, str(ROOT / "render_timeline.py"), str(path), "-o", str(out_dir)],
                capture_output=True,
                text=True,
            )
            self.assertEqual(0, res.returncode)

            # With --strict: returns 1
            res_strict = subprocess.run(
                [sys.executable, str(ROOT / "render_timeline.py"), str(path), "-o", str(out_dir), "--strict"],
                capture_output=True,
                text=True,
            )
            self.assertEqual(1, res_strict.returncode)

    def test_cli_missing_or_invalid_file_exit_code_2(self):
        res = subprocess.run(
            [sys.executable, str(ROOT / "render_timeline.py"), "non_existent.docx"],
            capture_output=True,
            text=True,
        )
        self.assertEqual(2, res.returncode)

    # --- Section 10: Canonical reference parity verification ---

    def test_canonical_reference_parity_exact(self):
        fixture = ROOT / "examples" / "timeline_example.docx"
        ref_dir = ROOT / "output" / "timeline_example"

        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "out"
            ret = timeline.cli([str(fixture), "-o", str(out_dir)])
            self.assertEqual(0, ret)

            # Byte-for-byte exact equality checks
            for fname in ["data.json", "qa_report.md", "index.html"]:
                self.assertTrue(
                    filecmp.cmp(out_dir / fname, ref_dir / fname, shallow=False),
                    f"Canonical artifact mismatch in {fname}",
                )

            # Verify known reference SHA-256 hashes
            hashes = {
                "data.json": "bc04256095ad9f6c1b4a57115702171b588e38fcb800e2aa1b01059cbf159964",
                "index.html": "d55b7d64c41ed165eb0f61b3fdd4711ae531c119eafc8ed5a5844f127ba5b56d",
                "qa_report.md": "581bb323f6a2600d9ee9af55b1f5320bc35d7d92e47a6e3d505c193f3714a158",
            }
            for fname, expected_hash in hashes.items():
                content = (out_dir / fname).read_bytes()
                actual_hash = hashlib.sha256(content).hexdigest()
                self.assertEqual(
                    expected_hash,
                    actual_hash,
                    f"SHA-256 hash mismatch for {fname}: expected {expected_hash}, got {actual_hash}",
                )

    # --- Section 11: Commercial Studio compatibility ---

    def test_commercial_studio_adapter_compatibility(self):
        studio_path = Path("/Users/jonathanwong/Desktop/startup/cloudpedagogy/projects/cloudpedagogy-word-conversion-studio")
        if not studio_path.exists():
            self.skipTest("Commercial Studio repository not found")

        fixture = ROOT / "examples" / "timeline_example.docx"
        saved_sys_path = list(sys.path)
        try:
            sys.path.insert(0, str(studio_path))
            from src.modules.timeline.adapter import TimelineAdapter
            adapter = TimelineAdapter()

            val_res = adapter.validate(fixture)
            self.assertTrue(val_res.is_valid, f"Validation failed: {val_res.errors}")

            with tempfile.TemporaryDirectory() as tmp:
                out_dir = Path(tmp) / "studio_out"
                conv_res = adapter.convert(fixture, out_dir, {})
                self.assertTrue(conv_res.is_valid, f"Conversion failed: {conv_res.errors}")
                self.assertTrue((out_dir / "index.html").is_file())
                self.assertTrue((out_dir / "data.json").is_file())
                self.assertTrue((out_dir / "qa_report.md").is_file())
                self.assertEqual(str(out_dir / "index.html"), conv_res.entry_point)
        finally:
            sys.path = saved_sys_path


if __name__ == "__main__":
    unittest.main()
