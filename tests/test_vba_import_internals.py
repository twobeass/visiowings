"""Tests for VBA-import header stripping and folder-annotation handling."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

from visiowings.vba_import import VisioVBAImporter


# --------------------------------------------------------------------------- #
# _strip_vba_header
# --------------------------------------------------------------------------- #
class TestStripVbaHeaderImport:
    def test_drops_attributes_when_keep_vb_name_false(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            'Attribute VB_Name = "Module1"\n'
            "Attribute VB_GlobalNameSpace = False\n"
            "Option Explicit\n"
            "Sub Foo()\n"
            "End Sub\n"
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Attribute VB_Name" not in cleaned
        assert "Attribute VB_GlobalNameSpace" not in cleaned
        assert "Option Explicit" in cleaned
        assert "Sub Foo()" in cleaned

    def test_preserves_only_vb_name_when_keep_vb_name_true(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            'Attribute VB_Name = "Mod1"\nAttribute VB_GlobalNameSpace = False\nSub Foo()\nEnd Sub\n'
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=True)
        assert 'Attribute VB_Name = "Mod1"' in cleaned
        assert "Attribute VB_GlobalNameSpace" not in cleaned

    def test_idempotent_on_already_clean_code(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = "Option Explicit\nSub Foo()\nEnd Sub\n"
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Sub Foo()" in cleaned
        assert "Option Explicit" in cleaned

    def test_handles_form_with_nested_begin_blocks(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            "VERSION 5.00\n"
            "Begin {GUID} UserForm1\n"
            '   Caption = "UserForm1"\n'
            "   Begin {GUID} Button1\n"
            '      Caption = "Click"\n'
            "   End\n"
            "End\n"
            'Attribute VB_Name = "UserForm1"\n'
            "Option Explicit\n"
            "Sub Click()\n"
            "End Sub\n"
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Begin {GUID}" not in cleaned
        assert "VERSION 5.00" not in cleaned
        assert "Attribute VB_Name" not in cleaned
        assert "Option Explicit" in cleaned
        assert "Sub Click()" in cleaned

    def test_empty_input_returns_empty(self):
        importer = VisioVBAImporter("dummy.vsdm")
        assert importer._strip_vba_header("", keep_vb_name=False).strip() == ""


# --------------------------------------------------------------------------- #
# _ensure_folder_annotation
# --------------------------------------------------------------------------- #
class TestEnsureFolderAnnotation:
    def _doc_info(self, folder_name: str = "drawing1_vsdx"):
        info = MagicMock()
        info.folder_name = folder_name
        return info

    def test_noop_when_rubberduck_disabled(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        path = Path("/ws/drawing1_vsdx/Folder/Module1.bas")
        content = "Sub Foo()\nEnd Sub\n"
        assert importer._ensure_folder_annotation(content, path, self._doc_info()) == content

    def test_injects_annotation_for_nested_path(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/Foo/Bar/Module1.bas")
        content = 'Attribute VB_Name = "Module1"\nOption Explicit\n'
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert '\'@Folder("Foo.Bar")' in new
        assert "Option Explicit" in new

    def test_replaces_stale_annotation(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/New/Module1.bas")
        content = 'Attribute VB_Name = "Module1"\n\'@Folder("Old.Path")\nOption Explicit\n'
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert '\'@Folder("New")' in new
        assert "Old.Path" not in new

    def test_no_annotation_at_document_root(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/Module1.bas")
        content = "Option Explicit\n"
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert "@Folder" not in new


# --------------------------------------------------------------------------- #
# UAT iter4 #6 — `import --rd` must mark a moved file as "different"
# --------------------------------------------------------------------------- #
class TestCompareUnderRubberduck:
    """`_compare_module_content` must apply the rubberduck folder
    transformation to the on-disk text before comparing, so that a
    `Helpers/BasicLogic.bas` whose body otherwise matches Visio is
    detected as "needs import" — the import then injects
    `'@Folder("Helpers")` into the live module (§F3).
    """

    def _doc_info(self, folder_name: str = "sample"):
        info = MagicMock()
        info.folder_name = folder_name
        return info

    def _component(self, code: str):
        comp = MagicMock(name="comp")
        comp.CodeModule.Lines.return_value = code
        comp.CodeModule.CountOfLines = len(code.splitlines())
        return comp

    def test_moved_file_under_rd_diffs_as_different(self, tmp_path):
        """`<seed>/sample/Helpers/BasicLogic.bas` whose body matches
        Visio's existing BasicLogic (without @Folder) must compare as
        DIFFERENT so the import path runs."""

        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        sample_dir = tmp_path / "sample" / "Helpers"
        sample_dir.mkdir(parents=True)
        bas = sample_dir / "BasicLogic.bas"
        bas.write_text(
            'Attribute VB_Name = "BasicLogic"\nOption Explicit\nSub Foo()\nEnd Sub\n',
            encoding="utf-8",
        )

        # Visio has the same body but NO @Folder annotation.
        comp = self._component("Option Explicit\nSub Foo()\nEnd Sub\n")

        diff_no_rd_doc, *_ = importer._compare_module_content(bas, comp)
        diff_rd, *_ = importer._compare_module_content(bas, comp, doc_info=self._doc_info())

        # Without doc_info we mirror the OLD behaviour (would skip).
        assert diff_no_rd_doc is False
        # With doc_info the path-derived annotation makes the file differ,
        # so the importer schedules a re-import.
        assert diff_rd is True

    def test_same_folder_annotation_compares_equal(self, tmp_path):
        """If the file already carries the right `'@Folder("Helpers")`
        and Visio's body matches identically, the file is correctly
        treated as up-to-date — no spurious re-imports."""

        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        helpers = tmp_path / "sample" / "Helpers"
        helpers.mkdir(parents=True)
        bas = helpers / "BasicLogic.bas"
        bas.write_text(
            'Attribute VB_Name = "BasicLogic"\n'
            '\'@Folder("Helpers")\n'
            "Option Explicit\nSub Foo()\nEnd Sub\n",
            encoding="utf-8",
        )

        # Visio mirrors the same annotated body.
        comp = self._component(
            '\'@Folder("Helpers")\nOption Explicit\nSub Foo()\nEnd Sub\n'
        )

        diff, *_ = importer._compare_module_content(bas, comp, doc_info=self._doc_info())
        assert diff is False

    def test_compare_without_rd_ignores_path(self, tmp_path):
        """When `use_rubberduck=False`, the path is irrelevant: the
        same body diffs as identical even if it lives in a subfolder."""

        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        nested = tmp_path / "sample" / "Helpers"
        nested.mkdir(parents=True)
        bas = nested / "BasicLogic.bas"
        bas.write_text(
            'Attribute VB_Name = "BasicLogic"\nOption Explicit\nSub Foo()\nEnd Sub\n',
            encoding="utf-8",
        )

        comp = self._component("Option Explicit\nSub Foo()\nEnd Sub\n")
        diff, *_ = importer._compare_module_content(bas, comp, doc_info=self._doc_info())
        assert diff is False


# --------------------------------------------------------------------------- #
# _find_document_for_file
# --------------------------------------------------------------------------- #
class TestFindDocumentForFile:
    def test_direct_parent_match_outside_rubberduck_mode(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        marker = MagicMock(name="doc")
        importer.document_map = {"drawing1": marker}
        path = Path("/ws/drawing1/Module1.bas")
        assert importer._find_document_for_file(path) is marker

    def test_rubberduck_walks_up_to_find_document(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        marker = MagicMock(name="doc")
        importer.document_map = {"drawing1": marker}
        path = Path("/ws/drawing1/Sub/Sub2/Module1.bas")
        assert importer._find_document_for_file(path) is marker

    def test_rubberduck_max_walk_depth_avoids_runaway_traversal(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        importer.document_map = {"drawing1": MagicMock()}
        # Build a deeply nested path with no match at any level.
        deep = Path("/" + "/".join(f"d{i}" for i in range(20)) + "/Module1.bas")
        # Should not raise, should not match.
        assert importer._find_document_for_file(deep) is None

    def test_non_rubberduck_falls_back_to_main_document(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        importer.document_map = {"drawing1": MagicMock()}
        importer.doc_manager = MagicMock()
        main_doc = MagicMock(name="main")
        importer.doc_manager.get_main_document.return_value = main_doc
        path = Path("/ws/unrelated/Module1.bas")
        assert importer._find_document_for_file(path) is main_doc

    def test_rubberduck_returns_none_when_no_match(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        importer.document_map = {"drawing1": MagicMock()}
        path = Path("/ws/elsewhere/Module1.bas")
        assert importer._find_document_for_file(path) is None


# --------------------------------------------------------------------------- #
# UAT iter3 #2 — --force / --non-interactive plumbing
# --------------------------------------------------------------------------- #
class TestForceFlagPlumbing:
    """`--force` must imply `always_yes`; `--non-interactive` must set
    `non_interactive`. Both keep the batch conflict prompt from raising
    `EOFError` in CI / UAT runs."""

    def test_defaults_are_interactive(self):
        importer = VisioVBAImporter("dummy.vsdm")
        assert importer.always_yes is False
        assert importer.non_interactive is False

    def test_always_yes_propagated(self):
        importer = VisioVBAImporter("dummy.vsdm", always_yes=True)
        assert importer.always_yes is True

    def test_non_interactive_propagated(self):
        importer = VisioVBAImporter("dummy.vsdm", non_interactive=True)
        assert importer.non_interactive is True

    def test_ephemeral_propagated(self):
        importer = VisioVBAImporter("dummy.vsdm", ephemeral=True)
        assert importer.ephemeral is True

    def test_clear_dirty_flag_resets_saved_on_every_touched_doc(self):
        """`--ephemeral` flips `Document.Saved = True` on every doc visited."""
        importer = VisioVBAImporter("dummy.vsdm", ephemeral=True)

        # Build a fake doc_manager exposing 3 documents.
        d1 = MagicMock(name="main.vsdm")
        d1.doc = MagicMock(name="vsdm_doc")
        d2 = MagicMock(name="stencil.vssm")
        d2.doc = MagicMock(name="vssm_doc")
        d3 = MagicMock(name="template.vstm")
        d3.doc = MagicMock(name="vstm_doc")

        importer.doc_manager = MagicMock()
        importer.doc_manager.get_all_documents_with_vba.return_value = [d1, d2, d3]

        importer._clear_dirty_flag()

        assert d1.doc.Saved is True
        assert d2.doc.Saved is True
        assert d3.doc.Saved is True

    def test_create_temp_file_raises_on_emoji_in_cp1252(self, tmp_path):
        """Iter3 #4 (DATA LOSS): an emoji in a cp1252 doc must raise upfront,
        BEFORE we touch the existing module, so the user keeps their code."""
        from visiowings.exceptions import EncodingIncompatibilityError

        importer = VisioVBAImporter("dummy.vsdm")
        importer.codepage = "cp1252"
        bas = tmp_path / "BasicLogic.bas"
        bas.write_text("' Earth: 🌍\nSub Foo()\nEnd Sub\n", encoding="utf-8")

        try:
            importer._create_temp_codepage_file(bas, "cp1252")
        except EncodingIncompatibilityError as exc:
            assert exc.file == "BasicLogic.bas"
            assert exc.codepage == "cp1252"
            assert "🌍" in exc.sample_chars
            assert "cp65001" in exc.message  # the hint points at UTF-8
        else:
            raise AssertionError("expected EncodingIncompatibilityError")

    def test_create_temp_file_succeeds_for_pure_ascii_in_cp1252(self, tmp_path):
        importer = VisioVBAImporter("dummy.vsdm")
        importer.codepage = "cp1252"
        bas = tmp_path / "BasicLogic.bas"
        bas.write_text("Option Explicit\nSub Foo()\nEnd Sub\n", encoding="utf-8")

        temp_path = importer._create_temp_codepage_file(bas, "cp1252")
        try:
            assert Path(temp_path).exists()
            # Content round-trips cleanly through cp1252.
            assert "Option Explicit" in Path(temp_path).read_text(encoding="cp1252")
        finally:
            Path(temp_path).unlink(missing_ok=True)

    def test_create_temp_file_emoji_ok_in_cp65001(self, tmp_path):
        """UTF-8 (cp65001) representation must NOT raise — that's the recommended override."""
        importer = VisioVBAImporter("dummy.vsdm")
        bas = tmp_path / "BasicLogic.bas"
        bas.write_text("' Earth: 🌍\nSub Foo()\nEnd Sub\n", encoding="utf-8")

        temp_path = importer._create_temp_codepage_file(bas, "utf-8")
        try:
            assert "🌍" in Path(temp_path).read_text(encoding="utf-8")
        finally:
            Path(temp_path).unlink(missing_ok=True)

    def test_clear_dirty_flag_survives_per_doc_errors(self):
        """A single read-only stencil must not stop us from clearing the rest."""
        importer = VisioVBAImporter("dummy.vsdm", ephemeral=True)

        good = MagicMock(name="ok")
        good.doc = MagicMock(name="ok_doc")

        # Build a doc whose Saved setter raises (e.g. read-only).
        bad = MagicMock(name="readonly")
        bad_doc = MagicMock(name="ro_doc")
        type(bad_doc).Saved = property(
            lambda self: True,
            lambda self, v: (_ for _ in ()).throw(Exception("read-only")),
        )
        bad.doc = bad_doc

        importer.doc_manager = MagicMock()
        importer.doc_manager.get_all_documents_with_vba.return_value = [bad, good]

        importer._clear_dirty_flag()

        # The healthy doc still gets the flag.
        assert good.doc.Saved is True

    def test_non_interactive_propagates_to_subparser(self):
        """`--non-interactive` flag on `import` reaches argparse."""
        from visiowings.cli import _build_parser

        parser = _build_parser()
        ns = parser.parse_args(["import", "--file", "x.vsdm", "--non-interactive"])
        assert ns.non_interactive is True
        assert ns.force is False

    def test_cli_force_implies_always_yes_in_import(self, monkeypatch, tmp_path):
        """`cmd_import` must pass `always_yes=force` to the importer."""

        from visiowings import cli

        captured: dict = {}

        class _StubImporter:
            def __init__(self, *_a, **kwargs):
                captured.update(kwargs)

            def import_modules_from_dir(self, *_a, **_kw):
                return 0

        # `cmd_import` lazy-imports VisioVBAImporter from .vba_import,
        # so we patch the attribute on the already-loaded module via
        # `sys.modules` rather than `import ... as` (which would conflict
        # with the top-level `from visiowings.vba_import import …`).
        monkeypatch.setattr(sys.modules["visiowings.vba_import"], "VisioVBAImporter", _StubImporter)
        monkeypatch.setattr(cli, "_validate_visio_file", lambda p: p)
        monkeypatch.setattr(cli, "_validate_readable_dir", lambda p, label: p)

        import argparse

        args = argparse.Namespace(
            file=tmp_path / "dummy.vsdm",
            input=str(tmp_path),
            force=True,
            non_interactive=False,
            debug=False,
            codepage=None,
            rubberduck=False,
        )
        cli.cmd_import(args)

        assert captured.get("force_document") is True
        assert captured.get("always_yes") is True
        assert captured.get("non_interactive") is False


# --------------------------------------------------------------------------- #
# UAT iter3 #1 — Option Explicit dedupe after import
# --------------------------------------------------------------------------- #
class TestDedupeOptionExplicit:
    """Iter3 #1: Visio auto-prepends Option Explicit when the VBE option
    "Require Variable Declaration" is on. Our exported `.bas` already
    has it; without the post-import dedupe every round-trip would
    accumulate one duplicate.
    """

    def _make_component(self, lines: list[str]):
        """Build a MagicMock that mimics a VBA component with a CodeModule."""

        comp = MagicMock(name="comp")
        cm = comp.CodeModule
        # `Lines(start, count)` in the real API returns a newline-joined
        # block; for the dedupe we only ask for one line at a time.
        state = {"lines": list(lines)}

        def _lines(start, count):
            # 1-based indexing, count == 1 in our scanner.
            return state["lines"][start - 1]

        def _delete(start, count):
            del state["lines"][start - 1 : start - 1 + count]
            cm.CountOfLines = len(state["lines"])

        cm.Lines.side_effect = _lines
        cm.DeleteLines.side_effect = _delete
        cm.CountOfLines = len(state["lines"])
        # Expose the mutating state for assertions
        comp._state = state  # type: ignore[attr-defined]
        return comp

    def test_keeps_a_lone_option_explicit(self):
        comp = self._make_component(["Option Explicit", "", "Sub Foo()", "End Sub"])
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 0
        assert comp._state["lines"][0] == "Option Explicit"

    def test_removes_one_duplicate(self):
        comp = self._make_component(
            ["Option Explicit", "Option Explicit", "", "Sub Foo()", "End Sub"]
        )
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 1
        assert comp._state["lines"].count("Option Explicit") == 1

    def test_removes_multiple_duplicates(self):
        comp = self._make_component(
            [
                "Option Explicit",
                "",
                "Option Explicit",
                "",
                "Option Explicit",
                "Sub Foo()",
                "End Sub",
            ]
        )
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 2
        assert comp._state["lines"].count("Option Explicit") == 1

    def test_case_insensitive(self):
        comp = self._make_component(
            ["Option Explicit", "OPTION EXPLICIT", "  option explicit  ", "Sub Foo()"]
        )
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 2
        # First-occurrence semantics: keep the literal first line, drop the rest.
        assert comp._state["lines"][0] == "Option Explicit"

    def test_stops_at_procedure_boundary(self):
        """Anything inside Sub/Function is not considered a declaration."""

        comp = self._make_component(
            [
                "Option Explicit",
                "Sub Foo()",
                "    Option Explicit",  # inside a procedure — not real
                "End Sub",
            ]
        )
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 0  # the second occurrence was past the boundary

    def test_handles_empty_module_without_error(self):
        comp = self._make_component([])
        removed = VisioVBAImporter._dedupe_option_explicit(comp)
        assert removed == 0

    def test_swallows_exceptions_in_codemodule_access(self):
        comp = MagicMock(name="broken")
        comp.CodeModule.CountOfLines = 5
        comp.CodeModule.Lines.side_effect = Exception("broken COM")
        # Must not raise; returns 0 on failure.
        assert VisioVBAImporter._dedupe_option_explicit(comp) == 0
