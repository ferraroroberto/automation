#!/usr/bin/env python3
"""Focused tests for foldersearcher_core.

Run from the repo root:
    & .\\.venv\\Scripts\\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
"""

import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import foldersearcher_core as core  # noqa: E402


class TestRootNormalization(unittest.TestCase):
    def test_backslashes_become_forward_slashes(self):
        self.assertEqual(core.normalize_root(r"E:\onedrive\Documentos"), "E:/onedrive/Documentos")

    def test_trailing_separator_stripped(self):
        self.assertEqual(core.normalize_root("E:/onedrive/Documentos/"), "E:/onedrive/Documentos")

    def test_blank_input_yields_empty(self):
        self.assertEqual(core.normalize_root("   "), "")
        self.assertEqual(core.normalize_root(""), "")

    def test_duplicates_removed_order_preserved(self):
        roots = core.normalize_roots([
            r"E:\b",
            "E:/a",
            "E:/b/",           # duplicate of the first, differently spelled
            r"e:\A",           # duplicate of the second, different case
            "  ",              # dropped
        ])
        self.assertEqual(roots, ["E:/b", "E:/a"])


class TestJoinPath(unittest.TestCase):
    def test_ordinary_parent(self):
        self.assertEqual(core.join_path("E:/docs", "clientes"), "E:/docs/clientes")

    def test_drive_root_does_not_double_the_separator(self):
        # A drive root keeps its trailing slash, so a naive join would give
        # "E://clientes" and split one folder into two result rows.
        self.assertEqual(core.join_path("E:/", "clientes"), "E:/clientes")

    def test_drive_root_child_matches_the_walk_key_spelling(self):
        self.assertEqual(
            core.join_path(core.normalize_root("E:\\"), "clientes"),
            core.normalize_root("E:\\clientes"),
        )


class TestConfigMigration(unittest.TestCase):
    def test_legacy_root_folder_migrates(self):
        config = core.FolderSearcherConfig.from_dict(
            {"root_folder": "E:/onedrive/Documentos"}, "index.txt"
        )
        self.assertEqual(config.root_paths, ["E:/onedrive/Documentos"])

    def test_defaults_applied_when_keys_absent(self):
        config = core.FolderSearcherConfig.from_dict({"root_folder": "E:/x"}, "index.txt")
        self.assertEqual(config.skip_depth, 4)
        self.assertTrue(config.prune_email_branches)

    def test_prune_defaults_true_but_explicit_false_is_kept(self):
        config = core.FolderSearcherConfig.from_dict(
            {"root_paths": ["E:/x"], "prune_email_branches": False}, "index.txt"
        )
        self.assertFalse(config.prune_email_branches)

    def test_root_paths_and_legacy_key_both_survive(self):
        config = core.FolderSearcherConfig.from_dict(
            {"root_paths": ["E:/new"], "root_folder": "E:/old"}, "index.txt"
        )
        self.assertEqual(config.root_paths, ["E:/new", "E:/old"])

    def test_invalid_skip_depth_falls_back_to_default(self):
        config = core.FolderSearcherConfig.from_dict(
            {"root_paths": ["E:/x"], "skip_depth": "nonsense"}, "index.txt"
        )
        self.assertEqual(config.skip_depth, 4)

    def test_negative_skip_depth_clamped_to_zero(self):
        config = core.FolderSearcherConfig.from_dict(
            {"root_paths": ["E:/x"], "skip_depth": -3}, "index.txt"
        )
        self.assertEqual(config.skip_depth, 0)

    def test_round_trip_through_disk_drops_legacy_key(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "foldersearcher.json")
            with open(path, "w", encoding="utf-8") as handle:
                json.dump({"root_folder": "E:/legacy", "structure_file": "old.txt"}, handle)

            loaded = core.load_config(path, "index.txt")
            self.assertEqual(loaded.root_paths, ["E:/legacy"])

            core.save_config(path, loaded)
            with open(path, "r", encoding="utf-8") as handle:
                written = json.load(handle)
            self.assertEqual(written["root_paths"], ["E:/legacy"])
            self.assertNotIn("root_folder", written)
            self.assertEqual(written["skip_depth"], 4)
            self.assertIs(written["prune_email_branches"], True)

    def test_missing_config_file_yields_empty_roots(self):
        with tempfile.TemporaryDirectory() as tmp:
            config = core.load_config(os.path.join(tmp, "absent.json"), "index.txt")
            self.assertEqual(config.root_paths, [])


class TestEmailFolderDetection(unittest.TestCase):
    def test_real_email_names_match(self):
        for name in ("info@acme.com", "jose.perez@empresa.es", "a@b.co"):
            self.assertTrue(core.is_email_folder_name(name), name)

    def test_names_seen_in_the_committed_index_do_not_match(self):
        # These are the only "@" folder names present in the shipped index.
        for name in ("@font-face", "@parcel", "@vscode", "cl@ve"):
            self.assertFalse(core.is_email_folder_name(name), name)


class TestMultiRootScanAndPersistence(unittest.TestCase):
    def _build_tree(self, base):
        """Two roots, one of which contains an email-owning item."""
        root_a = os.path.join(base, "rootA")
        root_b = os.path.join(base, "rootB")
        os.makedirs(os.path.join(root_a, "clientes", "acme guia genai", "info@acme.com"))
        os.makedirs(os.path.join(root_a, "clientes", "acme guia genai", "ventas@acme.com"))
        os.makedirs(os.path.join(root_a, "clientes", "acme guia genai", "muestras", "2024"))
        os.makedirs(os.path.join(root_a, "clientes", "acme guia genai", "temporal"))
        os.makedirs(os.path.join(root_a, "clientes", "acme guia genai", "inventario"))
        os.makedirs(os.path.join(root_b, "otros", "plain guia genai"))
        return core.normalize_root(root_a), core.normalize_root(root_b)

    def test_scan_covers_every_root(self):
        with tempfile.TemporaryDirectory() as tmp:
            root_a, root_b = self._build_tree(tmp)
            index = core.FolderIndex()
            index.scan([root_a, root_b])

            paths = index.all_paths()
            self.assertIn(f"{root_a}/clientes/acme guia genai", paths)
            self.assertIn(f"{root_b}/otros/plain guia genai", paths)
            self.assertEqual(index.roots, [root_a, root_b])

    def test_missing_root_is_skipped_not_fatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            root_a, _root_b = self._build_tree(tmp)
            index = core.FolderIndex()
            index.scan([root_a, core.normalize_root(os.path.join(tmp, "does-not-exist"))])
            self.assertIn(f"{root_a}/clientes", index.all_paths())

    def test_sectioned_round_trip_preserves_index(self):
        with tempfile.TemporaryDirectory() as tmp:
            root_a, root_b = self._build_tree(tmp)
            structure_file = os.path.join(tmp, "folder_structure.txt")

            original = core.FolderIndex()
            original.scan([root_a, root_b])
            original.save(structure_file)

            with open(structure_file, "r", encoding="utf-8") as handle:
                text = handle.read()
            self.assertIn(f"Root: {root_a}", text)
            self.assertIn(f"Root: {root_b}", text)

            reloaded = core.FolderIndex()
            reloaded.load(structure_file)
            self.assertEqual(reloaded.roots, original.roots)
            self.assertEqual(reloaded.all_paths(), original.all_paths())
            self.assertEqual(reloaded.owning_items(), original.owning_items())

    def test_legacy_headerless_index_loads_against_legacy_root(self):
        with tempfile.TemporaryDirectory() as tmp:
            structure_file = os.path.join(tmp, "legacy.txt")
            with open(structure_file, "w", encoding="utf-8") as handle:
                handle.write("Ana\n")
                handle.write("  Ana\\archive\n")
                handle.write("\n")
                handle.write("Ana\\archive\n")
                handle.write("  Ana\\archive\\bank\n")
                handle.write("\n")

            index = core.FolderIndex()
            index.load(structure_file, legacy_root="E:/onedrive/Documentos")

            self.assertEqual(index.roots, ["E:/onedrive/Documentos"])
            self.assertIn("E:/onedrive/Documentos/Ana/archive/bank", index.all_paths())

    def test_legacy_index_without_a_root_is_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            structure_file = os.path.join(tmp, "legacy.txt")
            with open(structure_file, "w", encoding="utf-8") as handle:
                handle.write("Ana\n  Ana\\archive\n")

            index = core.FolderIndex()
            self.assertEqual(index.load(structure_file, legacy_root=None), 0)

    def test_absent_index_file_loads_as_empty(self):
        with tempfile.TemporaryDirectory() as tmp:
            index = core.FolderIndex()
            self.assertEqual(index.load(os.path.join(tmp, "nope.txt")), 0)


class TestDisplayPath(unittest.TestCase):
    def test_skip_depth_trims_leading_components(self):
        self.assertEqual(
            core.display_path("E:/onedrive/Documentos/clientes/acme/muestras", 4),
            "acme/muestras",
        )

    def test_zero_skip_depth_keeps_everything(self):
        self.assertEqual(
            core.display_path("E:/onedrive/Documentos/clientes", 0),
            "E:/onedrive/Documentos/clientes",
        )

    def test_shallow_path_falls_back_to_last_component(self):
        self.assertEqual(core.display_path("E:/onedrive", 4), "onedrive")

    def test_display_does_not_alter_the_absolute_path(self):
        absolute = "E:/onedrive/Documentos/clientes/acme"
        core.display_path(absolute, 4)
        self.assertEqual(absolute, "E:/onedrive/Documentos/clientes/acme")


class TestSearchTermParsing(unittest.TestCase):
    def test_whitespace_and_semicolons_both_split(self):
        self.assertEqual(core.parse_search_terms("genai; guia  extra"), ["genai", "guia", "extra"])

    def test_terms_are_lowercased(self):
        self.assertEqual(core.parse_search_terms("GenAI GUIA"), ["genai", "guia"])

    def test_empty_input_yields_no_terms(self):
        self.assertEqual(core.parse_search_terms("   ;  "), [])


class TestSearchAndPruning(unittest.TestCase):
    def setUp(self):
        """An item owning two email folders plus three sibling branches."""
        self.item = "E:/docs/clientes/acme guia genai"
        self.index = core.FolderIndex(roots=["E:/docs"])
        self.index.entries = {
            "E:/docs": ["E:/docs/clientes"],
            "E:/docs/clientes": [self.item],
            self.item: [
                f"{self.item}/info@acme.com",
                f"{self.item}/ventas@acme.com",
                f"{self.item}/muestras",
                f"{self.item}/temporal",
                f"{self.item}/inventario",
            ],
            f"{self.item}/muestras": [f"{self.item}/muestras/2024"],
            f"{self.item}/temporal": [],
            f"{self.item}/inventario": [],
            f"{self.item}/info@acme.com": [],
            f"{self.item}/ventas@acme.com": [],
        }

    def test_owning_items_detected(self):
        self.assertEqual(self.index.owning_items(), {self.item})

    def test_pruning_collapses_variants_and_siblings_to_one_result(self):
        results = core.search(self.index, "genai guia", skip_depth=0, prune_email_branches=True)
        self.assertEqual([r.absolute_path for r in results], [self.item])

    def test_without_pruning_every_branch_is_returned(self):
        results = core.search(self.index, "genai guia", skip_depth=0, prune_email_branches=False)
        paths = [r.absolute_path for r in results]
        self.assertIn(f"{self.item}/muestras", paths)
        self.assertIn(f"{self.item}/temporal", paths)
        self.assertIn(f"{self.item}/inventario", paths)
        self.assertIn(f"{self.item}/info@acme.com", paths)
        self.assertIn(f"{self.item}/ventas@acme.com", paths)
        self.assertGreater(len(paths), 1)

    def test_deep_match_below_an_item_still_collapses(self):
        results = core.search(self.index, "2024", skip_depth=0, prune_email_branches=True)
        self.assertEqual([r.absolute_path for r in results], [self.item])

    def test_paths_without_an_owning_ancestor_pass_through(self):
        self.index.entries["E:/docs/otros"] = ["E:/docs/otros/plain genai guia"]
        self.index.entries["E:/docs/otros/plain genai guia"] = []
        results = core.search(self.index, "genai guia", skip_depth=0, prune_email_branches=True)
        paths = [r.absolute_path for r in results]
        self.assertIn(self.item, paths)
        self.assertIn("E:/docs/otros/plain genai guia", paths)
        self.assertEqual(len(paths), 2)

    def test_nested_items_collapse_to_the_shallowest_owner(self):
        inner = f"{self.item}/muestras/sub guia genai"
        self.index.entries[f"{self.item}/muestras"].append(inner)
        self.index.entries[inner] = [f"{inner}/contacto@sub.com"]
        self.index.entries[f"{inner}/contacto@sub.com"] = []

        self.assertEqual(self.index.owning_items(), {self.item, inner})
        results = core.search(self.index, "genai guia", skip_depth=0, prune_email_branches=True)
        self.assertEqual([r.absolute_path for r in results], [self.item])

    def test_search_is_case_insensitive_and_ands_terms(self):
        results = core.search(self.index, "GENAI ACME", skip_depth=0, prune_email_branches=True)
        self.assertEqual([r.absolute_path for r in results], [self.item])
        self.assertEqual(core.search(self.index, "genai nomatch", skip_depth=0), [])

    def test_empty_search_returns_nothing(self):
        self.assertEqual(core.search(self.index, "   ", skip_depth=0), [])

    def test_results_carry_absolute_and_display_paths(self):
        results = core.search(self.index, "genai guia", skip_depth=3, prune_email_branches=True)
        self.assertEqual(results[0].absolute_path, self.item)
        self.assertEqual(results[0].display_path, "acme guia genai")


if __name__ == "__main__":
    unittest.main()
