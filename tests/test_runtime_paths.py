"""Tests for install-root / frozen path helpers used by Deployment."""
from __future__ import annotations

import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from src.config import (
    AppConfig,
    ensure_runtime_data,
    get_bundle_root,
    get_install_root,
    knowledge_base_dir,
)


class RuntimeDataTests(unittest.TestCase):
    def test_install_root_points_at_project(self) -> None:
        root = get_install_root()
        self.assertTrue((root / "src" / "config.py").exists())
        self.assertEqual(root, get_bundle_root())

    def test_ensure_runtime_data_seeds_from_bundle_when_missing(self) -> None:
        project_root = get_install_root()
        project_kb = project_root / "data" / "knowledge_base" / "controls_catalog.json"
        if not project_kb.exists():
            self.skipTest("project knowledge_base missing")

        with TemporaryDirectory() as temp_dir:
            dest_root = Path(temp_dir)
            import src.config as config_mod

            original_bundle = config_mod.get_bundle_root
            config_mod.get_bundle_root = lambda: project_root
            try:
                ensure_runtime_data(dest_root)
            finally:
                config_mod.get_bundle_root = original_bundle

            dest_catalog = dest_root / "data" / "knowledge_base" / "controls_catalog.json"
            self.assertTrue(dest_catalog.exists())
            raw = json.loads(dest_catalog.read_text(encoding="utf-8"))
            self.assertIn("controls", raw)

    def test_ensure_runtime_data_does_not_overwrite_existing_catalog(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            kb = root / "data" / "knowledge_base"
            kb.mkdir(parents=True)
            marker = {"_schema_version": "test", "controls": []}
            catalog = kb / "controls_catalog.json"
            catalog.write_text(json.dumps(marker), encoding="utf-8")
            ensure_runtime_data(root)
            loaded = json.loads(catalog.read_text(encoding="utf-8"))
            self.assertEqual(loaded.get("_schema_version"), "test")

    def test_app_config_default_uses_install_root(self) -> None:
        cfg = AppConfig.default()
        self.assertEqual(cfg.output_dir, get_install_root() / "data" / "output")
        self.assertEqual(knowledge_base_dir(), get_install_root() / "data" / "knowledge_base")


if __name__ == "__main__":
    unittest.main()
