from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import app as app_module


class JobStoreTests(unittest.TestCase):
    def test_task_status_persists_status_flow(self) -> None:
        original_task_dir = app_module.TASK_DIR
        try:
            with tempfile.TemporaryDirectory() as tmp:
                app_module.TASK_DIR = Path(tmp)
                app_module.update_task("task123", status="queued", phase="queued", progress=0)
                app_module.update_task("task123", status="running", phase="crawler_manifest_products", progress=35)
                app_module.update_task("task123", status="succeeded", phase="done", progress=100)
                record = app_module.read_json(app_module.task_path("task123"))
        finally:
            app_module.TASK_DIR = original_task_dir

        self.assertEqual(record["status"], "succeeded")
        self.assertEqual(record["phase"], "done")
        self.assertEqual(record["progress"], 100)


if __name__ == "__main__":
    unittest.main()
