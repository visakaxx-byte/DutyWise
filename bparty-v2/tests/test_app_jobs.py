from __future__ import annotations

import asyncio
import tempfile
import unittest
from pathlib import Path

from fastapi.testclient import TestClient

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

    def test_task_elapsed_time_is_persisted_and_public(self) -> None:
        original_task_dir = app_module.TASK_DIR
        original_now_epoch = app_module.now_epoch
        try:
            with tempfile.TemporaryDirectory() as tmp:
                app_module.TASK_DIR = Path(tmp)
                app_module.now_epoch = lambda: 110.0
                app_module.update_task(
                    "task456",
                    status="running",
                    started_at_epoch=100.0,
                    manifest_path="/tmp/input.xlsx",
                    bill_path="/tmp/bill.pdf",
                    query_cache={"product": {}},
                    input={"manifest_path": "/tmp/input.xlsx"},
                )
                running_record = app_module.read_json(app_module.task_path("task456"))
                running_public = app_module.public_task_record(running_record)

                app_module.now_epoch = lambda: 132.5
                app_module.update_task("task456", status="succeeded", completed_at_epoch=132.5)
                completed_record = app_module.read_json(app_module.task_path("task456"))
                completed_public = app_module.public_task_record(completed_record)
        finally:
            app_module.TASK_DIR = original_task_dir
            app_module.now_epoch = original_now_epoch

        self.assertEqual(running_record["elapsed_seconds"], 10.0)
        self.assertEqual(running_public["elapsed_minutes"], 0.17)
        self.assertEqual(completed_record["elapsed_seconds"], 32.5)
        self.assertEqual(completed_public["elapsed_seconds"], 32.5)
        self.assertNotIn("manifest_path", completed_public)
        self.assertNotIn("bill_path", completed_public)
        self.assertNotIn("query_cache", completed_public)
        self.assertNotIn("input", completed_public)

    def test_public_task_record_exposes_display_task_no(self) -> None:
        payload = app_module.public_task_record(
            {
                "task_id": "uuid123",
                "display_task_no": "PO-2026/A-1",
                "status": "queued",
            }
        )

        self.assertEqual(payload["task_id"], "uuid123")
        self.assertEqual(payload["display_task_no"], "PO-2026/A-1")


class JobQueueTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_task_dir = app_module.TASK_DIR
        self.original_cancel_dir = app_module.CANCEL_DIR
        self.original_upload_dir = app_module.UPLOAD_DIR
        self.original_output_dir = app_module.OUTPUT_DIR
        self.original_account_path = app_module.ACCOUNT_PATH
        self.original_permission_path = app_module.PERMISSION_PATH
        self.original_retention_days = app_module.JOB_RETENTION_DAYS
        self.original_cancelled_task_ids = set(app_module.CANCELLED_TASK_IDS)
        self.original_auth_tokens = dict(app_module.AUTH_TOKENS)
        base = Path(self.temp_dir.name)
        app_module.TASK_DIR = base / "tasks"
        app_module.CANCEL_DIR = base / "cancelled_jobs"
        app_module.UPLOAD_DIR = base / "uploads"
        app_module.OUTPUT_DIR = base / "outputs"
        app_module.ACCOUNT_PATH = base / "accounts.json"
        app_module.PERMISSION_PATH = base / "permissions.json"
        app_module.TASK_DIR.mkdir(parents=True, exist_ok=True)
        app_module.CANCEL_DIR.mkdir(parents=True, exist_ok=True)
        app_module.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        app_module.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        app_module.JOB_RETENTION_DAYS = 30
        app_module.CANCELLED_TASK_IDS.clear()
        app_module.AUTH_TOKENS.clear()

    def tearDown(self) -> None:
        app_module.TASK_DIR = self.original_task_dir
        app_module.CANCEL_DIR = self.original_cancel_dir
        app_module.UPLOAD_DIR = self.original_upload_dir
        app_module.OUTPUT_DIR = self.original_output_dir
        app_module.ACCOUNT_PATH = self.original_account_path
        app_module.PERMISSION_PATH = self.original_permission_path
        app_module.JOB_RETENTION_DAYS = self.original_retention_days
        app_module.CANCELLED_TASK_IDS.clear()
        app_module.CANCELLED_TASK_IDS.update(self.original_cancelled_task_ids)
        app_module.AUTH_TOKENS.clear()
        app_module.AUTH_TOKENS.update(self.original_auth_tokens)
        self.temp_dir.cleanup()

    def auth_headers(self, user_id: str = "user1") -> dict[str, str]:
        token = f"test-token-{user_id}"
        app_module.AUTH_TOKENS[token] = user_id
        return {"Authorization": f"Bearer {token}"}

    def test_accounts_file_seeds_default_users_and_login_still_works(self) -> None:
        self.assertFalse(app_module.ACCOUNT_PATH.exists())

        with TestClient(app_module.app) as client:
            admin = client.post("/api/login", json={"username": "admin", "password": "admin123"})
            user = client.post("/api/login", json={"username": "user4", "password": "user123"})

        self.assertEqual(admin.status_code, 200)
        self.assertEqual(user.status_code, 200)
        self.assertTrue(app_module.ACCOUNT_PATH.exists())
        accounts = app_module.read_json(app_module.ACCOUNT_PATH)["accounts"]
        self.assertEqual(set(accounts), {"admin", "user1", "user2", "user3", "user4"})
        self.assertNotIn("password", accounts["admin"])
        self.assertTrue(str(accounts["admin"]["password_hash"]).startswith("pbkdf2_sha256$"))

    def test_admin_can_create_next_user_with_default_password(self) -> None:
        with TestClient(app_module.app) as client:
            created = client.post("/api/accounts", headers=self.auth_headers("admin"), json={})
            login = client.post("/api/login", json={"username": "user5", "password": "user123"})

        self.assertEqual(created.status_code, 200)
        self.assertEqual(created.json()["data"]["user_id"], "user5")
        self.assertEqual(created.json()["data"]["can_view_all_tasks"], False)
        self.assertEqual(login.status_code, 200)

    def test_admin_can_create_next_user_with_custom_password(self) -> None:
        with TestClient(app_module.app) as client:
            created = client.post("/api/accounts", headers=self.auth_headers("admin"), json={"password": "secret456"})
            old_login = client.post("/api/login", json={"username": "user5", "password": "user123"})
            new_login = client.post("/api/login", json={"username": "user5", "password": "secret456"})

        self.assertEqual(created.status_code, 200)
        self.assertEqual(old_login.status_code, 401)
        self.assertEqual(new_login.status_code, 200)

    def test_password_change_invalidates_user_tokens_and_new_password_works(self) -> None:
        old_headers = self.auth_headers("user1")

        with TestClient(app_module.app) as client:
            before = client.get("/api/me", headers=old_headers)
            changed = client.patch(
                "/api/accounts/user1/password",
                headers=self.auth_headers("admin"),
                json={"new_password": "newpass123"},
            )
            after = client.get("/api/me", headers=old_headers)
            old_login = client.post("/api/login", json={"username": "user1", "password": "user123"})
            new_login = client.post("/api/login", json={"username": "user1", "password": "newpass123"})

        self.assertEqual(before.status_code, 200)
        self.assertEqual(changed.status_code, 200)
        self.assertEqual(after.status_code, 401)
        self.assertEqual(old_login.status_code, 401)
        self.assertEqual(new_login.status_code, 200)

    def test_admin_password_change_requires_current_password(self) -> None:
        with TestClient(app_module.app) as client:
            bad = client.patch(
                "/api/accounts/admin/password",
                headers=self.auth_headers("admin"),
                json={"new_password": "admin456", "current_admin_password": "wrong"},
            )

        self.assertEqual(bad.status_code, 403)

    def test_admin_password_change_invalidates_admin_token(self) -> None:
        old_headers = self.auth_headers("admin")

        with TestClient(app_module.app) as client:
            changed = client.patch(
                "/api/accounts/admin/password",
                headers=old_headers,
                json={"new_password": "admin456", "current_admin_password": "admin123"},
            )
            after = client.get("/api/me", headers=old_headers)
            old_login = client.post("/api/login", json={"username": "admin", "password": "admin123"})
            new_login = client.post("/api/login", json={"username": "admin", "password": "admin456"})

        self.assertEqual(changed.status_code, 200)
        self.assertTrue(changed.json()["data"]["current_user_logged_out"])
        self.assertEqual(after.status_code, 401)
        self.assertEqual(old_login.status_code, 401)
        self.assertEqual(new_login.status_code, 200)

    def test_new_user_appears_in_account_and_permission_lists(self) -> None:
        app_module.update_task("first", status="queued", owner_user_id="user1", created_at="2026-05-18T09:00:00")

        with TestClient(app_module.app) as client:
            client.post("/api/accounts", headers=self.auth_headers("admin"), json={})
            accounts = client.get("/api/accounts", headers=self.auth_headers("admin"))
            permissions = client.get("/api/permissions", headers=self.auth_headers("admin"))
            restricted = client.get("/api/jobs", headers=self.auth_headers("user5"))
            grant = client.patch(
                "/api/permissions/user5",
                headers=self.auth_headers("admin"),
                json={"can_view_all_tasks": True},
            )
            unrestricted = client.get("/api/jobs", headers=self.auth_headers("user5"))

        self.assertEqual(accounts.status_code, 200)
        self.assertIn("user5", [user["user_id"] for user in accounts.json()["data"]["users"]])
        self.assertIn("user5", [user["user_id"] for user in permissions.json()["data"]["users"]])
        self.assertEqual(restricted.json()["data"]["jobs"], [])
        self.assertEqual(grant.status_code, 200)
        self.assertEqual([job["task_id"] for job in unrestricted.json()["data"]["jobs"]], ["first"])

    def test_job_queue_payload_only_includes_unfinished_tasks(self) -> None:
        original_max_concurrency = app_module.MAX_CONCURRENT_JOBS
        app_module.MAX_CONCURRENT_JOBS = 5
        self.addCleanup(setattr, app_module, "MAX_CONCURRENT_JOBS", original_max_concurrency)
        app_module.update_task(
            "done1",
            status="succeeded",
            phase="done",
            progress=100,
            created_at="2026-05-17T09:00:00",
        )
        app_module.update_task(
            "queued1",
            status="queued",
            phase="queued",
            progress=0,
            created_at="2026-05-17T09:01:00",
        )
        app_module.update_task(
            "running1",
            status="running",
            phase="parse",
            progress=5,
            created_at="2026-05-17T09:02:00",
            started_at_epoch=100.0,
        )

        payload = app_module.build_job_queue_payload()

        self.assertEqual(payload["max_concurrency"], 5)
        self.assertEqual(payload["running_count"], 1)
        self.assertEqual(payload["queued_count"], 1)
        self.assertEqual([job["task_id"] for job in payload["jobs"]], ["running1", "queued1"])
        self.assertEqual(payload["jobs"][0]["running_position"], 1)
        self.assertEqual(payload["jobs"][1]["queue_position"], 1)

    def test_delete_finished_task_record_removes_json_and_output_file(self) -> None:
        output_path = app_module.OUTPUT_DIR / "finished.xlsx"
        output_path.write_text("fake", encoding="utf-8")
        app_module.update_task(
            "finished",
            status="succeeded",
            phase="done",
            progress=100,
            created_at="2026-05-18T09:00:00",
            output_file=str(output_path),
        )

        result = app_module.delete_task_record("finished")

        self.assertEqual(result["status"], "deleted")
        self.assertFalse(app_module.task_path("finished").exists())
        self.assertFalse(output_path.exists())
        self.assertIn(str(output_path), result["removed_files"])

    def test_delete_failed_task_record_without_output_file(self) -> None:
        app_module.update_task(
            "failed_record",
            status="failed",
            phase="failed",
            progress=100,
            created_at="2026-05-18T09:00:00",
            error="boom",
        )

        result = app_module.delete_task_record("failed_record")

        self.assertEqual(result["status"], "deleted")
        self.assertFalse(app_module.task_path("failed_record").exists())

    def test_failed_task_public_payload_exposes_source_file_downloads(self) -> None:
        manifest_path = app_module.UPLOAD_DIR / "failed_manifest.xlsx"
        bill_path = app_module.UPLOAD_DIR / "failed_bill.pdf"
        manifest_path.write_bytes(b"manifest")
        bill_path.write_bytes(b"bill")
        app_module.update_task(
            "failed_files",
            status="failed",
            phase="failed",
            progress=100,
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
            manifest_path=str(manifest_path),
            bill_path=str(bill_path),
            input={"manifest_path": str(manifest_path), "bill_path": str(bill_path)},
        )

        payload = app_module.public_task_record(app_module.read_json(app_module.task_path("failed_files")))

        self.assertEqual(payload["source_files"]["manifest"]["label"], "清单")
        self.assertEqual(payload["source_files"]["bill"]["label"], "提单")
        self.assertEqual(payload["source_files"]["manifest"]["download_url"], "/api/task-files/failed_files/manifest")
        self.assertEqual(payload["source_files"]["bill"]["download_url"], "/api/task-files/failed_files/bill")
        self.assertNotIn("manifest_path", payload)
        self.assertNotIn("bill_path", payload)
        self.assertNotIn("input", payload)

    def test_failed_task_source_files_are_downloadable_to_owner(self) -> None:
        manifest_path = app_module.UPLOAD_DIR / "owned_manifest.xlsx"
        bill_path = app_module.UPLOAD_DIR / "owned_bill.pdf"
        manifest_path.write_bytes(b"manifest bytes")
        bill_path.write_bytes(b"bill bytes")
        app_module.update_task(
            "failed_download",
            status="failed",
            phase="failed",
            progress=100,
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
            manifest_path=str(manifest_path),
            bill_path=str(bill_path),
        )

        with TestClient(app_module.app) as client:
            manifest = client.get("/api/task-files/failed_download/manifest", headers=self.auth_headers("user1"))
            bill = client.get("/api/task-files/failed_download/bill", headers=self.auth_headers("user1"))
            other_user = client.get("/api/task-files/failed_download/manifest", headers=self.auth_headers("user2"))

        self.assertEqual(manifest.status_code, 200)
        self.assertEqual(manifest.content, b"manifest bytes")
        self.assertEqual(bill.status_code, 200)
        self.assertEqual(bill.content, b"bill bytes")
        self.assertEqual(other_user.status_code, 404)

    def test_task_source_file_download_is_only_for_failed_tasks(self) -> None:
        manifest_path = app_module.UPLOAD_DIR / "succeeded_manifest.xlsx"
        manifest_path.write_bytes(b"manifest")
        app_module.update_task(
            "succeeded_download",
            status="succeeded",
            phase="done",
            progress=100,
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
            manifest_path=str(manifest_path),
        )

        with TestClient(app_module.app) as client:
            response = client.get("/api/task-files/succeeded_download/manifest", headers=self.auth_headers("user1"))

        self.assertEqual(response.status_code, 404)
        payload = app_module.public_task_record(app_module.read_json(app_module.task_path("succeeded_download")))
        self.assertNotIn("source_files", payload)

    def test_delete_unfinished_record_is_rejected(self) -> None:
        app_module.update_task(
            "queued_record",
            status="queued",
            phase="queued",
            progress=0,
            created_at="2026-05-18T09:00:00",
        )

        with self.assertRaises(app_module.HTTPException) as context:
            app_module.delete_task_record("queued_record")

        self.assertEqual(context.exception.status_code, 409)
        self.assertTrue(app_module.task_path("queued_record").exists())

    def test_delete_task_keeps_internal_id_and_display_no_separate(self) -> None:
        app_module.update_task(
            "internal-uuid",
            status="succeeded",
            phase="done",
            progress=100,
            display_task_no="PO-2026/ABC-01",
            created_at="2026-05-18T09:00:00",
        )

        payload = app_module.public_task_record(app_module.read_json(app_module.task_path("internal-uuid")))

        self.assertEqual(payload["task_id"], "internal-uuid")
        self.assertEqual(payload["display_task_no"], "PO-2026/ABC-01")

    def test_create_job_requires_and_persists_display_task_no(self) -> None:
        async def fake_run_clearance_job(task_id: str, task_record: dict) -> None:
            return None

        original_run_clearance_job = app_module.run_clearance_job
        original_running_tasks = dict(app_module.RUNNING_TASKS)
        app_module.run_clearance_job = fake_run_clearance_job
        app_module.RUNNING_TASKS.clear()
        self.addCleanup(setattr, app_module, "run_clearance_job", original_run_clearance_job)
        self.addCleanup(app_module.RUNNING_TASKS.clear)
        self.addCleanup(app_module.RUNNING_TASKS.update, original_running_tasks)

        manifest = b"manifest"
        bill = b"%PDF-1.4"
        with TestClient(app_module.app) as client:
            response = client.post(
                "/api/jobs",
                headers=self.auth_headers("user1"),
                data={
                    "display_task_no": "PO-2026/05-A1",
                    "target_tax_amount": "700",
                    "target_item_count": "10",
                },
                files={
                    "manifest": ("manifest.xlsx", manifest, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    "bill": ("bill.pdf", bill, "application/pdf"),
                },
            )

        self.assertEqual(response.status_code, 200)
        payload = response.json()["data"]
        self.assertEqual(payload["display_task_no"], "PO-2026/05-A1")
        self.assertNotEqual(payload["task_id"], payload["display_task_no"])
        record = app_module.read_json(app_module.task_path(payload["task_id"]))
        self.assertEqual(record["display_task_no"], "PO-2026/05-A1")
        self.assertEqual(record["owner_user_id"], "user1")

    def test_process_download_url_is_accessible_to_owner(self) -> None:
        async def fake_build_clearance(
            manifest_path,
            bill_path,
            output_dir,
            requested_profile,
            *,
            target_tax_amount,
            target_item_count,
            query_cache=None,
            progress_callback=None,
        ):
            output_path = app_module.OUTPUT_DIR / "process-result.xlsx"
            output_path.write_bytes(b"xlsx")
            return {
                "stats": {"rows": 1},
                "flow": [],
                "manifest": {},
                "bill": {},
                "bill_categories": [],
                "output_rows": [],
                "filter_summary": {},
                "output_file": str(output_path),
            }

        original_build_clearance = app_module.build_clearance
        app_module.build_clearance = fake_build_clearance
        self.addCleanup(setattr, app_module, "build_clearance", original_build_clearance)

        with TestClient(app_module.app) as client:
            response = client.post(
                "/api/process",
                headers=self.auth_headers("user1"),
                data={"target_tax_amount": "700", "target_item_count": "10"},
                files={
                    "manifest": ("manifest.xlsx", b"manifest", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    "bill": ("bill.pdf", b"%PDF-1.4", "application/pdf"),
                },
            )
            payload = response.json()["data"]
            download = client.get(payload["download_url"], headers=self.auth_headers("user1"))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(payload["owner_user_id"], "user1")
        self.assertTrue(app_module.task_path(payload["task_id"]).exists())
        self.assertEqual(download.status_code, 200)

    def test_job_api_requires_login(self) -> None:
        app_module.update_task(
            "owned",
            status="queued",
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
        )

        with TestClient(app_module.app) as client:
            response = client.get("/api/jobs")

        self.assertEqual(response.status_code, 401)

    def test_sub_account_only_sees_own_tasks(self) -> None:
        app_module.update_task(
            "mine",
            status="queued",
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
        )
        app_module.update_task(
            "others",
            status="queued",
            owner_user_id="user2",
            created_at="2026-05-18T09:01:00",
        )

        with TestClient(app_module.app) as client:
            response = client.get("/api/jobs", headers=self.auth_headers("user1"))

        self.assertEqual(response.status_code, 200)
        payload = response.json()["data"]
        self.assertEqual([job["task_id"] for job in payload["jobs"]], ["mine"])

    def test_admin_sees_all_tasks(self) -> None:
        app_module.update_task(
            "first",
            status="queued",
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
        )
        app_module.update_task(
            "second",
            status="queued",
            owner_user_id="user2",
            created_at="2026-05-18T09:01:00",
        )

        with TestClient(app_module.app) as client:
            response = client.get("/api/jobs", headers=self.auth_headers("admin"))

        self.assertEqual(response.status_code, 200)
        payload = response.json()["data"]
        self.assertEqual([job["task_id"] for job in payload["jobs"]], ["first", "second"])

    def test_admin_can_grant_sub_account_view_all_tasks(self) -> None:
        app_module.update_task(
            "first",
            status="queued",
            owner_user_id="user1",
            created_at="2026-05-18T09:00:00",
        )
        app_module.update_task(
            "second",
            status="queued",
            owner_user_id="user2",
            created_at="2026-05-18T09:01:00",
        )

        with TestClient(app_module.app) as client:
            grant = client.patch(
                "/api/permissions/user2",
                headers=self.auth_headers("admin"),
                json={"can_view_all_tasks": True},
            )
            response = client.get("/api/jobs", headers=self.auth_headers("user2"))

        self.assertEqual(grant.status_code, 200)
        self.assertEqual(response.status_code, 200)
        payload = response.json()["data"]
        self.assertEqual([job["task_id"] for job in payload["jobs"]], ["first", "second"])

    def test_view_all_sub_account_cannot_delete_other_users_task(self) -> None:
        app_module.update_task(
            "others_finished",
            status="succeeded",
            owner_user_id="user2",
            created_at="2026-05-18T09:01:00",
        )

        with TestClient(app_module.app) as client:
            client.patch(
                "/api/permissions/user1",
                headers=self.auth_headers("admin"),
                json={"can_view_all_tasks": True},
            )
            visible = client.get("/api/jobs/others_finished", headers=self.auth_headers("user1"))
            delete = client.delete("/api/jobs/others_finished", headers=self.auth_headers("user1"))

        self.assertEqual(visible.status_code, 200)
        self.assertEqual(delete.status_code, 403)
        self.assertTrue(app_module.task_path("others_finished").exists())

    def test_retention_cleanup_removes_only_old_finished_records(self) -> None:
        current_epoch = 1_800_000_000.0
        old_epoch = current_epoch - 31 * 24 * 60 * 60
        fresh_epoch = current_epoch - 10 * 24 * 60 * 60
        old_output = app_module.OUTPUT_DIR / "old.xlsx"
        fresh_output = app_module.OUTPUT_DIR / "fresh.xlsx"
        old_output.write_text("old", encoding="utf-8")
        fresh_output.write_text("fresh", encoding="utf-8")
        app_module.update_task(
            "old_done",
            status="succeeded",
            completed_at_epoch=old_epoch,
            output_file=str(old_output),
        )
        app_module.update_task(
            "fresh_done",
            status="succeeded",
            completed_at_epoch=fresh_epoch,
            output_file=str(fresh_output),
        )
        app_module.update_task(
            "old_running",
            status="running",
            created_at="2020-01-01T00:00:00",
        )

        result = app_module.cleanup_expired_job_records(retention_days=30, current_epoch=current_epoch)

        self.assertEqual(result["deleted_records"], 1)
        self.assertFalse(app_module.task_path("old_done").exists())
        self.assertFalse(old_output.exists())
        self.assertTrue(app_module.task_path("fresh_done").exists())
        self.assertTrue(fresh_output.exists())
        self.assertTrue(app_module.task_path("old_running").exists())


class JobConcurrencyTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_task_dir = app_module.TASK_DIR
        self.original_cancel_dir = app_module.CANCEL_DIR
        self.original_output_dir = app_module.OUTPUT_DIR
        self.original_max_concurrency = app_module.MAX_CONCURRENT_JOBS
        self.original_job_semaphore = app_module.JOB_SEMAPHORE
        self.original_running_tasks = dict(app_module.RUNNING_TASKS)
        self.original_cancelled_task_ids = set(app_module.CANCELLED_TASK_IDS)
        self.original_build_clearance = app_module.build_clearance
        app_module.TASK_DIR = Path(self.temp_dir.name) / "tasks"
        app_module.CANCEL_DIR = Path(self.temp_dir.name) / "cancelled_jobs"
        app_module.OUTPUT_DIR = Path(self.temp_dir.name) / "outputs"
        app_module.TASK_DIR.mkdir(parents=True, exist_ok=True)
        app_module.CANCEL_DIR.mkdir(parents=True, exist_ok=True)
        app_module.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        app_module.RUNNING_TASKS.clear()
        app_module.CANCELLED_TASK_IDS.clear()
        app_module.MAX_CONCURRENT_JOBS = 5
        app_module.JOB_SEMAPHORE = asyncio.Semaphore(5)

    async def asyncTearDown(self) -> None:
        pending = [task for task in app_module.RUNNING_TASKS.values() if not task.done()]
        for task in pending:
            task.cancel()
        if pending:
            await asyncio.gather(*pending, return_exceptions=True)
        app_module.TASK_DIR = self.original_task_dir
        app_module.CANCEL_DIR = self.original_cancel_dir
        app_module.OUTPUT_DIR = self.original_output_dir
        app_module.MAX_CONCURRENT_JOBS = self.original_max_concurrency
        app_module.JOB_SEMAPHORE = self.original_job_semaphore
        app_module.RUNNING_TASKS.clear()
        app_module.RUNNING_TASKS.update(self.original_running_tasks)
        app_module.CANCELLED_TASK_IDS.clear()
        app_module.CANCELLED_TASK_IDS.update(self.original_cancelled_task_ids)
        app_module.build_clearance = self.original_build_clearance
        self.temp_dir.cleanup()

    def make_task_record(self, task_id: str) -> dict:
        return {
            "task_id": task_id,
            "status": "queued",
            "phase": "queued",
            "progress": 0,
            "message": "等待后台任务启动",
            "created_at": f"2026-05-17T09:00:{task_id[-2:]}",
            "updated_at": "2026-05-17T09:00:00",
            "input": {
                "manifest_path": f"/tmp/{task_id}.xlsx",
                "bill_path": f"/tmp/{task_id}.pdf",
                "requested_profile": "auto",
                "target_tax_amount": 700.0,
                "target_item_count": 10,
            },
            "query_cache": {"product": {}, "hs": {}, "bill": {}},
        }

    def write_queued_task(self, task_id: str) -> dict:
        record = self.make_task_record(task_id)
        app_module.write_json(app_module.task_path(task_id), record)
        return record

    def fake_result(self, name: str) -> dict:
        return {
            "stats": {"fake": name},
            "flow": [],
            "manifest": {},
            "bill": {},
            "bill_categories": [],
            "output_rows": [],
            "filter_summary": {},
            "output_file": str(app_module.OUTPUT_DIR / f"{name}.xlsx"),
        }

    async def test_background_jobs_are_limited_to_five_running_tasks(self) -> None:
        active = 0
        max_active = 0
        active_lock = asyncio.Lock()
        five_started = asyncio.Event()
        release_jobs = asyncio.Event()

        async def fake_build_clearance(
            manifest_path,
            bill_path,
            output_dir,
            requested_profile,
            *,
            target_tax_amount,
            target_item_count,
            query_cache=None,
            progress_callback=None,
        ):
            nonlocal active, max_active
            async with active_lock:
                active += 1
                max_active = max(max_active, active)
                if active == 5:
                    five_started.set()
            if progress_callback:
                await progress_callback({"stage": "fake", "status": "running", "progress": 42, "message": "fake"})
            await release_jobs.wait()
            async with active_lock:
                active -= 1
            return self.fake_result(Path(manifest_path).stem)

        app_module.build_clearance = fake_build_clearance
        tasks = []
        for index in range(6):
            task_id = f"task{index:02d}"
            record = self.write_queued_task(task_id)
            tasks.append(app_module.schedule_clearance_job(task_id, record))

        await asyncio.wait_for(five_started.wait(), timeout=2)
        statuses = [app_module.read_json(app_module.task_path(f"task{index:02d}"))["status"] for index in range(6)]
        self.assertEqual(statuses.count("running"), 5)
        self.assertEqual(statuses.count("queued"), 1)
        self.assertLessEqual(max_active, 5)

        release_jobs.set()
        await asyncio.gather(*tasks)
        await asyncio.sleep(0)

        self.assertLessEqual(max_active, 5)
        self.assertEqual(app_module.RUNNING_TASKS, {})
        final_statuses = [app_module.read_json(app_module.task_path(f"task{index:02d}"))["status"] for index in range(6)]
        self.assertEqual(final_statuses.count("succeeded"), 6)

    async def test_failed_job_releases_slot_for_next_queued_job(self) -> None:
        app_module.MAX_CONCURRENT_JOBS = 1
        app_module.JOB_SEMAPHORE = asyncio.Semaphore(1)
        calls: list[str] = []

        async def fake_build_clearance(
            manifest_path,
            bill_path,
            output_dir,
            requested_profile,
            *,
            target_tax_amount,
            target_item_count,
            query_cache=None,
            progress_callback=None,
        ):
            task_name = Path(manifest_path).stem
            calls.append(task_name)
            if task_name == "fail":
                raise RuntimeError("boom")
            return self.fake_result(task_name)

        app_module.build_clearance = fake_build_clearance
        fail_record = self.write_queued_task("fail")
        next_record = self.write_queued_task("next")
        await asyncio.gather(
            app_module.schedule_clearance_job("fail", fail_record),
            app_module.schedule_clearance_job("next", next_record),
        )
        await asyncio.sleep(0)

        self.assertEqual(calls, ["fail", "next"])
        self.assertEqual(app_module.read_json(app_module.task_path("fail"))["status"], "failed")
        self.assertEqual(app_module.read_json(app_module.task_path("next"))["status"], "succeeded")
        self.assertEqual(app_module.RUNNING_TASKS, {})

    async def test_cancel_queued_job_deletes_record(self) -> None:
        record = self.write_queued_task("queued_cancel")
        app_module.write_json(app_module.task_path("queued_cancel"), record)

        result = await app_module.cancel_job_record("queued_cancel")

        self.assertEqual(result["status"], "cancelled")
        self.assertFalse(app_module.task_path("queued_cancel").exists())
        self.assertTrue(app_module.cancel_marker_path("queued_cancel").exists())
        self.assertEqual(app_module.RUNNING_TASKS, {})

    async def test_cancel_waiting_background_job_never_starts_after_slot_releases(self) -> None:
        app_module.MAX_CONCURRENT_JOBS = 1
        app_module.JOB_SEMAPHORE = asyncio.Semaphore(1)
        calls: list[str] = []
        first_started = asyncio.Event()
        release_first = asyncio.Event()
        second_started = asyncio.Event()

        async def fake_build_clearance(
            manifest_path,
            bill_path,
            output_dir,
            requested_profile,
            *,
            target_tax_amount,
            target_item_count,
            query_cache=None,
            progress_callback=None,
        ):
            task_name = Path(manifest_path).stem
            calls.append(task_name)
            if task_name == "blocked":
                first_started.set()
                await release_first.wait()
            if task_name == "queued_later":
                second_started.set()
            return self.fake_result(task_name)

        app_module.build_clearance = fake_build_clearance
        blocked_record = self.write_queued_task("blocked")
        queued_record = self.write_queued_task("queued_later")
        blocked_task = app_module.schedule_clearance_job("blocked", blocked_record)
        queued_task = app_module.schedule_clearance_job("queued_later", queued_record)

        await asyncio.wait_for(first_started.wait(), timeout=2)
        await asyncio.sleep(0)
        self.assertEqual(app_module.read_json(app_module.task_path("queued_later"))["status"], "queued")

        result = await app_module.cancel_job_record("queued_later")
        release_first.set()
        await asyncio.gather(blocked_task, queued_task, return_exceptions=True)

        self.assertEqual(result["status"], "cancelled")
        self.assertEqual(calls, ["blocked"])
        self.assertFalse(second_started.is_set())
        self.assertFalse(app_module.task_path("queued_later").exists())
        self.assertTrue(app_module.cancel_marker_path("queued_later").exists())
        self.assertEqual(app_module.RUNNING_TASKS, {})

    async def test_cancel_running_job_cancels_task_and_deletes_record(self) -> None:
        started = asyncio.Event()
        release = asyncio.Event()

        async def fake_build_clearance(
            manifest_path,
            bill_path,
            output_dir,
            requested_profile,
            *,
            target_tax_amount,
            target_item_count,
            query_cache=None,
            progress_callback=None,
        ):
            started.set()
            await release.wait()
            return self.fake_result(Path(manifest_path).stem)

        app_module.build_clearance = fake_build_clearance
        record = self.write_queued_task("running_cancel")
        task = app_module.schedule_clearance_job("running_cancel", record)
        await asyncio.wait_for(started.wait(), timeout=2)

        result = await app_module.cancel_job_record("running_cancel")
        await asyncio.gather(task, return_exceptions=True)

        self.assertEqual(result["status"], "cancelled")
        self.assertFalse(app_module.task_path("running_cancel").exists())
        self.assertTrue(app_module.cancel_marker_path("running_cancel").exists())
        self.assertEqual(app_module.RUNNING_TASKS, {})
        self.assertIn("running_cancel", app_module.CANCELLED_TASK_IDS)


if __name__ == "__main__":
    unittest.main()
