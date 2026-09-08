"""Unit tests for the generic background task runner (cdf_helper.tasks).

Hermetic: no Flask, no network. Exercises the submit/status lifecycle directly.
"""
import sys
import time

import pytest

from cdf_helper.tasks import submit, status


def test_submit_runs_work_and_reaches_done():
    seen = []
    tid = submit(lambda on_status: seen.append("ran") or {"ok": 1})
    for _ in range(50):
        s = status(tid)
        if s["status"] != "running":
            break
        time.sleep(0.01)
    assert s["status"] == "done", s
    assert s["result"] == {"ok": 1}, s
    assert seen == ["ran"], "work_fn must have executed"


def test_on_status_appends_to_live_log():
    def work(on_status):
        on_status("start")
        on_status("batch 1/2")
        return None

    tid = submit(work)
    for _ in range(50):
        s = status(tid)
        if s["status"] != "running":
            break
        time.sleep(0.01)
    assert s["status"] == "done"
    assert s["log"] == ["start", "batch 1/2"], s


def test_worker_exception_surfaces_as_error():
    def work(on_status):
        raise RuntimeError("boom")

    tid = submit(work)
    for _ in range(50):
        s = status(tid)
        if s["status"] != "running":
            break
        time.sleep(0.01)
    assert s["status"] == "error", s
    assert s["error"] == "boom", s


def test_unknown_task_is_missing():
    assert status("nope-not-a-real-id")["status"] == "missing"


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))