"""Generic background task runner.

submit(work_fn) runs work_fn(on_status) in a daemon thread and returns a task id.
status(task_id) returns the live state: running / done / error / missing.
"""
import threading
import uuid
from typing import Any, Callable, Dict

_lock = threading.Lock()
_tasks: Dict[str, Dict[str, Any]] = {}


def submit(work_fn: Callable[[Callable[[str], None]], Any]) -> str:
    task_id = uuid.uuid4().hex[:16]
    record: Dict[str, Any] = {
        "status": "running",
        "log": [],
        "error": None,
        "result": None,
    }
    with _lock:
        _tasks[task_id] = record

    def on_status(msg: str) -> None:
        with _lock:
            record["log"].append(msg)

    def _run() -> None:
        try:
            result = work_fn(on_status)
        except Exception as e:
            with _lock:
                record["status"] = "error"
                record["error"] = str(e)
        else:
            with _lock:
                record["status"] = "done"
                record["result"] = result

    threading.Thread(target=_run, daemon=True).start()
    return task_id


def status(task_id: str) -> Dict[str, Any]:
    with _lock:
        record = _tasks.get(task_id)
        if record is None:
            return {"status": "missing", "log": [], "error": None, "result": None}
        return {
            "status": record["status"],
            "log": list(record["log"]),
            "error": record["error"],
            "result": record["result"],
        }
