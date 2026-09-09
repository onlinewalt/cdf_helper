"""Background task runner for long-running work (e.g. AI enrichment + generation).

A deep, generic module: it knows nothing about Flask, parsing, or the 报关清单.
Callers submit a plain callable; the module runs it on a daemon thread, keeps a
live progress log (via an on_status callback), and exposes the task's state.

Public interface:
    submit(work_fn, *, on_status=None) -> str   # start a task, return its id
    status(task_id) -> dict                      # current {status, log, error, result}

Everything else is private. Tasks prune themselves (on finish and on access) so the
registry does not grow without bound.
"""

import threading
import time
import uuid
from typing import Callable, Optional

_MAX_AGE = 900  # seconds a finished task is retained before pruning
_LOG_MAX = 200  # cap on log entries kept per task


class _Task:
    def __init__(self, work_fn: Callable, on_status: Optional[Callable]):
        self._work = work_fn
        self._on_status = on_status
        self.status = "running"
        self.log = []
        self.error = None
        self.result = None
        self.ts = time.time()

    def _report(self, msg: str) -> None:
        self.log.append(msg)
        if len(self.log) > _LOG_MAX:
            self.log = self.log[-_LOG_MAX:]
        if self._on_status:
            self._on_status(msg)

    def run(self) -> None:
        try:
            self.result = self._work(self._report)
            self.status = "done"
        except Exception as e:  # surface failures to the caller via status()
            self.status = "error"
            self.error = str(e)
        finally:
            self.ts = time.time()


_registry = {}
_lock = threading.Lock()


def _new_id() -> str:
    return uuid.uuid4().hex[:16]


def _prune(now=None) -> None:
    now = now if now is not None else time.time()
    with _lock:
        stale = [tid for tid, t in _registry.items() if now - t.ts > _MAX_AGE]
        for tid in stale:
            _registry.pop(tid, None)


def submit(work_fn: Callable, *, on_status: Optional[Callable] = None) -> str:
    """Start work_fn on a daemon thread and return its task id.

    work_fn receives one argument (an on_status callback that appends a line to the
    task's live log) and returns the task's result payload. Any exception raised is
    captured and surfaced via status()["error"].
    """
    task = _Task(work_fn, on_status)
    tid = _new_id()
    with _lock:
        _registry[tid] = task
    threading.Thread(target=task.run, daemon=True).start()
    return tid


def status(task_id: str) -> dict:
    """Return the task's current state; prunes stale tasks as a side effect."""
    _prune()
    with _lock:
        task = _registry.get(task_id)
    if task is None:
        return {"status": "missing", "log": [], "error": None, "result": None}
    return {
        "status": task.status,
        "log": list(task.log),
        "error": task.error,
        "result": task.result,
    }