import time

from _worker import run, WorkerError


def main(payload):
    if payload.get("boom"):
        raise WorkerError("BOOM", "requested failure", "test hint")
    if payload.get("sleep_ms"):
        time.sleep(payload["sleep_ms"] / 1000)
    return {"echo": payload}


run(main)
