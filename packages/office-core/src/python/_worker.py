import json
import sys
import traceback


class WorkerError(Exception):
    def __init__(self, code, message, hint):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint


def run(main):
    try:
        payload = json.load(sys.stdin)
        print(json.dumps({"ok": True, "data": main(payload)}))
    except WorkerError as e:
        print(json.dumps({"ok": False, "error": {"code": e.code, "message": e.message, "hint": e.hint}}))
    except Exception:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": {
                        "code": "PYTHON_EXCEPTION",
                        "message": traceback.format_exc(limit=5),
                        "hint": "Unexpected worker failure; likely a bug in opencode-office.",
                    },
                }
            )
        )
