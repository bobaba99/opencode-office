import { expect, test } from "bun:test"
import { OfficeError } from "../src/errors"
import { runWorker } from "../src/worker"

test("echoes a payload through the venv python", async () => {
  const out = await runWorker<{ echo: { n: number } }>("echo.py", { n: 42 })
  expect(out.echo.n).toBe(42)
}, 180_000)

test("propagates WorkerError as OfficeError with code and hint", async () => {
  try {
    await runWorker("echo.py", { boom: true })
    expect.unreachable()
  } catch (e) {
    expect(e).toBeInstanceOf(OfficeError)
    expect((e as OfficeError).code).toBe("BOOM")
    expect((e as OfficeError).hint).toBe("test hint")
  }
}, 30_000)

test("kills a hung worker and reports WORKER_TIMEOUT", async () => {
  try {
    await runWorker("echo.py", { sleep_ms: 5_000 }, { timeoutMs: 500 })
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("WORKER_TIMEOUT")
  }
}, 30_000)
