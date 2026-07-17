import { expect, test } from "bun:test"

test("workspace resolves office-core", async () => {
  const pkg = await Bun.file(new URL("../package.json", import.meta.url)).json()
  expect(pkg.name).toBe("@opencode-office/core")
})
