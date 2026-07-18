import { createHash } from "node:crypto"
import { existsSync } from "node:fs"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "./errors"
import { defaultCacheDir } from "./runtime"

export async function backupFile(file: string, cacheDir = defaultCacheDir()): Promise<string> {
  let data: Buffer
  try {
    data = await readFile(file)
  } catch (e) {
    throw new OfficeError(
      "FILE_OPEN",
      `Cannot read ${file} for backup: ${e instanceof Error ? e.message : String(e)}`,
      "Check that the path exists and is readable before editing.",
    )
  }
  const hash = createHash("sha256").update(data).digest("hex").slice(0, 12)
  const dir = path.join(cacheDir, "backups")
  await mkdir(dir, { recursive: true })
  const dest = path.join(dir, `${hash}-${path.basename(file)}`)
  if (!existsSync(dest)) await writeFile(dest, data)
  return dest
}
