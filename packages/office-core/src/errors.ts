export class OfficeError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly hint: string,
  ) {
    super(message)
    this.name = "OfficeError"
  }
}

export function toToolError(err: unknown): { code: string; message: string; hint: string } {
  if (err instanceof OfficeError) return { code: err.code, message: err.message, hint: err.hint }
  const message = err instanceof Error ? err.message : String(err)
  return {
    code: "INTERNAL",
    message,
    hint: "Likely a bug in opencode-office. Re-run once to confirm, then report it.",
  }
}
