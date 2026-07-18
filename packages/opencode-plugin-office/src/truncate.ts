export function truncateForModel(text: string, limit = 24_000): string {
  if (text.length <= limit) return text
  const total = text.length
  return text.slice(0, limit) + `\n[truncated: showing ${limit} of ${total} chars — use mode:"outline" or a target ID to narrow]`
}
