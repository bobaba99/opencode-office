import { createHash } from "node:crypto"
import path from "node:path"
import { readDocx, readPptx, runWorker } from "../packages/office-core/src"
import { FIXTURE_DIR } from "../packages/office-core/test/fixtures"
import type { CardScore } from "./cards"

async function sha256File(file: string): Promise<string> {
  const data = Buffer.from(await Bun.file(file).arrayBuffer())
  return createHash("sha256").update(data).digest("hex")
}

function fail(notes: string[], score: Partial<CardScore> = {}): CardScore {
  return { success: false, fidelity: false, notes, ...score }
}

// Card 1: docx-replace
export async function checkDocxReplace(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "edit-report.docx")

  let doc: Awaited<ReturnType<typeof readDocx>>
  try {
    doc = await readDocx(file, "full")
  } catch (e) {
    return fail([`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const growth = doc.elements.find((e) => e.type === "paragraph" && e.text.includes("Growth was"))
  if (growth?.type === "paragraph" && growth.text === "Growth was robust this quarter overall.") {
    success = true
  } else {
    notes.push(`expected "Growth was robust this quarter overall.", found ${JSON.stringify(growth && growth.type === "paragraph" ? growth.text : null)}`)
  }

  if (growth) {
    try {
      const probe = await runWorker<{ comment_refs: number }>("docx_probe.py", { file, target: growth.id })
      if (probe.comment_refs !== 1) {
        fidelity = false
        notes.push(`comment_refs expected 1, got ${probe.comment_refs}`)
      }
    } catch (e) {
      fidelity = false
      notes.push(`comment probe failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  } else {
    fidelity = false
    notes.push("could not locate the growth paragraph to probe comment refs")
  }

  const tracked = doc.elements.find((e) => e.type === "paragraph" && e.text.includes("Reviewed text"))
  if (!tracked || tracked.type !== "paragraph" || JSON.stringify(tracked.tracked_insertions) !== JSON.stringify(["with tracked insertion"])) {
    fidelity = false
    notes.push("tracked-insertion paragraph changed or missing")
  }

  const table = doc.elements.find((e) => e.type === "table")
  if (!table || table.type !== "table" || table.text !== "K | V\nalpha | one") {
    fidelity = false
    notes.push(`table cells changed: ${table && table.type === "table" ? table.text : "missing"}`)
  }

  const first = doc.elements[0]
  if (!first || first.type !== "paragraph" || first.style !== "Heading 1" || first.text !== "Edit Playground") {
    fidelity = false
    notes.push("p:0 is no longer the untouched 'Edit Playground' Heading 1")
  }

  return { success, fidelity, notes }
}

// Card 2: docx-insert
export async function checkDocxInsert(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "report.docx")

  let doc: Awaited<ReturnType<typeof readDocx>>
  try {
    doc = await readDocx(file, "content")
  } catch (e) {
    return fail([`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const regionalIdx = doc.elements.findIndex((e) => e.type === "paragraph" && e.text === "Regional Breakdown")
  const nextStepsIdx = doc.elements.findIndex((e) => e.type === "paragraph" && e.text === "Next Steps")

  if (regionalIdx === -1 || nextStepsIdx === -1 || nextStepsIdx <= regionalIdx) {
    notes.push(`missing/out-of-order headings: Regional Breakdown@${regionalIdx} Next Steps@${nextStepsIdx}`)
  } else {
    const nextSteps = doc.elements[nextStepsIdx]!
    const after = doc.elements.slice(nextStepsIdx + 1, nextStepsIdx + 3)
    const bulletTexts = after.map((e) => (e.type === "paragraph" ? e.text : "(table)"))
    const bulletStyles = after.map((e) => (e.type === "paragraph" ? e.style : "(table)"))
    const isHeading2 = nextSteps.type === "paragraph" && nextSteps.style === "Heading 2"
    const bulletsOk = bulletTexts.join(",") === "Hire,Ship" && bulletStyles.every((s) => s === "List Bullet")
    if (isHeading2 && bulletsOk) {
      success = true
    } else {
      notes.push(`Next Steps section malformed: style=${nextSteps.type === "paragraph" ? nextSteps.style : "?"} bullets=${JSON.stringify(bulletTexts)} styles=${JSON.stringify(bulletStyles)}`)
    }
  }

  const expectedPre = [
    "Quarterly Report",
    "Q3 revenue grew 12% year over year.",
    "Regional Breakdown",
    "EMEA led growth for the third consecutive quarter.",
    "Prepared by the finance team.",
  ]
  let lastIdx = -1
  for (const text of expectedPre) {
    const idx = doc.elements.findIndex((e) => e.type === "paragraph" && e.text === text)
    if (idx === -1 || idx <= lastIdx) {
      fidelity = false
      notes.push(`pre-existing paragraph out of order or missing: ${text}`)
    }
    lastIdx = idx
  }
  const table = doc.elements.find((e) => e.type === "table")
  if (!table || table.type !== "table" || table.text !== "Region | Revenue\nEMEA | $4.2M") {
    fidelity = false
    notes.push(`table changed: ${table && table.type === "table" ? table.text : "missing"}`)
  }

  return { success, fidelity, notes }
}

// Card 3: docx-table
export async function checkDocxTable(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "report.docx")

  let doc: Awaited<ReturnType<typeof readDocx>>
  try {
    doc = await readDocx(file, "content")
  } catch (e) {
    return fail([`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const table = doc.elements.find((e) => e.type === "table")
  if (table?.type === "table" && table.text) {
    const rows = table.text.split("\n").map((r) => r.split(" | "))
    if (rows[1]?.[1] === "$5.1M") {
      success = true
    } else {
      notes.push(`cell(1,1) is ${JSON.stringify(rows[1]?.[1])}, expected "$5.1M"`)
    }
    if (rows[0]?.[0] !== "Region" || rows[0]?.[1] !== "Revenue" || rows[1]?.[0] !== "EMEA") {
      fidelity = false
      notes.push("untouched table cells changed")
    }
  } else {
    notes.push("table missing")
    fidelity = false
  }

  const paragraphTexts = [
    "Quarterly Report",
    "Q3 revenue grew 12% year over year.",
    "Regional Breakdown",
    "EMEA led growth for the third consecutive quarter.",
    "Prepared by the finance team.",
  ]
  for (const text of paragraphTexts) {
    if (!doc.elements.some((e) => e.type === "paragraph" && e.text === text)) {
      fidelity = false
      notes.push(`paragraph changed/missing: ${text}`)
    }
  }

  return { success, fidelity, notes }
}

// Card 4: docx-create — no pre-existing content, fidelity is always true.
export async function checkDocxCreate(arena: string): Promise<CardScore> {
  const notes: string[] = []
  const file = path.join(arena, "summary.docx")

  let doc: Awaited<ReturnType<typeof readDocx>>
  try {
    doc = await readDocx(file, "content")
  } catch (e) {
    return { success: false, fidelity: true, notes: [`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`] }
  }

  const expectedTexts = ["Weekly Summary", "All systems nominal.", "Uptime 99.9%", "Zero incidents"]
  const texts = doc.elements.map((e) => (e.type === "paragraph" ? e.text : "(table)"))
  const styles = doc.elements.map((e) => (e.type === "paragraph" ? e.style : "(table)"))
  const textsMatch = texts.length === 4 && expectedTexts.every((t, i) => texts[i] === t)
  const headingOk = typeof styles[0] === "string" && styles[0].startsWith("Heading")
  const bulletsOk = styles[2] === "List Bullet" && styles[3] === "List Bullet"

  const success = textsMatch && headingOk && bulletsOk
  if (!success) {
    notes.push(`unexpected structure: texts=${JSON.stringify(texts)} styles=${JSON.stringify(styles)}`)
  }

  return { success, fidelity: true, notes }
}

// Card 5: pptx-retitle
export async function checkPptxRetitle(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "deck.pptx")

  let deck: Awaited<ReturnType<typeof readPptx>>
  try {
    deck = await readPptx(file, "content")
  } catch (e) {
    return fail([`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const s1 = deck.slides[1]
  if (s1?.title === "Q3 Highlights" && s1.notes === "Two minutes max.") {
    success = true
  } else {
    notes.push(`slide 1 title/notes mismatch: title=${s1?.title} notes=${s1?.notes}`)
  }

  const s0 = deck.slides[0]
  const s0Subtitle = s0?.shapes?.some((sh) => sh.text === "Finance Team") ?? false
  if (s0?.title !== "Q3 Review" || !s0Subtitle) {
    fidelity = false
    notes.push(`slide 0 title/subtitle changed: title=${s0?.title}`)
  }

  return { success, fidelity, notes }
}

// Card 6: pptx-insert
export async function checkPptxInsert(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "deck.pptx")

  let outline: Awaited<ReturnType<typeof readPptx>>
  try {
    outline = await readPptx(file, "outline")
  } catch (e) {
    return fail([`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const titles = outline.slides.map((s) => s.title)
  if (titles.length === 3 && titles[0] === "Q3 Review" && titles[1] === "Agenda" && titles[2] === "Highlights") {
    const content = await readPptx(file, "content", "s:1")
    const slide1 = content.slides[0]!
    const bodyText = (slide1.shapes ?? []).map((sh) => sh.text).join("\n")
    if (slide1.layout === "Title and Content" && bodyText.includes("Numbers") && bodyText.includes("Risks")) {
      success = true
    } else {
      notes.push(`slide 1 layout/body mismatch: layout=${slide1.layout} body=${JSON.stringify(bodyText)}`)
    }
  } else {
    notes.push(`outline mismatch: ${JSON.stringify(titles)}`)
  }

  const content = await readPptx(file, "content")
  const first = content.slides[0]
  const last = content.slides[content.slides.length - 1]
  const firstTexts = [...(first?.shapes ?? []).map((sh) => sh.text)].sort()
  const lastTexts = [...(last?.shapes ?? []).map((sh) => sh.text)].sort()
  const expectedFirst = ["Q3 Review", "Finance Team"].sort()
  const expectedLast = ["Highlights", "Revenue up 12%\nEMEA leads growth"].sort()
  if (JSON.stringify(firstTexts) !== JSON.stringify(expectedFirst)) {
    fidelity = false
    notes.push(`first slide shapes changed: ${JSON.stringify(firstTexts)}`)
  }
  if (JSON.stringify(lastTexts) !== JSON.stringify(expectedLast)) {
    fidelity = false
    notes.push(`last slide shapes changed: ${JSON.stringify(lastTexts)}`)
  }

  return { success, fidelity, notes }
}

// Card 7: pptx-image
type PictureProbe = { pictures: Array<{ id: string; part: string; sha256: string; content_type: string }> }

export async function checkPptxImage(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "edit-deck.pptx")
  const swap = path.join(arena, "swap.png")
  const pristine = path.join(FIXTURE_DIR, "edit-deck.pptx")
  // The picture-only slide named in the prompt ("the third slide") — its own position never
  // shifts as a result of this edit, since duplicating only appends/inserts a NEW slide
  // elsewhere in the deck; the original slide's index is stable regardless of where the copy
  // lands.
  const ORIGINAL_ID = "s:2/sh:0"

  let outline: Awaited<ReturnType<typeof readPptx>>
  let probe: PictureProbe
  let swapHash: string
  let pristineOutline: Awaited<ReturnType<typeof readPptx>>
  let pristineProbe: PictureProbe
  try {
    outline = await readPptx(file, "outline")
    probe = await runWorker<PictureProbe>("pptx_probe.py", { file })
    swapHash = await sha256File(swap)
    pristineOutline = await readPptx(pristine, "outline")
    pristineProbe = await runWorker<PictureProbe>("pptx_probe.py", { file: pristine })
  } catch (e) {
    return fail([`could not read/probe ${file}: ${e instanceof Error ? e.message : String(e)}`])
  }

  const expectedCount = pristineOutline.slides.length + 1
  const countOk = outline.slides.length === expectedCount
  if (!countOk) {
    notes.push(`expected ${expectedCount} slides (original ${pristineOutline.slides.length} + 1 duplicate), got ${outline.slides.length}`)
  }

  const pristineOriginal = pristineProbe.pictures.find((p) => p.id === ORIGINAL_ID)
  const currentOriginal = probe.pictures.find((p) => p.id === ORIGINAL_ID)
  const originalIntact = !!pristineOriginal && !!currentOriginal && currentOriginal.sha256 === pristineOriginal.sha256
  if (!originalIntact) {
    fidelity = false
    notes.push(`original slide's picture (${ORIGINAL_ID}) changed: got ${currentOriginal?.sha256}, expected ${pristineOriginal?.sha256}`)
  }

  // The prompt doesn't pin where the duplicate lands — native duplicate_slide places it right
  // after the source, but a different tool path (e.g. a generic script) could append it at the
  // end instead. Locate it by CONTENT rather than a fixed index: any picture, other than the
  // untouched original, whose sha256 matches swap.png.
  const swapped = probe.pictures.find((p) => p.id !== ORIGINAL_ID && p.sha256 === swapHash)
  if (!swapped) {
    notes.push(`no picture outside ${ORIGINAL_ID} matches swap.png's sha256 (${swapHash})`)
  }

  success = countOk && originalIntact && !!swapped

  return { success, fidelity, notes }
}

// Card 8: pptx-create
export async function checkPptxCreate(arena: string): Promise<CardScore> {
  const notes: string[] = []
  let success = false
  let fidelity = true
  const file = path.join(arena, "plan.pptx")
  const template = path.join(arena, "edit-deck.pptx")

  try {
    const deck = await readPptx(file, "content")
    if (deck.slides.length === 2) {
      const [s0, s1] = deck.slides as [NonNullable<(typeof deck.slides)[0]>, NonNullable<(typeof deck.slides)[0]>]
      const bulletText = (s1.shapes ?? []).map((sh) => sh.text).join("\n")
      const bulletsOk = ["Design", "Build", "Ship"].every((b) => bulletText.includes(b))
      if (s0.title === "Kickoff" && s0.layout === "Title Slide" && s1.title === "Timeline" && s1.layout === "Title and Content" && bulletsOk) {
        success = true
      } else {
        notes.push(`slide mismatch: ${JSON.stringify(deck.slides)}`)
      }
    } else {
      notes.push(`expected 2 slides, got ${deck.slides.length}`)
    }
  } catch (e) {
    notes.push(`could not read ${file}: ${e instanceof Error ? e.message : String(e)}`)
  }

  try {
    const templateHash = await sha256File(template)
    const pristineHash = await sha256File(path.join(FIXTURE_DIR, "edit-deck.pptx"))
    if (templateHash !== pristineHash) {
      fidelity = false
      notes.push("template file (edit-deck.pptx) was modified")
    }
  } catch (e) {
    fidelity = false
    notes.push(`could not hash template: ${e instanceof Error ? e.message : String(e)}`)
  }

  return { success, fidelity, notes }
}
