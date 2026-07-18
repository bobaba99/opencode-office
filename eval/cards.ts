import path from "node:path"
import * as checks from "./score"

export type CardScore = { success: boolean; fidelity: boolean; notes: string[] }

export type Card = {
  id: string
  files: string[]
  prompt: (arena: string) => string
  check: (arena: string) => Promise<CardScore>
}

// Prompts name absolute paths and the concrete change; they deliberately do NOT name
// tools (office_read/office_edit/etc.) — tool selection is what the battery measures.
export const CARDS: Card[] = [
  {
    id: "docx-replace",
    files: ["edit-report.docx"],
    prompt: (arena) =>
      `In ${path.join(arena, "edit-report.docx")}, change the word 'strong' to 'robust' in the growth sentence. Change nothing else.`,
    check: checks.checkDocxReplace,
  },
  {
    id: "docx-insert",
    files: ["report.docx"],
    prompt: (arena) =>
      `In ${path.join(arena, "report.docx")}, add a new section right after the 'Regional Breakdown' heading: a level-2 heading 'Next Steps' followed by bullet points 'Hire' and 'Ship'.`,
    check: checks.checkDocxInsert,
  },
  {
    id: "docx-table",
    files: ["report.docx"],
    prompt: (arena) => `In ${path.join(arena, "report.docx")}, update the EMEA revenue figure in the table to $5.1M.`,
    check: checks.checkDocxTable,
  },
  {
    id: "docx-create",
    files: [],
    prompt: (arena) =>
      `Create ${path.join(arena, "summary.docx")}: a heading 'Weekly Summary', a paragraph 'All systems nominal.', then bullet points 'Uptime 99.9%' and 'Zero incidents'.`,
    check: checks.checkDocxCreate,
  },
  {
    id: "pptx-retitle",
    files: ["deck.pptx"],
    prompt: (arena) =>
      `In ${path.join(arena, "deck.pptx")}, retitle the second slide to 'Q3 Highlights' and set its speaker notes to 'Two minutes max.'`,
    check: checks.checkPptxRetitle,
  },
  {
    id: "pptx-insert",
    files: ["deck.pptx"],
    prompt: (arena) =>
      `In ${path.join(arena, "deck.pptx")}, insert a new slide titled 'Agenda' with bullet points 'Numbers' and 'Risks' directly after the title slide, using the 'Title and Content' layout.`,
    check: checks.checkPptxInsert,
  },
  {
    id: "pptx-image",
    files: ["edit-deck.pptx", "swap.png"],
    prompt: (arena) =>
      `In ${path.join(arena, "edit-deck.pptx")}, duplicate the third slide (the one with the picture) and on the DUPLICATE swap the picture for ${path.join(arena, "swap.png")}. Leave the original slide untouched.`,
    check: checks.checkPptxImage,
  },
  {
    id: "pptx-create",
    files: ["edit-deck.pptx"],
    prompt: (arena) =>
      `Create ${path.join(arena, "plan.pptx")} using ${path.join(arena, "edit-deck.pptx")} as the template, with exactly two slides: 'Kickoff' on the 'Title Slide' layout, and 'Timeline' on 'Title and Content' with bullets 'Design', 'Build', 'Ship'.`,
    check: checks.checkPptxCreate,
  },
]
