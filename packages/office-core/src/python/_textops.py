def para_text(p):
    return "".join(run.text or "" for run in p.runs)


def replace_in_paragraph(p, anchor, replacement):
    """Replace the first occurrence of anchor, preserving run formatting.

    The replacement text lands in the first run that overlaps the anchor
    (inheriting its formatting); every other overlapped run keeps only the
    parts of its text outside the anchor. Returns True if a replacement
    happened.
    """
    text = para_text(p)
    start = text.find(anchor)
    if start < 0:
        return False
    end = start + len(anchor)
    pos = 0
    replaced = False
    for run in p.runs:
        run_text = run.text or ""
        run_start, run_end = pos, pos + len(run_text)
        pos = run_end
        if run_end <= start or run_start >= end:
            continue
        head = run_text[: max(0, start - run_start)]
        tail = run_text[max(0, min(len(run_text), end - run_start)):]
        if not replaced:
            run.text = head + replacement + tail
            replaced = True
        else:
            run.text = head + tail
    return replaced
