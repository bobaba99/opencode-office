from _worker import WorkerError
from pptx.oxml.ns import qn


def find_layout(prs, name):
    names = []
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            names.append(layout.name)
            if layout.name == name:
                return layout
    raise WorkerError("LAYOUT_NOT_FOUND", f"No slide layout named {name!r}", f"Available layouts: {', '.join(names)}")


def move_entry(prs, from_index, to_index):
    lst = prs.slides._sldIdLst
    entry = list(lst)[from_index]
    lst.remove(entry)
    remaining = list(lst)
    to_index = max(0, min(to_index, len(remaining)))
    if to_index == len(remaining):
        lst.append(entry)
    else:
        remaining[to_index].addprevious(entry)


def delete_entry(prs, index):
    lst = prs.slides._sldIdLst
    entry = list(lst)[index]
    prs.part.drop_rel(entry.get(qn("r:id")))
    lst.remove(entry)
