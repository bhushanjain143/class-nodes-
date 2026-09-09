"""Remove speaker notes from a deck, keeping a local copy that still has them."""
import shutil
import sys

from pptx import Presentation


def strip(path, keep_master=True):
    if keep_master:
        master = path.replace(".pptx", "_TeacherNotes.pptx")
        shutil.copyfile(path, master)
        print(f"kept notes copy: {master}")

    prs = Presentation(path)
    dropped = 0
    for slide in prs.slides:
        if not slide.has_notes_slide:
            continue
        part = slide.part
        for rid, rel in list(part.rels.items()):
            if rel.reltype.endswith("/notesSlide"):
                part.drop_rel(rid)
                dropped += 1
    prs.save(path)

    check = Presentation(path)
    remaining = [i + 1 for i, s in enumerate(check.slides) if s.has_notes_slide]
    print(f"{path}: removed {dropped} notes pages, "
          f"slides still carrying one: {remaining if remaining else 'none'}")
    return not remaining


if __name__ == "__main__":
    ok = all(strip(p) for p in sys.argv[1:])
    sys.exit(0 if ok else 1)
