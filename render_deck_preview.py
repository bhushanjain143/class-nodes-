"""Render selected PowerPoint slides to PNG using the installed PowerPoint app."""
import os
import sys
import shutil

import win32com.client


def render(pptx_path, out_dir, slide_numbers):
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir, ignore_errors=True)
    os.makedirs(out_dir, exist_ok=True)

    app = win32com.client.Dispatch("PowerPoint.Application")
    deck = app.Presentations.Open(pptx_path, WithWindow=False)
    try:
        for n in slide_numbers:
            target = os.path.join(out_dir, f"slide_{n:02d}.png")
            deck.Slides(n).Export(target, "PNG", 1600, 900)
            print(target)
    finally:
        deck.Close()
        app.Quit()


if __name__ == "__main__":
    deck_path = sys.argv[1]
    output = sys.argv[2]
    numbers = [int(x) for x in sys.argv[3].split(",")]
    render(deck_path, output, numbers)
