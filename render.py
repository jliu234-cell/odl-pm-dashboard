#!/usr/bin/env python3
"""Bake data.json into the self-contained index.html (no external assets).

Replaces the `/*__DATA__*/` placeholder in template.html with the built data so
the result is a single file anyone can open from Drive/Canvas, offline.
Called automatically at the end of build.py; can also be run standalone.
"""
import os, json

HERE = os.path.dirname(os.path.abspath(__file__))


def render():
    with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
        tpl = f.read()
    with open(os.path.join(HERE, "data.json"), encoding="utf-8") as f:
        data = f.read()
    # embed safely: close-script sequences can't appear inside a JSON string,
    # but guard anyway by escaping the only sequence that could end the tag.
    data = data.replace("</", "<\\/")
    html = tpl_replace(tpl, data)
    out = os.path.join(HERE, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    return out, len(html)


def tpl_replace(tpl, data):
    marker = "/*__DATA__*/"
    if marker not in tpl:
        raise SystemExit("template.html missing /*__DATA__*/ marker")
    return tpl.replace(marker, "window.DATA = " + data + ";")


if __name__ == "__main__":
    out, n = render()
    print(f"wrote {out}  ({n // 1024} KB)")
