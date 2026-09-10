#!/usr/bin/env python3
"""Derive a content-free EIP memo template from a real memo deck.

Why this exists: EIP's memo branding (Nunito, the green/grey palette, the logo and
copyright line on the master) lives inside real deal decks, not in a standalone .potx.
This script lifts the branding out and leaves every trace of the deal behind, so the
template can be committed and shared without leaking a live transaction.

Run it again whenever EIP's branding changes -- point it at the newest memo:

    python scripts/make_template.py NEWEST_MEMO.pptx assets/eip-memo-template.pptx

What gets removed, and why each one matters:

  - every slide                  the deal content itself
  - docProps/app.xml titles      PowerPoint caches the full slide-title list here, which
                                 is a complete table of contents of the source deal
  - ppt/authors.xml              deal-team names and @energyimpactpartners.com addresses
  - ppt/changesInfos/            co-authoring history, same names and addresses
  - ppt/revisionInfo.xml         revision identifiers tied to the source document
  - customXml/                   SharePoint content-type metadata and Bloomberg add-in
                                 shape maps keyed to the deleted slides
  - ppt/tags/                    Macabacus/DocHarmony add-in tags describing the source
                                 deck's own section structure
  - p14:sectionLst               stale section list still naming the deleted slide IDs
  - non-EIP pictures on master   the target company's logo sits on the master alongside
                                 EIP's, and survives slide deletion

The script always finishes with a scan (see verify_clean) and exits non-zero if anything
on the denylist survives, so a silent leak fails loudly instead of shipping.
"""

import argparse
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from pptx import Presentation

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Parts that carry provenance rather than design. Prefixes, matched against the
# archive path.
DROP_PREFIXES = (
    "ppt/authors.xml",
    "ppt/changesInfos/",
    "ppt/revisionInfo.xml",
    "ppt/tags/",
    "customXml/",
)

# Terms that must not survive into a shareable template. Deal-specific names are
# supplied by the caller; these are the structural giveaways that apply to any memo.
DEFAULT_DENYLIST = [
    r"@energyimpactpartners\.com",
    r"\bEBITDA\b",
    r"\bIRR\b",
    r"\bMOIC\b",
    r"Sources & Uses",
    r"Opportunity Summary",
]

MINIMAL_APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" \
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
<Application>Microsoft Office PowerPoint</Application>\
<PresentationFormat>Widescreen</PresentationFormat>\
<Slides>0</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips>\
<ScaleCrop>false</ScaleCrop><Company>Energy Impact Partners</Company>\
<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>\
</Properties>"""

MINIMAL_CUSTOM_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" \
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>"""


def strip_slides_and_logos(src: Path, tmp_pptx: Path, keep_pattern: str) -> None:
    """Delete every slide, then drop any picture on the master that isn't EIP's own."""
    prs = Presentation(str(src))

    sld_id_lst = prs.slides._sldIdLst
    for sld in list(sld_id_lst):
        prs.part.drop_rel(sld.get(f"{R_NS}id"))
        sld_id_lst.remove(sld)

    keep_re = re.compile(keep_pattern, re.I)
    for master in prs.slide_masters:
        for shape in list(master.shapes):
            if shape.shape_type != 13:  # PICTURE
                continue
            descr = shape._element.nvPicPr.cNvPr.get("descr") or ""
            name = shape.name or ""
            if keep_re.search(descr) or keep_re.search(name):
                print(f"  keep picture on master: {descr[:70]!r}")
            else:
                print(f"  DROP picture on master: {descr[:70]!r}")
                shape._element.getparent().remove(shape._element)

    cp = prs.core_properties
    cp.title = "EIP Investment Memo Template"
    cp.author = "Energy Impact Partners"
    cp.last_modified_by = "Energy Impact Partners"
    for field in ("subject", "comments", "category", "keywords", "identifier", "revision"):
        try:
            setattr(cp, field, 1 if field == "revision" else "")
        except (ValueError, TypeError):
            pass

    prs.save(str(tmp_pptx))


def _resolve(base_part: str, target: str) -> str:
    """Resolve a relationship Target against the part that declares it."""
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(posixpath.dirname(base_part))  # strip _rels/
    return posixpath.normpath(posixpath.join(base_dir, target))


def _should_drop(part: str) -> bool:
    return any(part == p or part.startswith(p) for p in DROP_PREFIXES)


def _kept_rel_ids(part: str, blobs, dropped) -> set:
    """Relationship IDs still valid for `part` after the dropped parts are removed."""
    rels_name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    rels_blob = blobs.get(rels_name)
    if rels_blob is None:
        return set()
    kept = set()
    for m in re.finditer(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels_blob.decode("utf8", "ignore")):
        if _resolve(rels_name, m.group(2)) not in dropped:
            kept.add(m.group(1))
    return kept


def scrub_package(tmp_pptx: Path, out: Path) -> None:
    """Rewrite the archive without the provenance parts, patching every reference."""
    with zipfile.ZipFile(tmp_pptx) as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}

    dropped = {n for n in names if _should_drop(n)}

    # Drop media that nothing references any more (the company logo we just removed).
    referenced = set()
    for name, blob in blobs.items():
        if name.endswith(".rels") and name not in dropped:
            for m in re.finditer(r'Target="([^"]+)"', blob.decode("utf8", "ignore")):
                referenced.add(_resolve(name, m.group(1)))
    for name in names:
        if name.startswith("ppt/media/") and name not in referenced:
            dropped.add(name)

    for name in sorted(dropped):
        print(f"  drop part: {name}")

    out_blobs = {}
    for name, blob in blobs.items():
        if name in dropped:
            continue

        if name == "docProps/app.xml":
            blob = MINIMAL_APP_XML.encode("utf8")
        elif name == "docProps/custom.xml":
            blob = MINIMAL_CUSTOM_XML.encode("utf8")
        elif name == "[Content_Types].xml":
            text = blob.decode("utf8")
            for part in dropped:
                text = re.sub(
                    r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(part), "", text
                )
            blob = text.encode("utf8")
        elif name.endswith(".rels"):
            text = blob.decode("utf8")
            for m in list(re.finditer(r"<Relationship\b[^>]*/>", text)):
                tgt = re.search(r'Target="([^"]+)"', m.group(0))
                if tgt and _resolve(name, tgt.group(1)) in dropped:
                    text = text.replace(m.group(0), "")
            blob = text.encode("utf8")
        elif name.endswith(".xml") and name.startswith("ppt/"):
            text = blob.decode("utf8")

            if name == "ppt/presentation.xml":
                # Stale section list still enumerating the deleted slides.
                text = re.sub(
                    r"<p:ext uri=\"\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}\">.*?</p:ext>",
                    "",
                    text,
                    flags=re.S,
                )

            # An element pointing at a relationship we deleted (e.g. <p:tags r:id="rId1"/>,
            # left behind by removing the add-in tag parts) must go entirely. Merely
            # stripping the r:id attribute leaves a schema-invalid element, and PowerPoint
            # reports the whole file as corrupt.
            kept = _kept_rel_ids(name, blobs, dropped)
            for m in list(re.finditer(r"<\w+:\w+\b[^>]*\br:id=\"(rId\d+)\"[^>]*/>", text)):
                if m.group(1) not in kept:
                    text = text.replace(m.group(0), "")
            blob = text.encode("utf8")

        out_blobs[name] = blob

    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in out_blobs.items():
            zout.writestr(name, blob)


# Human-readable payloads inside an OOXML part. Scanning raw XML instead would flag
# every OOXML element that happens to share a word with the denylist -- a company named
# "Guide" collides with <p:guideLst>, one named "Line" with <a:ln>. Everything a reader
# or a metadata pane can actually surface lives in a text node or one of these attributes.
TEXT_NODES = re.compile(
    r"<(?:a:t|vt:lpstr|dc:\w+|cp:\w+|ds:\w+)>([^<]+)</", re.I
)
TEXT_ATTRS = re.compile(
    r'\b(?:name|descr|title|author|userId|initials|Target|val)="([^"]*)"', re.I
)


def readable_strings(blob: bytes):
    """Yield the strings in a part that a person could actually read."""
    text = blob.decode("utf8", "ignore")
    for m in TEXT_NODES.finditer(text):
        yield m.group(1)
    for m in TEXT_ATTRS.finditer(text):
        yield m.group(1)


def verify_clean(path: Path, extra_terms) -> bool:
    """Scan the readable content of every part for anything identifying the source deal."""
    patterns = list(DEFAULT_DENYLIST) + [r"\b%s\b" % re.escape(t) for t in extra_terms]
    findings = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            for s in readable_strings(z.read(name)):
                for pat in patterns:
                    if re.search(pat, s, re.I):
                        findings.append((name, pat, s[:120]))

    prs = Presentation(str(path))
    print(f"\nTemplate: {path}  ({path.stat().st_size / 1024:.0f} KB)")
    print(f"  slides: {len(prs.slides)}  layouts: {len(prs.slide_masters[0].slide_layouts)}")

    if findings:
        print(f"\n  LEAK: {len(findings)} match(es) of denylisted terms survived:")
        for name, pat, ctx in findings[:20]:
            print(f"    [{name}] /{pat}/ ...{ctx}...")
        return False
    print("  scan: clean -- no denylisted terms found")
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", type=Path, help="a real EIP memo .pptx to lift branding from")
    ap.add_argument("output", type=Path, help="where to write the content-free template")
    ap.add_argument(
        "--keep-pattern",
        default=r"energyimpactpartners|\bEIP\b",
        help="pictures on the master matching this (name or alt-text) are kept; all others "
        "are dropped as deal-specific. Check the keep/DROP report before trusting the result.",
    )
    ap.add_argument(
        "--deny",
        nargs="*",
        default=[],
        help="extra terms to fail on, e.g. the target company and sponsor names",
    )
    args = ap.parse_args()

    print(f"Deriving template from {args.source.name}")
    with tempfile.TemporaryDirectory() as td:
        tmp_pptx = Path(td) / "stripped.pptx"
        strip_slides_and_logos(args.source, tmp_pptx, args.keep_pattern)
        scrub_package(tmp_pptx, args.output)

    if not verify_clean(args.output, args.deny):
        print("\nRefusing to ship a template that still names the source deal.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
