#!/usr/bin/env python3
"""google+deepl combined engine — translate once with Google, once with DeepL, show both.

This is a thin orchestrator over the existing CLI (``machine-translate-docx.py``): it runs the
SAME source file through Google (primary) and DeepL (secondary) as two ordinary single-engine
passes, then merges DeepL's translation column onto the Google docx as a new 4th column
(``merge_columns.py``) so the reader sees both engines side by side:

    col 0 (index) . col 1 (source) . col 2 (Google) . col 3 (DeepL)

It re-implements none of the translation pipeline — each pass is a normal CLI run — so it stays
fully compatible with the existing engines, languages, and Selenium setup. If DeepL lacks the
target language, or its pass / the merge fails, it degrades to the Google-only output rather
than losing the result.

Usage:
    python combine_google_deepl.py --docxfile FILE.docx --destlang fa [--srclang en] [--showbrowser]

Output (next to the input): ``FILE_<LANG>_Google_Deepl.docx`` (or ``FILE_<LANG>_Google.docx``
on a DeepL/merge degrade).
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

CLI = Path(__file__).with_name("machine-translate-docx.py")
_SAVED_RE = re.compile(r"Saved file name:\s*(.+\.docx)\s*$", re.IGNORECASE | re.MULTILINE)


def _run_pass(engine, src_copy, destlang, srclang, showbrowser, python_exe):
    """Run one single-engine CLI pass on ``src_copy``; return the saved output Path.

    Mirrors how each engine is normally invoked (``--split`` so the result is already
    distributed into rows). Raises RuntimeError on a non-zero exit or if no output is found."""
    cmd = [
        python_exe, str(CLI),
        "--docxfile", str(src_copy),
        "--destlang", destlang,
        "--srclang", srclang,
        "--engine", engine,
        "--split",
        "--silent", "--exitonsuccess",
    ]
    if engine == "google" and showbrowser:
        cmd.append("--showbrowser")
    print("[combine] %s pass: %s" % (engine, " ".join(cmd)))
    proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(CLI.parent))
    sys.stdout.write(proc.stdout or "")
    sys.stderr.write(proc.stderr or "")
    if proc.returncode != 0:
        raise RuntimeError("%s pass exited %d" % (engine, proc.returncode))
    # Prefer the CLI's own "Saved file name:" line; fall back to the deterministic name.
    matches = _SAVED_RE.findall(proc.stdout or "")
    if matches:
        out = Path(matches[-1].strip())
        if out.exists():
            return out
    # Fallback: the CLI saves <stem>_<LANG>.docx next to the input.
    for cand in sorted(src_copy.parent.glob(src_copy.stem + "_*.docx")):
        if cand != src_copy:
            return cand
    raise RuntimeError("%s pass produced no output docx" % engine)


def _lang_suffix(google_out):
    """Extract the ISO-639-2/B suffix the CLI appended, e.g. 'PER' from '..._PER.docx'."""
    stem = google_out.stem
    return stem.rsplit("_", 1)[-1] if "_" in stem else ""


def combine(docxfile, destlang, srclang="en", showbrowser=False, python_exe=None, landscape=True):
    """Run both passes + merge. Returns the served output Path (combined, or Google-only on
    a DeepL/merge degrade). ``landscape`` rotates the combined output so the four columns fit
    (pass False / --no-landscape to keep portrait)."""
    from merge_columns import merge_second_engine_column

    python_exe = python_exe or sys.executable
    src = Path(docxfile).resolve()
    if not src.exists():
        raise FileNotFoundError(src)

    workdir = Path(tempfile.mkdtemp(prefix="combine_gd_"))
    try:
        g_in = workdir / ("google_" + src.name)
        d_in = workdir / ("deepl_" + src.name)
        shutil.copy2(src, g_in)
        shutil.copy2(src, d_in)

        google_out = _run_pass("google", g_in, destlang, srclang, showbrowser, python_exe)
        lang = _lang_suffix(google_out)
        combined = src.with_name("%s_%s_Google_Deepl.docx" % (src.stem, lang))
        google_only = src.with_name("%s_%s_Google.docx" % (src.stem, lang))

        try:
            deepl_out = _run_pass("deepl", d_in, destlang, srclang, showbrowser, python_exe)
        except Exception as exc:
            print("[combine] DeepL pass failed (%s) -- serving Google-only" % exc)
            shutil.copy2(google_out, google_only)
            return google_only

        try:
            merge_second_engine_column(str(google_out), str(deepl_out), str(combined),
                                       primary_label="Google", secondary_label="DeepL",
                                       landscape=landscape)
            print("[combine] merged Google+DeepL -> %s" % combined)
            return combined
        except Exception as exc:
            print("[combine] column merge failed (%s) -- serving Google-only" % exc)
            shutil.copy2(google_out, google_only)
            return google_only
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="google+deepl combined engine (side-by-side columns)")
    ap.add_argument("--docxfile", "-d", required=True, help="Input .docx")
    ap.add_argument("--destlang", "--dl", required=True, help="Target language (2-letter code, e.g. fa)")
    ap.add_argument("--srclang", "-sl", default="en", help="Source language (default en)")
    ap.add_argument("--showbrowser", "-b", action="store_true", help="Show the Chrome browser for the Google pass")
    ap.add_argument("--no-landscape", action="store_true",
                    help="Keep portrait orientation (default rotates the combined output to landscape)")
    ap.add_argument("--python", default=sys.executable, help="Python interpreter for the sub-passes")
    args = ap.parse_args()
    out = combine(args.docxfile, args.destlang, args.srclang, args.showbrowser, args.python,
                  landscape=not args.no_landscape)
    print("Saved file name: %s" % out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
