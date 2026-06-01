# google+deepl combined engine

Translate a source `.docx` with **both** Google and DeepL and get a single output that shows
the two translations **side by side**, so a reviewer can compare and pick the better line:

| col 0 (index) | col 1 (source) | col 2 (Google) | col 3 (DeepL) |
|---------------|----------------|----------------|---------------|

It is a thin orchestrator over the existing CLI — it re-implements **none** of the
translation pipeline. Each engine runs as a normal single-engine pass, then the two outputs
are merged by appending DeepL's translation column onto the Google document as a new 4th
column. That keeps it fully compatible with the existing engines, languages, fonts, RTL
handling, and Selenium setup.

## Usage

```bash
python src/combine_google_deepl.py --docxfile "News Scroll - table.docx" --destlang fa
# optional: --srclang en   --showbrowser   --no-landscape   --python /path/to/python
```

The output is **landscape** by default so the four side-by-side columns are comfortably
visible; pass `--no-landscape` to keep portrait. Each engine's name is written in the cell
directly **below** the destination-language label ("Google" under the Google column, "DeepL"
under the DeepL column) — that cell is the blank separator beneath the language row, so the
label never overwrites a subtitle line.

Output, written next to the input:

- `News Scroll - table_PER_Google_Deepl.docx` — both columns, side by side.
- `News Scroll - table_PER_Google.docx` — **degrade** output if DeepL lacks the target
  language, or its pass / the merge fails. The Google result is never lost.

You can also merge two already-translated outputs directly:

```bash
python src/merge_columns.py google_out.docx deepl_out.docx combined.docx DeepL
```

## How it works

1. The source is copied to two scratch files and translated once with `--engine google` and
   once with `--engine deepl` (both with `--split`, so each result is row-distributed). The
   two passes share the source's phrase grouping, so they have identical row geometry.
2. `merge_columns.merge_second_engine_column` appends DeepL's translation cell (`<w:tc>`,
   col 2) onto each Google row as a new column. The cell is a deep-copy of DeepL's, so RTL
   direction, run style, and destination font carry over verbatim. The table grid is widened
   by one column. Each engine's name is written in the empty cell directly below the
   destination-language label (cell (1,2)), and the page is rotated to landscape unless
   `--no-landscape` is given.
3. **Safety:** the merge refuses to run (and the caller degrades to Google-only) on row-count
   drift or an irregular/merged-cell grid in either document, so it never emits scrambled
   side-by-side cells.

## Files

| File | Role |
|------|------|
| `src/merge_columns.py` | Pure `python-docx` column merge (no Selenium, no globals). Importable + runnable standalone. |
| `src/combine_google_deepl.py` | Orchestrator: two single-engine CLI passes + merge, with graceful degrade. |
| `tests/test_merge_columns.py` | Unit tests (`pip install python-docx pytest`). |

## Requirements

`python-docx` (already used by the project). The translation passes need the same Chrome /
Selenium setup the Google and DeepL engines already use.
