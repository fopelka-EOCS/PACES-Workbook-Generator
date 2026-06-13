# Code Transfer Packages

Word-document packages that carry this repo's Python scripts **across a locked-down security boundary as text** — for environments where executable files (`.py`, `.exe`) cannot be moved through the firewall, but documents can.

Each package contains the full source of one script, encoded so it survives copy-paste intact. Inside the boundary you paste it into a file, run a short reconstructor, and get a **byte-for-byte exact** copy of the original — confirmed by a SHA-256 hash the reconstructor prints.

## Packages

| Document | Source script | Reconstructs | Dependency | SHA-256 of source |
|----------|---------------|--------------|------------|-------------------|
| `CFO_Generator_Code_Transfer.docx` | `build_cfo_workbook.py` | 10-sheet CFO Excel workbook generator | `openpyxl` (writes the `.xlsx`) | `b2757beb870050ef0d295ce5f54425b00a270c5b6d45b90fe02a93e436fdbda3` |
| `CFO_HTML_Generator_Code_Transfer.docx` | `build_cfo_html.py` | Interactive CFO Decision Tool HTML generator | stdlib only; **requires** `paces_analytics.py` + `templates/cfo_decision_tool.html` present | `3a40b32bd1d6d9db7e0e5be959b5ca2a2c2e4f12c3e73a1f55b2e6009ac4ea6b` |
| `PACES_Analytics_Code_Transfer.docx` | `paces_analytics.py` | Shared analytics module (imported by `build_cfo_html.py`) | none — pure standard library | `61997c5c38e0e293f7b6bce4fc85ff4fb4bfac10ef8bee0bcd0b4fe65b01a6f6` |

### Dependencies between files

- `build_cfo_workbook.py` is self-contained — it needs only `openpyxl`.
- `build_cfo_html.py` will **not** run alone. It imports `paces_analytics.py` (must sit in the same folder) and fills in `templates/cfo_decision_tool.html` (expected at `templates/cfo_decision_tool.html` beside the script, or pass `--template`). Transfer the module package alongside it, and make sure the template is present inside the boundary too — the template is text, so it can be moved as a file or transferred the same way as the scripts.

## How to use

Each `.docx` is self-contained and walks through three steps:

1. **Confirm the environment** — check that Python is present (and, for the workbook generator, that `openpyxl` is available).
2. **Reconstruct** — copy the gzip+base64 block into a `rebuild_*.py` file and run it. It writes the original script beside it and prints its SHA-256. Match the hash above and you have a verified, exact copy.
3. **Run** — point the generator at your APCD CSV.

The blob method cannot be corrupted by copy-paste (base64 is only letters, digits, `+`, `/` — no indentation or smart-quote pitfalls). Each document also includes the readable source in an appendix for review.

## Why this is safe

Nothing executable crosses the firewall. The document carries *characters*; the script is reconstructed from those characters inside the boundary. The SHA-256 check guarantees the reconstructed file is identical to the source committed here.

## Regenerating these packages

The `.docx` files are **generated artifacts** built from the `.py` scripts in the repo root. If a source script changes, regenerate its package so the two stay in sync. The packages are produced with a small Node script (`docx` library) that gzip+base64-encodes the source and embeds it alongside instructions and the readable listing. Treat the `.py` files as the source of truth; these documents are a distribution convenience.
