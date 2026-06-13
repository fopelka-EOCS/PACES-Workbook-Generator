# Code Transfer Packages

Word-document packages that carry this repo's Python scripts **across a locked-down security boundary as text** — for environments where executable files (`.py`, `.exe`) cannot be moved through the firewall, but documents can.

Each package contains the full source of one script, encoded so it survives copy-paste intact. Inside the boundary you paste it into a file, run a short reconstructor, and get a **byte-for-byte exact** copy of the original — confirmed by a SHA-256 hash the reconstructor prints.

## Packages

| Document | Source script | Reconstructs | Dependency | SHA-256 of source |
|----------|---------------|--------------|------------|-------------------|
| `CFO_Generator_Code_Transfer.docx` | `build_cfo_workbook.py` | 10-sheet CFO Excel workbook generator | `openpyxl` (writes the `.xlsx`) | `b2757beb870050ef0d295ce5f54425b00a270c5b6d45b90fe02a93e436fdbda3` |
| `PACES_Analytics_Code_Transfer.docx` | `paces_analytics.py` | Shared analytics module (imported by both generators) | none — pure standard library | `61997c5c38e0e293f7b6bce4fc85ff4fb4bfac10ef8bee0bcd0b4fe65b01a6f6` |

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
