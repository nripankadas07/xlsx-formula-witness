# xlsx-formula-witness

Inspect XLSX shared-formula groups and expand supported A1 references with structural findings and master-cell evidence.

For spreadsheet export developers inspecting missing or surprising shared-formula cells.

## Quickstart

Python3.10+; no external service or runtime dependencies. Linux shell:

```sh
git clone https://github.com/nripankadas07/xlsx-formula-witness.git
cd xlsx-formula-witness
python -m venv .venv
. .venv/bin/activate
python -m pip install .
xlsx-formula-witness --make-example demo.xlsx
xlsx-formula-witness demo.xlsx
```

The generated real workbook expands B3 to A3+$A$1. [Expected JSON](example-output.json).

## Input and interpretation

Run `xlsx-formula-witness --make-example demo.xlsx` first to create an invented real OOXML workbook. Creation refuses to overwrite an existing file. Inspection is read only. No user data or macros are executed.

Read-only conventional xl/worksheets/*.xml parts in transitional OOXML; 4MiB compressed AND total uncompressed ZIP,100 entries,16 worksheets,5000 cells per sheet,8192 formula characters. Strict UTF-8 worksheet XML, no DTD/entities/duplicate archive paths or cells. Master is the unique ref-bearing shared member; follower text is overridden. Outside-range members and overlapping master ranges are findings. Lexical A1 translation supports absolute/relative coordinates, ranges, quoted strings, numbers, function tokens, operators and TRUE/FALSE. Names, sheet/external/table references, translated out-of-bounds references are reported unsupported. This is NOT a formula grammar validator/evaluator, cache-freshness check, full XLSX validator or relationship resolver. Nonshared/array/data-table formulas are listed but not analyzed. Missing nonshared cells inside a shared range are not invented or treated as shared. No Excel application execution.

Exit0: completed result without a finding/conflict; exit1: documented schema/shared-formula/stockout finding or merge conflict; exit2: malformed/unsupported input, stderr and no partial JSON report. Event-time drops are normal policy decisions and use exit0. A completed report is evidence only about the supplied model and supported subset. No automatic live-system changes.

## Verification

```sh
python -m unittest -v
python -m compileall -q xlsx_formula_witness.py
python -m pip wheel --no-deps . -w dist
```

Shared/absolute/string/function/range translation; real ZIP roundtrip; missing/outside/duplicate/overlapping masters; follower override; archive/XML rejection. Optional oracle_check.py compared15 translations and a real shared workbook with installed openpyxl3.1.5 locally; not a CI/runtime dependency. [Validation record](VALIDATION.md). CI checks Python3.10/3.12/3.14 and a wheel-installed console from outside the checkout. Windows/macOS and production-size workloads are not verified. No benchmark or superiority claim.

## Scope and provenance

Alpha utility, original AI-assisted implementation and invented fixtures. Standard algorithms were reimplemented; no third-party implementation/prose copied. MIT. [Candidate brief](BRIEF.json), [dated research and alternatives](RESEARCH.md), [support](SUPPORT.md), [contribution guide](CONTRIBUTING.md), [security](SECURITY.md).

Existing CSV and accounting tools do not inspect OOXML shared master/follower structures. This reads a real spreadsheet package and never evaluates or changes it.
