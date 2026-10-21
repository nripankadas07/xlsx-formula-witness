# Validation — 21 October 2026

Shared/absolute/string/function/range translation; real ZIP roundtrip; missing/outside/duplicate/overlapping masters; follower override; archive/XML rejection. Optional oracle_check.py compared15 translations and a real shared workbook with installed openpyxl3.1.5 locally; not a CI/runtime dependency.

All behavior tests passed locally on Python3.12/Linux. Local compile, offline wheel build, new isolated venv installation and installed console from outside checkout passed; failure cases and input preservation passed. Local runtime: Python3.12.14, Linux x86_64. Remote Python3.10/3.12/3.14 matrix and public heads must be verified before LIVE. No competitor benchmark.
