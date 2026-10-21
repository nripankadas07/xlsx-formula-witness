# Research — 21 October 2026

Demand is **inferred**, not a verified request for this tool. A shared follower can omit its formula text, making raw ZIP inspection misleading; caches can be stale.

[Exact Most-stars search](https://github.com/search?q=xlsx%20formula&type=repositories&s=stars&o=desc), observed 2026-10-21T06:19:16.581Z. Highest-star relevant comparable found in the inspected pool: `qax-os/excelize`. This is not an exhaustive global ranking.

| Comparable | Stars | Last push | License metadata | Observed workflow |
| --- | ---: | --- | --- | --- |
| [qax-os/excelize](https://github.com/qax-os/excelize) | 21008 | 2026-10-20T03:36:41Z | BSD-3-Clause | Broad Go Excel read/write and calculation library; current README requires Go1.26+. |
| [jmcnamara/XlsxWriter](https://github.com/jmcnamara/XlsxWriter) | 3983 | 2026-08-04T23:41:16Z | BSD-2-Clause | Python XLSX writer with broad formatting/chart/formula output. |
| [apache/poi](https://github.com/apache/poi) | 2278 | 2026-10-19T17:58:20Z | unknown | Java Office binary/OOXML library and shared-formula token handling. |

## Choice and tradeoffs

Inspect XLSX shared-formula groups and expand supported A1 references with structural findings and master-cell evidence. Existing CSV and accounting tools do not inspect OOXML shared master/follower structures. This reads a real spreadsheet package and never evaluates or changes it.

The established alternatives are broader and often the better choice for their supported ecosystems. This tool provides one bounded inspectable report with an invented example and no runtime dependencies. READMEs/docs were read for install and example workflow; competitor clean install, time to first useful result, reliability, support response and performance were not measured. Own tests prove only the documented subset. Stars are research context, not performance or adoption evidence.

Keyword search coverage is narrow: Avro results include generators/profilers, event-time results include general streaming platforms, inventory-expiry results include unrelated homelab/certificate inventory, and three-way results include plugins and unrelated scraped text. Known relevant established tools were added directly. Spreadsheet comparison uses Excelize rather than an obsolete GitHub openpyxl mirror.

[Exact metadata and sampled current issues/PRs](RESEARCH_PROVENANCE.json). The samples are context, not verified competitor defects or requests for our new tool. Relevant Avro compatibility, POI SharedFormula and Git merge-file source was inspected; standard algorithms are original reimplementations. Microsoft CellFormula and Beam semantics informed the bounded contracts. No third-party code, branding or prose is copied.

Primary references:
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
- https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas

Optional local openpyxl3.1.5 oracle: 15 supported translations and reading our real shared-formula workbook agreed. This does not establish full Excel compatibility. `python oracle_check.py` requires a separately installed openpyxl; it is not needed by the tool or CI.
