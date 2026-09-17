# Validation scope

The public Beta validates reusable extraction and publication behavior rather
than certifying every SEC issuer.

The current regression suite covers:

- reported, derived and calculated value provenance;
- exact derivation operands, dates, units and filing references;
- standalone-quarter reconstruction only when source intervals prove it;
- nil, missing and unsupported values remaining distinct from zero;
- dimensional facts not being promoted into consolidated totals;
- evidence-gated financial front sheets;
- source, lineage, semantic review and QA workbook sheets;
- workbook manifest verification and fault injection.

The September 2026 LITE review covered 1,139 displayed financial cells, 1,074
unique data points, seven fault-injection cases and 25 Excel-exported print
pages. That result applies to the reviewed artifact and filings only. Industry
mapping remained partial and company-specific KPIs were not configured.

To run the local checks:

```bash
python -m pytest -q -p no:cacheprovider
ruff check pipeline_audit.py excel_lineage.py sec_data_cli.py review_luna tests tools
mypy --ignore-missing-imports pipeline_audit.py excel_lineage.py sec_data_cli.py
```
