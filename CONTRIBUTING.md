# Contributing

Bug reports and focused pull requests are welcome. A useful issue includes the
ticker, accession number, affected period, workbook cell, expected SEC value,
and the source URL or XBRL concept.

## Development checks

```bash
python -m pytest -q
ruff check pipeline_audit.py excel_lineage.py sec_data_cli.py review_luna tests tools
mypy --ignore-missing-imports pipeline_audit.py excel_lineage.py sec_data_cli.py
```

Fix extraction and validation rules in the reusable pipeline. Do not patch one
generated workbook or weaken a check to make a fixture pass. Unsupported facts
should remain blank and enter the review queue.

Never commit SEC contact identities, API keys, local caches, generated output,
or untrusted pickle files.
