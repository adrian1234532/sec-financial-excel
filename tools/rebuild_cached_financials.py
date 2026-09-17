"""Replay a trusted local extraction checkpoint into a new, immutable output directory."""

import argparse
import gzip
import json
import pickle
import socket
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import sec_data  # noqa: E402
from excel_lineage import save_pivot_xlsx, verify_workbook_manifest  # noqa: E402
from pipeline_audit import (  # noqa: E402
    attach_income_face_evidence,
    enrich_output_audit,
    require_complete_output_evidence,
    retain_derivation_inputs,
)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--checkpoint', type=Path, required=True, help='Trusted local pickle; never accept downloaded pickle files')
    parser.add_argument('--ticker', required=True)
    parser.add_argument('--year-end-month', type=int, default=12)
    parser.add_argument('--sec-cache', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--frames', type=Path, help='Replay only the export stage from trusted local engine frames')
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=False)

    def offline(*_args, **_kwargs):
        raise OSError('Offline replay forbids network access')

    socket.socket.connect = offline
    socket.create_connection = offline
    with args.checkpoint.open('rb') as stream:
        checkpoint = pickle.load(stream)
    sec_data._restore_native_mutable_state(checkpoint['native_state'])
    by_filing = defaultdict(list)
    for fact in checkpoint['all_facts']:
        by_filing[fact.get('Accession')].append(fact)
    facts = []
    evidence = []
    documents = {}
    import hashlib
    for accession, rows in by_filing.items():
        compact = str(accession).replace('-', '')
        candidates = [path for path in args.sec_cache.glob(f'*{compact}*htm-*') if not path.name.endswith('.meta')]
        if len(candidates) != 1:
            raise ValueError(f'{accession}: expected one explicitly identified cached SEC document, found {len(candidates)}')
        document = gzip.decompress(candidates[0].read_bytes())
        documents[accession] = document
        facts.extend(attach_income_face_evidence(rows, document))
        evidence.append({'accession': accession, 'url': rows[0].get('FilingUrl'),
                         'cache_file': candidates[0].name, 'raw_sha256': hashlib.sha256(document).hexdigest()})
    print(f'Building {args.ticker} from {len(facts)} cached facts, {len(evidence)} SEC documents', flush=True)
    if args.frames:
        with args.frames.open('rb') as stream:
            quarterly, annual = pickle.load(stream)
        roots = sec_data._CF_OPERATING_PARENTS | sec_data._CF_INVESTING_PARENTS | sec_data._CF_FINANCING_PARENTS
        for frame in (quarterly, annual):
            print(f'Annotating {len(frame.attrs["fact_audit"])} frozen audit rows', flush=True)
            audit = frame.attrs['fact_audit']
            rows = []
            for accession, group in audit.groupby('Accession', dropna=False):
                records = group.to_dict('records')
                rows.extend(attach_income_face_evidence(records, documents[accession]) if accession in documents else records)
            frame.attrs['fact_audit'] = pd.DataFrame(rows)
            print('Retaining recursive operands and applying evidence eligibility', flush=True)
            retain_derivation_inputs(frame, facts)
            enrich_output_audit(frame, sec_data.GLOBAL_CALC_PARENT, roots)
            require_complete_output_evidence(frame)
    else:
        quarterly = sec_data.build_pivoted_data(facts, args.ticker, args.year_end_month)
        state = sec_data._snapshot_native_mutable_state()
        try:
            annual = sec_data.build_annual_pivoted_data(facts, args.ticker, args.year_end_month, limit=5)
        finally:
            sec_data._restore_native_mutable_state(state)
    workbook = args.output_dir / f'{args.ticker}_financials_corrected.xlsx'
    print('Writing workbook and frozen manifest', flush=True)
    save_pivot_xlsx(quarterly, workbook, annual_pivot=annual)
    with (args.output_dir / 'frames.pkl').open('wb') as stream:
        pickle.dump((quarterly, annual), stream, pickle.HIGHEST_PROTOCOL)
    (args.output_dir / 'replay.json').write_text(json.dumps({'timestamp': datetime.now(timezone.utc).isoformat(),
        'checkpoint_sha256': hashlib.sha256(args.checkpoint.read_bytes()).hexdigest(),
        'engine_frames_sha256': hashlib.sha256(args.frames.read_bytes()).hexdigest() if args.frames else None, 'evidence': evidence,
        'network': 'FORBIDDEN', 'artifact_status': 'DRAFT'}, indent=2), encoding='utf-8')
    errors = verify_workbook_manifest(workbook)
    print(json.dumps({'workbook': str(workbook), 'integrity_errors': errors}), flush=True)
    if errors:
        raise SystemExit(2)


if __name__ == '__main__':
    main()
