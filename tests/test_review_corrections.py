import json
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from excel_lineage import _context, _unit, save_pivot_xlsx, verify_workbook_manifest
from pipeline_audit import (
    attach_income_face_evidence,
    enrich_output_audit,
    retain_derivation_inputs,
)
from review_luna.run_review import (
    close,
    front_values,
    sha256,
    xbrl_match,
    xml_cache_for,
)
from sec_data import _validate_and_repair_segment_data


def test_review_uses_configured_edgar_cache_and_ignores_only_xlsx_float_noise(
    tmp_path, monkeypatch
):
    cache = tmp_path / "edgar"
    artifact = cache / "_tcache" / "www.sec.gov" / "filing-000163397826057358-doc.htm-a1b2c3d4"
    artifact.parent.mkdir(parents=True)
    artifact.write_bytes(b"<html/>")
    monkeypatch.setenv("EDGAR_LOCAL_DATA_DIR", str(cache))
    xml_cache_for.cache_clear()
    assert xml_cache_for("0001633978-26-057358") == artifact
    assert close(25.9, Decimal("25.899999999999999"))
    assert not close(25.9, Decimal("25.9001"))
    xml_cache_for.cache_clear()


def test_dimensional_fact_is_not_a_conflict_with_consolidated_amount():
    lineage = {'Source ID': 'F1', 'XBRL Concept': 'Revenues', 'Context': '2024-04-01..2024-06-30',
               'Metric': 'Revenue', 'Unit': 'USD', 'Value': 100}
    fact = {'concept': 'Revenues', 'entity': '1652044', 'kind': 'duration',
            'start': '2024-04-01', 'end': '2024-06-30', 'dimensions': [], 'axes': [],
            'unit': 'USD', 'namespace': 'http://fasb.org/us-gaap/2024', 'value': 100,
            'face_income': True, 'id': 'f1', 'xpath': '/table/fact', 'context': 'c1',
            'raw_lexical': '100', 'scale': '0', 'sign': '', 'artifact_sha256': 'abc', 'decimals': 'INF'}
    segment = dict(fact, value=30, dimensions=['goog:CloudMember'], axes=['us-gaap:StatementBusinessSegmentsAxis'])
    result, _ = xbrl_match(lineage, {'F1': [fact, segment]}, {'F1': {'Accession': '0001652044-24-000001'}})
    assert result == 'PASS'


def test_review_uses_archive_entity_cik_when_accession_prefix_differs():
    lineage = {
        'Source ID': 'F1',
        'XBRL Concept': '{http://fasb.org/us-gaap/2026}Revenues',
        'Context': '2025-06-29..2026-06-27',
        'Metric': 'Revenue',
        'Unit': 'USD',
        'Value': 100,
    }
    fact = {
        'concept': 'Revenues', 'entity': '0001633978', 'kind': 'duration',
        'start': '2025-06-29', 'end': '2026-06-27', 'dimensions': [],
        'axes': [], 'unit': 'USD', 'namespace': 'http://fasb.org/us-gaap/2026',
        'value': Decimal('100'), 'face_income': True, 'id': 'f1',
        'xpath': '/table/fact', 'context': 'c1', 'raw_lexical': '100',
        'scale': '0', 'sign': '', 'artifact_sha256': 'abc', 'decimals': 'INF',
    }
    source = {
        'Accession': '0001628280-26-057358',
        'SEC Link': (
            'https://www.sec.gov/Archives/edgar/data/1633978/'
            '000162828026057358/lite-20260627.htm'
        ),
    }
    result, _ = xbrl_match(lineage, {'F1': [fact]}, {'F1': source})
    assert result == 'PASS'


def test_segment_ytd6_repair_is_derived_with_relocatable_inputs():
    category, label = '4c_Segments_Geographic_Countries', 'Revenue - United States'
    frame = pd.DataFrame(
        {'2024-Q1': [41_100_000.0, 200_000_000.0],
         '2024-Q2': [103_000_000.0, 220_000_000.0]},
        index=pd.MultiIndex.from_tuples([
            (category, label), ('1_Income_Statement', 'Revenue')]))
    common = {
        'Category': category, 'Label': label,
        'Concept': 'RevenueFromContractWithCustomerIncludingAssessedTax',
        'SourceExactUnitKind': 'USD', 'SourceDimensionAxes': 'GeographyAxis',
        'SourceDimensionMembers': 'UnitedStatesMember',
        'SourceKind': 'xbrl', 'SourceClassificationConfidence': 1.0,
    }
    audit = pd.DataFrame([
        dict(common, Period='2024-Q1', Value=41_100_000.0,
             Accession='Q1', Start='2024-01-01', End='2024-03-31',
             SourceReportedValue='41100000', SourceCellIdentity='/q1'),
        dict(common, Period='2024-Q2', Value=103_000_000.0,
             Accession='H1', Start='2024-01-01', End='2024-06-30',
             SourceReportedValue='103000000', SourceCellIdentity='/h1'),
    ])
    result = _validate_and_repair_segment_data(frame, audit)
    assert result.loc[(category, label), '2024-Q2'] == 61_900_000.0
    record = result.attrs['fact_audit'].query("Period == '2024-Q2'").iloc[0]
    assert record.SourceKind == 'derived'
    assert record.Start == '2024-04-01'
    assert record.SourceReportedValue != record.SourceReportedValue  # NaN
    inputs = json.loads(record.SourceInputFacts)
    assert [item['Accession'] for item in inputs] == ['H1', 'Q1']
    assert [item['SourceCellIdentity'] for item in inputs] == ['/h1', '/q1']
    assert sum(Decimal(str(item['Value'])) * Decimal(item['Coefficient'])
               for item in inputs) == Decimal('61900000')


def test_large_reported_segment_quarter_is_not_subtracted_without_interval_proof():
    category, label = '4c_Segments_Geographic_Countries', 'Revenue - United States'
    frame = pd.DataFrame(
        {'2024-Q1': [41_100_000.0], '2024-Q2': [103_000_000.0]},
        index=pd.MultiIndex.from_tuples([(category, label)]))
    common = {
        'Category': category, 'Label': label,
        'Concept': 'RevenueFromContractWithCustomerIncludingAssessedTax',
        'SourceExactUnitKind': 'USD', 'SourceDimensionAxes': 'GeographyAxis',
        'SourceDimensionMembers': 'UnitedStatesMember', 'SourceKind': 'xbrl',
        'SourceClassificationConfidence': 1.0,
    }
    audit = pd.DataFrame([
        dict(common, Period='2024-Q1', Value=41_100_000.0,
             Accession='Q1', Start='2023-07-02', End='2023-09-30',
             SourceReportedValue='41100000', SourceCellIdentity='/q1'),
        # This is a standalone Q2 interval, despite being >1.8x Q1.
        dict(common, Period='2024-Q2', Value=103_000_000.0,
             Accession='Q2', Start='2023-10-01', End='2023-12-30',
             SourceReportedValue='103000000', SourceCellIdentity='/q2'),
    ])
    result = _validate_and_repair_segment_data(frame, audit)
    assert result.loc[(category, label), '2024-Q2'] == 103_000_000.0
    record = result.attrs['fact_audit'].query("Period == '2024-Q2'").iloc[0]
    assert record.SourceKind == 'xbrl'
    assert pd.isna(record.get('SourceDerivation'))


def test_income_face_total_wins_over_supplemental_revenue():
    xml = b'''<html xmlns:ix="http://www.xbrl.org/2013/inlineXBRL"><context id="c"><startDate>2024-04-01</startDate><endDate>2024-06-30</endDate></context>
    <table><ix:nonFraction name="us-gaap:Revenues" contextRef="c" unitRef="usd" scale="6">84,742</ix:nonFraction>
    <ix:nonFraction name="us-gaap:OperatingIncomeLoss" contextRef="c" unitRef="usd"/><ix:nonFraction name="us-gaap:EarningsPerShareBasic" contextRef="c" unitRef="usd"/></table></html>'''
    base = {'Category': '1_Income_Statement', 'Label': 'Revenue', 'Start': '2024-04-01', 'End': '2024-06-30', 'DimCount': 0}
    facts = [dict(base, Concept='Revenues', Value='84742000000'),
             dict(base, Concept='RevenueFromContractWithCustomerExcludingAssessedTax', Value='84640000000')]
    result = attach_income_face_evidence(facts, xml)
    assert result[0]['SourceFaceRevenue'] is True
    assert result[1]['Category'] == '6_Disclosures'
    assert facts[1]['Category'] == '1_Income_Statement'


def test_other_investing_display_applies_source_weight_once():
    frame = pd.DataFrame({'2026-Q1': [-996]}, index=pd.MultiIndex.from_tuples([('3_Cash_Flow', 'Other Investing Activities')]))
    frame.attrs['fact_audit'] = pd.DataFrame([{'Category': '3_Cash_Flow', 'Label': 'Other Investing Activities',
        'Period': '2026-Q1', 'Concept': 'PaymentsForProceedsFromOtherInvestingActivities', 'Value': -996,
        'SourceReportedValue': '996', 'Start': '2026-01-01', 'End': '2026-03-31'}])
    result = enrich_output_audit(frame, {'PaymentsForProceedsFromOtherInvestingActivities': {('CFI', -1)}}, {'CFI'})
    record = result.attrs['fact_audit'].iloc[0]
    assert result.iloc[0, 0] == record.Value == 996
    assert record.SourceDisplaySign == -1


def test_revenue_share_payable_is_money_not_share_count():
    assert _unit('Accrued Revenue Share', {}) == 'USD'
    assert _unit('Shares Outstanding Basic', {}) == 'shares'
    assert _unit('Anything', {'SourceExactUnitKind': 'USD/shares'}) == 'USD/share'


def test_missing_raw_provenance_does_not_erase_original_dimensions():
    context = _context({'Start': '2026-01-01', 'End': '2026-03-31',
                        'SourceEvidenceAxes': float('nan'), 'SourceEvidenceMembers': float('nan'),
                        'SourceDimensionAxes': 'SegmentAxis', 'SourceDimensionMembers': 'Cloud'})
    assert context == '2026-01-01..2026-03-31 | SegmentAxis; Cloud'


def test_other_investing_quarter_uses_original_sec_input_signs():
    concept = 'PaymentsForProceedsFromOtherInvestingActivities'
    frame = pd.DataFrame({'2026-Q2': [-363]}, index=pd.MultiIndex.from_tuples([('3_Cash_Flow', 'Other Investing Activities')]))
    frame.attrs['fact_audit'] = pd.DataFrame([{'Category': '3_Cash_Flow', 'Label': 'Other Investing Activities',
        'Period': '2026-Q2', 'Concept': concept, 'Value': -363, 'Start': '2026-04-01', 'End': '2026-06-30',
        'SourceDerivationYTDValue': -1359, 'SourceDerivationYTDConcept': concept, 'SourceDerivationYTDAccession': 'H1',
        'SourceDerivationQuarterValue': -996, 'SourceDerivationQuarterConcept': concept, 'SourceDerivationQuarterAccession': 'Q1'}])
    facts = [dict(Concept=concept, Accession=accession, Value=value, SourceReportedValue=str(value),
                  Start='2026-01-01', End=end) for accession, value, end in [('H1', 1359, '2026-06-30'), ('Q1', 996, '2026-03-31')]]
    retain_derivation_inputs(frame, facts)
    record = frame.attrs['fact_audit'].iloc[0]
    assert frame.iloc[0, 0] == record.Value == 363
    assert [item['Value'] for item in json.loads(record.SourceInputFacts)] == ['1359', '996']


def small_workbook(path: Path) -> None:
    frame = pd.DataFrame({'2026-Q1': [100, 40]}, index=pd.MultiIndex.from_tuples([('1_Income_Statement', 'Revenue'), ('1_Income_Statement', 'Net Income')]))
    frame.attrs['ticker'] = 'TEST'
    frame.attrs['fact_audit'] = pd.DataFrame([{'Ticker': 'TEST', 'Category': '1_Income_Statement',
        'Label': label, 'Period': '2026-Q1', 'Value': value, 'Accession': '0001652044-26-000048',
        'FilingUrl': 'https://www.sec.gov/test', 'Concept': concept, 'SourceKind': 'xbrl',
        'Start': '2026-01-01', 'End': '2026-03-31', 'SourceClassificationConfidence': 1}
        for label, value, concept in [('Revenue', 100, 'Revenues'), ('Net Income', 40, 'NetIncomeLoss')]])
    save_pivot_xlsx(frame, path)


def test_repeated_overview_number_is_financial_and_mutations_bypass_binary_gate(tmp_path):
    path = tmp_path / 'test.xlsx'
    small_workbook(path)
    assert verify_workbook_manifest(path) == []
    book = load_workbook(path)
    front = front_values(book, {})
    assert any(row['Sheet'] == '00_OVERVIEW' and row['Scope'] == 'financial' for row in front)
    for sheet, cell in [('02_QUARTERLY', 'C6'), ('00_OVERVIEW', 'F7')]:
        original = book[sheet][cell].value
        book[sheet][cell].value = original + 1
        book.save(path)
        manifest_path = path.with_suffix('.manifest.json')
        manifest = json.loads(manifest_path.read_text())
        manifest['workbook'].update(sha256=sha256(path), bytes=path.stat().st_size)
        manifest_path.write_text(json.dumps(manifest))
        errors = verify_workbook_manifest(path)
        assert errors and all('hash does not match' not in error for error in errors)
        book[sheet][cell].value = original
    book.close()
