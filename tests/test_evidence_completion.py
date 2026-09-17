import gzip
import json
from decimal import Decimal

import pandas as pd
import pytest
from lxml import etree

from pipeline_audit import (
    attach_income_face_evidence,
    require_complete_output_evidence,
    retain_derivation_inputs,
)
from review_luna.run_review import table_match


def test_html_cost_cannot_inherit_revenue_semantics():
    document = b'''<html xmlns:ix="http://www.xbrl.org/2013/inlineXBRL" xmlns:g="http://fasb.org/us-gaap/2024">
    <context id="c"><startDate>2024-01-01</startDate><endDate>2024-03-31</endDate></context>
    <unit id="u"><measure>g:USD</measure></unit><table><tr><td>Cloud</td><td>
    <ix:nonFraction name="g:RevenueFromContractWithCustomerExcludingAssessedTax" contextRef="c" unitRef="u" scale="6">100</ix:nonFraction>
    </td></tr></table></html>'''
    row = {'Label': 'Other Costs and Expenses - Cloud', 'Concept': 'HTMLBusinessBreakdown',
           'SourceRawLabel': 'Cloud', 'Value': '100000000', 'Start': '2024-01-01',
           'End': '2024-03-31', 'SourceSemanticType': 'Other Costs and Expenses'}
    result = attach_income_face_evidence([row], document)[0]
    assert result['SourceEvidenceStatus'] == 'REJECTED'
    assert result['SourceClassificationConfidence'] == 0
    assert row['Concept'] == 'HTMLBusinessBreakdown'


@pytest.mark.parametrize('header', ['Three Months Ended March 31, | 2024', 'Three Months Ended | March 31, | 2024'])
def test_table_target_column_and_independent_relocation(tmp_path, monkeypatch, header):
    document = b'''<html><p>USD in millions</p><table><tr><td/><td colspan="2">Three Months Ended March 31,</td></tr>
    <tr><td/><td>2024</td><td>2025</td></tr><tr><td>TAC</td><td>100</td><td>120</td></tr></table></html>'''
    source = {'Concept': 'HTMLBusinessBreakdown', 'Label': 'Cost of Revenue - TAC', 'SourceRawLabel': 'TAC',
              'SourceValueHeader': header, 'SourceRawValue': '100',
              'SourceUnitScale': 1000000, 'Value': 100000000, 'Start': '2024-01-02', 'End': '2024-03-31'}
    result = attach_income_face_evidence([source], document)[0]
    assert result['Start'] == '2024-01-01'
    evidence = json.loads(result['SourceTableEvidence'])
    path = tmp_path / 'sec.gz'
    path.write_bytes(gzip.compress(document))
    monkeypatch.setattr('review_luna.run_review.xml_cache_for', lambda _: path)
    lineage = {'Table Evidence': result['SourceTableEvidence'], 'Source ID': 'F1',
               'Context': '2024-01-01..2024-03-31', 'Unit': 'USD', 'Value': 100000000}
    assert table_match(lineage, {'F1': {'Accession': 'A'}})[0] == 'PASS'
    # Refresh the artifact digest after swapping year columns: this must still
    # fail on the target column, not merely on a stale file hash.
    changed = document.replace(b'<td>2024</td><td>2025</td>', b'<td>2025</td><td>2024</td>')
    path.write_bytes(gzip.compress(changed))
    import hashlib
    evidence['artifact_sha256'] = hashlib.sha256(changed).hexdigest()
    lineage['Table Evidence'] = json.dumps(evidence)
    assert table_match(lineage, {'F1': {'Accession': 'A'}})[0] != 'PASS'


def test_derived_html_row_must_not_reuse_annual_table_as_quarter():
    row = {'Concept': 'HTMLBusinessBreakdown', 'SourceDerivation': 'annual_minus_ytd9',
           'Start': '2024-10-01', 'End': '2024-12-31', 'Value': 100,
           'SourceRawLabel': 'TAC', 'SourceRawValue': 400}
    result = attach_income_face_evidence([row], b'<html/>')[0]
    assert result['Start'] == row['Start']
    assert 'SourceTableEvidence' not in result


@pytest.mark.parametrize('start,end,fiscal_start,prior_end', [('2024-10-01', '2024-12-31', '2024-01-01', '2024-09-30'),
                                                         ('2024-04-01', '2024-06-30', '2023-07-01', '2024-03-31')])
def test_dimensional_derivation_keeps_exact_operand_dates_and_members(start, end, fiscal_start, prior_end):
    label = 'Costs and Expenses - Cloud'
    frame = pd.DataFrame({'2024-Q4': [30]}, index=pd.MultiIndex.from_tuples([('4a_Segments_Business', label)]))
    frame.attrs['fact_audit'] = pd.DataFrame([{'Category': '4a_Segments_Business', 'Label': label,
        'Period': '2024-Q4', 'Value': 30, 'Concept': 'CostsAndExpenses', 'DimCount': 1,
        'Start': start, 'End': end, 'SourceDerivationAnnualValue': 100,
        'SourceDerivationAnnualAccession': 'FY', 'SourceDerivationAnnualConcept': 'CostsAndExpenses',
        'SourceDerivationYTD9Value': 70, 'SourceDerivationYTD9Accession': '9M', 'SourceDerivationYTD9Concept': 'CostsAndExpenses'}])
    base = {'Concept': 'CostsAndExpenses', 'Label': label, 'Start': fiscal_start,
            'DimCount': 1, 'SourceEvidenceAxes': 'g:SegmentAxis', 'SourceEvidenceMembers': 'g:CloudMember'}
    inputs = [dict(base, End=end, Value=100, Accession='FY'), dict(base, End=prior_end, Value=70, Accession='9M')]
    retain_derivation_inputs(frame, inputs)
    retained = json.loads(frame.attrs['fact_audit'].iloc[0].SourceInputFacts)
    assert sum(Decimal(x['Value']) * Decimal(x['Coefficient']) for x in retained) == 30
    assert retained[0]['SourceEvidenceMembers'] == 'g:CloudMember'
    assert retained[1]['End'] == prior_end


def test_unreported_calculated_zero_is_withheld_not_promoted_to_reported():
    frame = pd.DataFrame()
    frame.attrs['fact_audit'] = pd.DataFrame([{'Category': '2_Balance_Sheet', 'Label': 'Short-term Borrowings',
        'Concept': 'CommercialPaper', 'SourceKind': 'xbrl', 'Value': 0, 'IsCalculated': True}])
    require_complete_output_evidence(frame)
    assert frame.attrs['fact_audit'].iloc[0].SourceClassificationConfidence == 0
    assert frame.attrs['fact_audit'].iloc[0].Value == 0  # Retained as an auditable candidate.


def test_unproven_zero_is_removed_from_front_but_reported_zero_survives():
    index = pd.MultiIndex.from_tuples([
        ('3_Cash_Flow', 'Unsupported Detail'),
        ('3_Cash_Flow', 'Reported Zero'),
    ])
    frame = pd.DataFrame({'2025-Q1': [0.0, 0.0]}, index=index)
    frame.attrs['fact_audit'] = pd.DataFrame([
        {'Category': '3_Cash_Flow', 'Label': 'Unsupported Detail',
         'Period': '2025-Q1', 'Value': 0.0, 'Concept': 'CalculatedDetail',
         'SourceKind': 'calculated', 'SourceDerivation': 'sum_components',
         'SourceInputFacts': json.dumps([{
             'Value': '0', 'Coefficient': '1', 'Accession': 'A',
             'Concept': 'MissingComponent', 'Start': '2025-01-01',
             'End': '2025-03-31', 'Unit': 'USD'}])},
        {'Category': '3_Cash_Flow', 'Label': 'Reported Zero',
         'Period': '2025-Q1', 'Value': 0.0, 'Concept': 'ReportedZero',
         'SourceKind': 'xbrl', 'SourceCellIdentity': '/html/body/ix:nonFraction[1]'},
    ])
    require_complete_output_evidence(frame)
    assert pd.isna(frame.loc[('3_Cash_Flow', 'Unsupported Detail'), '2025-Q1'])
    assert frame.loc[('3_Cash_Flow', 'Reported Zero'), '2025-Q1'] == 0
    rejected = frame.attrs['fact_audit'].iloc[0]
    assert rejected.SourceAdmissionRule == 'MISSING_RECURSIVE_OPERANDS'
    assert rejected.Value == 0  # The rejected candidate remains auditable.


def test_late_calculated_cell_without_exact_audit_is_withheld():
    key = ('3_Cash_Flow', 'Late Debt Repair')
    frame = pd.DataFrame({'2026-Q2': [-2_500_000.0]},
                         index=pd.MultiIndex.from_tuples([key]))
    frame.attrs['fact_audit'] = pd.DataFrame([{
        'Category': '3_Cash_Flow', 'Label': 'Different Detail',
        'Period': '2026-Q2', 'Value': -2_500_000.0,
        'Concept': 'NetCashProvidedByUsedInFinancingActivities',
        'SourceKind': 'xbrl', 'SourceCellIdentity': '/reported'}])
    require_complete_output_evidence(frame)
    assert pd.isna(frame.at[key, '2026-Q2'])


def test_grid_preserves_rowspan_physical_locator():
    from pipeline_audit import logical_table_grid
    grid = logical_table_grid(etree.fromstring(b'<table><tr><td rowspan="2">Label</td><td colspan="2">2024</td></tr><tr><td>1</td><td>2</td></tr></table>'))
    assert grid[1][0][1:] == (0, 0)
    assert grid[1][2][1:] == (1, 1)


def test_same_local_concept_in_different_namespaces_is_not_auto_linked():
    document = b'''<html xmlns:ix="http://www.xbrl.org/2013/inlineXBRL" xmlns:a="https://taxonomy/a" xmlns:b="https://taxonomy/b">
    <context id="c"><startDate>2024-01-01</startDate><endDate>2024-03-31</endDate></context><unit id="u"><measure>a:USD</measure></unit>
    <ix:nonFraction name="a:Amount" contextRef="c" unitRef="u">100</ix:nonFraction>
    <ix:nonFraction name="b:Amount" contextRef="c" unitRef="u">100</ix:nonFraction></html>'''
    row = {'Concept': 'Amount', 'Value': 100, 'Start': '2024-01-01', 'End': '2024-03-31', 'DimCount': 0}
    result = attach_income_face_evidence([row], document)[0]
    assert 'SourceCellIdentity' not in result


@pytest.mark.parametrize('reported_sum', [60, 61])
@pytest.mark.parametrize('existing_payload', [None, '[{"Value":999,"Accession":"STALE"}]'])
def test_annual_minus_three_quarters_retains_real_quarters_not_an_unlocated_sum(reported_sum, existing_payload):
    label = 'Cash Taxes Paid'
    frame = pd.DataFrame({'2024-Q1': [10], '2024-Q2': [20], '2024-Q3': [30], '2024-Q4': [40]},
                         index=pd.MultiIndex.from_tuples([('3_Cash_Flow', label)]))
    records = [{'Category': '3_Cash_Flow', 'Label': label, 'Concept': 'IncomeTaxesPaid', 'Period': f'2024-Q{i}',
                'Value': value, 'Start': start, 'End': end, 'Accession': f'Q{i}'}
               for i, value, start, end in [(1, 10, '2024-01-01', '2024-03-31'), (2, 20, '2024-04-01', '2024-06-30'),
                                           (3, 30, '2024-07-01', '2024-09-30')]]
    records.append({'Category': '3_Cash_Flow', 'Label': label, 'Concept': 'IncomeTaxesPaid', 'Period': '2024-Q4',
                    'Value': 40, 'Start': '2024-10-01', 'End': '2024-12-31', 'SourceDerivation': 'exact_concept_annual_minus_q1_q2_q3',
                    'SourceDerivationAnnualAccession': 'FY', 'SourceDerivationAnnualConcept': 'IncomeTaxesPaid',
                    'SourceDerivationAnnualValue': 100, 'SourceDerivationQ1Q2Q3SumValue': reported_sum,
                    'SourceInputFacts': existing_payload})
    frame.attrs['fact_audit'] = pd.DataFrame(records)
    retain_derivation_inputs(frame, [{'Concept': 'IncomeTaxesPaid', 'Accession': 'FY', 'Label': label,
                                      'Value': 100, 'Start': '2024-01-01', 'End': '2024-12-31'}])
    row = frame.attrs['fact_audit'].iloc[3]
    if reported_sum == 60:
        inputs = json.loads(row.SourceInputFacts)
        assert [i['Accession'] for i in inputs] == ['FY', 'Q1', 'Q2', 'Q3']
        assert sum(Decimal(str(i['Value'])) * Decimal(i['Coefficient']) for i in inputs) == 40
    else:
        assert pd.isna(row.get('SourceInputFacts'))


def test_existing_input_payload_does_not_allow_eps_quarter_subtraction():
    frame = pd.DataFrame()
    frame.attrs['fact_audit'] = pd.DataFrame([{'Category': '1_Income_Statement', 'Period': '2024-Q4',
        'Concept': 'EarningsPerShareDiluted', 'SourceDerivation': 'annual_minus_nine_months',
        'SourceInputFacts': '[{"Value":1}]', 'Value': 1}])
    require_complete_output_evidence(frame)
    assert frame.attrs['fact_audit'].iloc[0].SourceAdmissionRule == 'NON_ADDITIVE_QUARTER_RECONSTRUCTION'


def test_required_review_items_are_not_silently_truncated_at_thirty():
    from excel_lineage import _semantic_review_queue
    records = [{'_Category': '3_Cash_Flow', 'Metric': f'Unknown {i}', 'Period': '2024-Q1',
                'Value': i, 'Review Required': 'YES', 'Method': 'MISSING_RECURSIVE_OPERANDS',
                'Display Scale': 1, 'Display Unit': 'USD'} for i in range(35)]
    result = _semantic_review_queue(records, {'2024-Q1'}, {('3_Cash_Flow', r['Metric']) for r in records})
    assert len(result) == 35
    assert all('withheld' in r['Finding'] for r in result)


def test_wrapped_annual_table_uses_local_unit_paragraph(tmp_path, monkeypatch):
    document = b'<html><p>Cost of revenues (in millions, except percentages):</p><div><table><tr><td/><td>Year Ended December 31,</td></tr><tr><td/><td>2024</td></tr><tr><td>TAC</td><td>100</td></tr></table></div></html>'
    source = {'Concept': 'HTMLBusinessBreakdown', 'Label': 'Cost of Revenue - TAC', 'SourceRawLabel': 'TAC',
              'SourceValueHeader': 'Year Ended December 31, | 2024', 'SourceRawValue': '100',
              'SourceUnitScale': 1000000, 'Value': 100000000, 'End': '2024-12-31'}
    result = attach_income_face_evidence([source], document)[0]
    path = tmp_path / 'annual.gz'
    path.write_bytes(gzip.compress(document))
    monkeypatch.setattr('review_luna.run_review.xml_cache_for', lambda _: path)
    lineage = {'Table Evidence': result['SourceTableEvidence'], 'Source ID': 'F1',
               'Context': '2024-01-01..2024-12-31', 'Unit': 'USD', 'Value': 100000000}
    assert table_match(lineage, {'F1': {'Accession': 'FY'}})[0] == 'PASS'
