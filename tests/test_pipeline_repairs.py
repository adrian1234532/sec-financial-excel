from unittest.mock import patch

import pandas as pd

import sec_data
from excel_lineage import _display_transform, _origin
from pipeline_audit import enrich_output_audit, prefer_cashflow_face_candidates


def test_console_diagnostics_replace_unencodable_text_instead_of_aborting():
    class ConsoleStream:
        def __init__(self):
            self.errors = "strict"

        def reconfigure(self, **options):
            self.errors = options["errors"]

    stdout = ConsoleStream()
    stderr = ConsoleStream()
    with patch.object(sec_data.sys, "stdout", stdout), patch.object(
        sec_data.sys, "stderr", stderr
    ):
        sec_data._configure_console_text_errors()
    assert stdout.errors == "replace"
    assert stderr.errors == "replace"


def test_cash_flow_total_investment_gain_has_its_own_mapping():
    tags = sec_data.CONCEPT_MAP['Gain/Loss on Investments (CF)']['tags']
    assert tags[0] == 'DebtAndEquitySecuritiesGainLoss'


def test_geography_regions_and_country_are_one_footing_basis():
    frame = pd.DataFrame(
        {'2026-Q2': [100, 60, 40]},
        index=pd.MultiIndex.from_tuples([
            ('1_Income_Statement', 'Revenue'),
            ('4b_Segments_Geographic_Regions', 'Revenue - International'),
            ('4c_Segments_Geographic_Countries', 'Revenue - United States'),
        ]),
    )
    with patch.object(sec_data, '_rescue_segments_from_html') as rescue:
        sec_data._audit_segment_footing(frame, 'TEST', 12)
    rescue.assert_not_called()


def test_cash_flow_display_sign_does_not_change_stored_value():
    scale, unit = _display_transform('Gain/Loss on Investments (CF)', 'USD',
                                     {'SourceDisplaySign': -1})
    assert scale == -1e-6
    assert unit == '$mm'


def test_engine_arithmetic_is_calculated_not_reported_quarter_derivation():
    assert _origin('1_Income_Statement', {
        'SourceKind': 'calculated', 'SourceDerivationFormula': 'Revenue - Cost of Revenue',
    }) == 'C'


def test_cashflow_face_total_blocks_direct_quarter_footnote_component():
    facts = pd.DataFrame([
        {'Category': '3_Cash_Flow', 'Label': 'Investment adjustment', 'FY': 2026,
         'Concept': 'FaceTotal', 'Value': 135803, 'Duration': 180},
        {'Category': '3_Cash_Flow', 'Label': 'Investment adjustment', 'FY': 2026,
         'Concept': 'FootnoteEquityOnly', 'Value': 99031, 'Duration': 90},
        {'Category': '1_Income_Statement', 'Label': 'Investment gain', 'FY': 2026,
         'Concept': 'FootnoteEquityOnly', 'Value': 99031, 'Duration': 90},
    ])
    selected = prefer_cashflow_face_candidates(facts, {'FaceTotal': {'3_Cash_Flow'}})
    assert selected.Value.tolist() == [135803, 99031]


def test_gross_profit_records_exact_operands_and_rejects_different_dates():
    keys = [('1_Income_Statement', label) for label in ('Revenue', 'Cost of Revenue', 'Gross Profit')]
    frame = pd.DataFrame({'2026-Q2': [120, 46, 74]}, index=pd.MultiIndex.from_tuples(keys))
    frame.attrs['fact_audit'] = pd.DataFrame([
        {'Category': category, 'Label': label, 'Period': '2026-Q2', 'Value': value,
         'Concept': concept, 'Accession': '0001652044-26-000071',
         'Start': '2026-04-01', 'End': '2026-06-30'}
        for (category, label), value, concept in zip(keys[:2], [120, 46], ['Revenues', 'CostOfRevenue'])
    ])
    bad = frame.copy()
    bad.attrs['fact_audit'] = frame.attrs['fact_audit'].copy()
    bad.attrs['fact_audit'].loc[1, 'Start'] = '2026-01-01'
    assert len(enrich_output_audit(bad, {}, set()).attrs['fact_audit']) == 2
    result = enrich_output_audit(frame, {}, set()).attrs['fact_audit']
    output = result.iloc[-1]
    assert output.Label == 'Gross Profit'
    assert output.SourceKind == 'calculated'
    assert 'CostOfRevenue' in output.SourceInputFacts


def test_missing_eps_and_weighted_shares_are_not_guessed_from_adjacent_quarter():
    labels = ['Revenue', 'Cost of Revenue', 'Net Income', 'EPS Basic', 'Shares Outstanding Basic']
    frame = pd.DataFrame({'2026-Q2': [300, 80, 200, None, None],
                          '2026-Q1': [200, 60, 100, 2, 50]},
                         index=pd.MultiIndex.from_tuples([('1_Income_Statement', label) for label in labels]))
    output = sec_data.calculate_kpis(frame)
    guessed = output[(output.Period == '2026-Q2') & output.Label.isin(
        ['EPS Basic', 'EPS Diluted', 'Shares Outstanding Basic', 'Shares Outstanding Diluted'])]
    assert guessed.empty


def test_annual_weighted_average_cannot_be_trusted_as_q4():
    frame = pd.DataFrame({'2025-Q4': [120]}, index=pd.MultiIndex.from_tuples([
        ('1_Income_Statement', 'Shares Outstanding Basic')]))
    frame.attrs['fact_audit'] = pd.DataFrame([{
        'Category': '1_Income_Statement', 'Label': 'Shares Outstanding Basic', 'Period': '2025-Q4',
        'Value': 120, 'Concept': 'WeightedAverageNumberOfSharesOutstandingBasic',
        'Start': '2025-01-01', 'End': '2025-12-31', 'SourceClassificationConfidence': 1,
    }])
    audit = enrich_output_audit(frame, {}, set()).attrs['fact_audit']
    assert audit.iloc[0].SourceClassificationConfidence == 0
    assert audit.iloc[0].SourceAdmissionRule == 'REJECTED_NON_ADDITIVE_PERIOD_MISMATCH'
