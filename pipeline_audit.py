"""Attach evidence to finished engine values without selecting or repairing them."""

import hashlib
import json
import re
from collections import defaultdict
from collections.abc import Mapping
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, localcontext
from typing import Any

import pandas as pd
from lxml import etree


def _normalized(value: Any) -> str:
    return re.sub(r'[^a-z0-9]', '', str(value).lower())


def _lexical_decimal(node: Any) -> Decimal:
    lexical = ''.join(node.itertext()).strip().replace(',', '').replace('\xa0', '')
    lexical = lexical.replace('(', '-').replace(')', '')
    value = Decimal(lexical)
    with localcontext() as ctx:
        ctx.prec = max(50, len(value.as_tuple().digits) + abs(int(node.get('scale', '0'))) + 5)
        value = value.scaleb(int(node.get('scale', '0')))
    if node.get('sign') == '-':
        value = -value
    if not value.is_finite():
        raise InvalidOperation
    return value


def logical_table_grid(table: Any) -> list[list[tuple[Any, int, int]]]:
    """Logical columns retain their physical cell and row/column positions."""
    grid: list[list[tuple[Any, int, int]]] = []
    occupied: dict[tuple[int, int], tuple[Any, int, int]] = {}
    rows = table.xpath('./*[local-name()="tr"] | ./*[local-name()="tbody" or local-name()="thead"]/*[local-name()="tr"]')
    for ri, row in enumerate(rows):
        ci = 0
        for physical, cell in enumerate(row.xpath('./*[local-name()="td" or local-name()="th"]')):
            while (ri, ci) in occupied:
                ci += 1
            width, height = int(cell.get('colspan', '1')), int(cell.get('rowspan', '1'))
            for dy in range(height):
                for dx in range(width):
                    occupied[ri + dy, ci + dx] = cell, ri, physical
            ci += width
        last = max((column for rr, column in occupied if rr == ri), default=-1)
        grid.append([occupied.get((ri, col), (None, ri, -1)) for col in range(last + 1)])
    return grid


def table_cell_evidence(root: Any, row: Mapping[str, Any], prepared: Any = None) -> dict[str, Any] | None:
    """Locate a frozen HTML row and period column; never choose the first number."""
    label = _normalized(row.get('SourceRawLabel') or row.get('SourceLabel'))
    header = str(row.get('SourceValueHeader') or '')
    years = re.findall(r'\b(?:19|20)\d{2}\b', header)
    if not label or not years or not row.get('SourceRawValue'):
        return None
    year = years[-1]
    try:
        expected = Decimal(str(row['SourceRawValue']).replace(',', '').replace('(', '-').replace(')', ''))
        if not expected.is_finite():
            return None
    except InvalidOperation:
        return None
    matches = []
    tables = prepared if prepared is not None else [(table, logical_table_grid(table)) for table in root.xpath('//*[local-name()="table"]')]
    for table, grid in tables:
        for ri, cells in enumerate(grid):
            if not cells or cells[0][0] is None or _normalized(''.join(cells[0][0].itertext())) != label:
                continue
            for column, (cell, local_row, physical_col) in enumerate(cells):
                if cell is None or physical_col < 1:
                    continue
                preceding = [' '.join(x[0].itertext()).strip() for rr in grid[:ri] if column < len(rr) and (x := rr[column])[0] is not None]
                if year not in preceding:
                    continue
                period_header = ' '.join(preceding)
                heading = ' '.join(header.split('|')[:-1]).strip()
                # Combined headers must identify the exact duration column.
                if _normalized(heading) not in _normalized(period_header):
                    continue
                try:
                    value = _lexical_decimal(cell)
                except (InvalidOperation, ValueError):
                    continue
                if value != expected:
                    continue
                months = {'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
                          'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12}
                end_match = re.search(r'(' + '|'.join(months) + r')\s+(\d{1,2})', heading.lower())
                if end_match is None:
                    continue
                end = date(int(year), months[end_match[1]], int(end_match[2]))
                width = 12 if 'year ended' in heading.lower() else next((n for word, n in [('three', 3), ('six', 6), ('nine', 9)] if word + ' months' in heading.lower()), None)
                if width is None:
                    continue
                start_month = end.month - width + 1
                if start_month < 1:
                    continue  # Unestablished fiscal calendars require explicit evidence.
                start = date(end.year, start_month, 1)
                scale = Decimal(str(row.get('SourceUnitScale', '1')))
                unit_caption = ' '.join(table.itertext()) + ' '.join(' '.join(node.itertext()) for node in table.xpath('preceding-sibling::*[position() <= 4]'))
                # SEC wraps some tables in an otherwise empty div. Its local
                # introductory paragraph carries the unit, outside that wrapper.
                parent = table.getparent()
                if parent is not None and len(parent.xpath('./*[local-name()="table"]')) == 1:
                    unit_caption += ' '.join(' '.join(node.itertext()) for node in parent.xpath('preceding-sibling::*[position() <= 5]'))
                if scale != Decimal('1000000') or 'in millions' not in unit_caption.lower():
                    continue
                matches.append({'kind': 'TABLE', 'xpath': root.getroottree().getpath(cell),
                                'table_xpath': root.getroottree().getpath(table), 'local_row': local_row,
                                'physical_column': physical_col, 'logical_column': column, 'label': str(row.get('SourceRawLabel')),
                                'header': heading, 'year': year, 'raw': str(value), 'scale': str(scale),
                                'start': start.isoformat(), 'end': end.isoformat(), 'unit': 'USD',
                                'unit_caption': unit_caption[:1000],
                                'artifact_sha256': hashlib.sha256(etree.tostring(root)).hexdigest()})
    signatures = {(m['start'], m['end'], m['raw'], m['label']) for m in matches}
    return matches[0] if len(signatures) == 1 else None


def attach_income_face_evidence(facts: list[dict[str, Any]], document: bytes | str) -> list[dict[str, Any]]:
    """Locate revenue in the filed income table, rather than infer it from tag rank.

    Only tables containing operating income and EPS qualify. Evidence is matched
    by concept, exact dates and lexical value; a supplemental revenue is retained
    as a disclosure when that same filing reports a different face total.
    """
    root = etree.fromstring(document.encode() if isinstance(document, str) else document,
                            parser=etree.XMLParser(resolve_entities=False, no_network=True))
    contexts = {}
    context_nodes = {}
    for context in root.xpath('//*[local-name()="context"]'):
        context_nodes[context.get('id')] = context
        contexts[context.get('id')] = (
            str(context.xpath('string(.//*[local-name()="startDate"])')) or None,
            str(context.xpath('string(.//*[local-name()="endDate"])') or context.xpath('string(.//*[local-name()="instant"])')),
            bool(context.xpath('.//*[local-name()="explicitMember" or local-name()="typedMember"]')),
        )
    evidence: dict[tuple[Any, ...], dict[str, Any]] = {}
    face_periods = set()
    for table in root.xpath('//*[local-name()="table"]'):
        nodes = table.xpath('.//*[@contextRef and @unitRef and @name]')
        concepts = {str(node.get('name')).split(':')[-1] for node in nodes}
        if 'OperatingIncomeLoss' not in concepts or not any(name.startswith('EarningsPerShare') for name in concepts):
            continue
        for node in nodes:
            concept = str(node.get('name')).split(':')[-1]
            if not any(row.get('Label') == 'Revenue' and row.get('Concept') == concept for row in facts):
                continue
            start, end, dimensional = contexts.get(node.get('contextRef'), (None, '', True))
            if dimensional or not start:
                continue
            lexical = ''.join(node.itertext()).strip()
            try:
                value = Decimal(lexical.replace(',', '')).scaleb(int(node.get('scale', '0')))
                if node.get('sign') == '-':
                    value = -value
                if not value.is_finite():
                    continue
            except (InvalidOperation, ValueError):
                continue
            key: tuple[Any, ...] = concept, start, end, value
            evidence[key] = {'SourceCellIdentity': root.getroottree().getpath(node),
                             'SourceRawValue': lexical, 'SourceFaceRevenue': True,
                             'SourceContextID': node.get('contextRef')}
            face_periods.add((start, end))
    result = []
    raw_index: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    prepared_tables = [(table, logical_table_grid(table)) for table in root.xpath('//*[local-name()="table"]')]
    units = {node.get('id'): '/'.join(str(measure.text).split(':')[-1] for measure in node.xpath('.//*[local-name()="measure"]'))
             for node in root.xpath('//*[local-name()="unit"]')}
    for node in root.xpath('//*[@contextRef and @unitRef and @name]'):
        start, end, dimensional = contexts.get(node.get('contextRef'), (None, '', True))
        try:
            amount = _lexical_decimal(node)
        except (InvalidOperation, ValueError):
            continue
        name = str(node.get('name'))
        concept = name.split(':')[-1]
        namespace = node.nsmap.get(name.split(':')[0], '')
        context_node = context_nodes.get(node.get('contextRef'))
        if context_node is None:
            continue
        members = context_node.xpath('.//*[local-name()="explicitMember"]')
        tr = node.xpath('ancestor::*[local-name()="tr"][1]')
        first = tr[0].xpath('./*[local-name()="td" or local-name()="th"][1]') if tr else []
        physical_label = _normalized(''.join(first[0].itertext())) if first else ''
        raw_index.setdefault((concept, start, end, amount), []).append({
            'SourceContextID': node.get('contextRef'), 'SourceRawValue': ''.join(node.itertext()).strip(),
            'SourceCellIdentity': root.getroottree().getpath(node),
            'SourceExactUnitKind': units.get(node.get('unitRef')),
            'SourceEvidenceAxes': ' | '.join(str(member.get('dimension')) for member in members),
            'SourceEvidenceMembers': ' | '.join(str(member.text) for member in members),
            'SourceConceptQName': '{' + namespace + '}' + concept, '_dimensional': dimensional,
            '_row_label': physical_label})
    other_facts: dict[tuple[Any, ...], set[Decimal]] = {}
    for node in root.xpath('//*[@contextRef and @unitRef and @name]'):
        if str(node.get('name')).split(':')[-1] != 'PaymentsForProceedsFromOtherInvestingActivities':
            continue
        start, end, dimensional = contexts.get(node.get('contextRef'), (None, '', True))
        if dimensional:
            continue
        try:
            amount = Decimal(''.join(node.itertext()).strip().replace(',', '')).scaleb(int(node.get('scale', '0')))
            if node.get('sign') == '-':
                amount = -amount
            other_facts.setdefault((start, end), set()).add(amount)
        except (InvalidOperation, ValueError):
            continue
    for original in facts:
        row = dict(original)
        if row.get('Concept') == 'HTMLBusinessBreakdown' and str(row.get('SourceDerivation', '')).strip() in {'', 'nan', 'None'}:
            row['SourceOriginalConcept'] = row['Concept']
            # A table-local numeric tag outranks a prefix inherited from nearby
            # prose. In particular, a Revenue tag is not an expense observation.
            raw_label = _normalized(row.get('SourceRawLabel') or row.get('SourceLabel'))
            tagged = []
            for key, appearances in raw_index.items():
                if key[2] != row.get('End') or key[3] != Decimal(str(row.get('Value'))):
                    continue
                for appearance in appearances:
                    if appearance['_row_label'] == raw_label:
                        tagged.append((key, appearance))
            html_signatures = {(key[0], key[1], key[2], item['SourceEvidenceAxes'], item['SourceEvidenceMembers'], item['SourceConceptQName']) for key, item in tagged}
            if len(html_signatures) == 1:
                key, appearance = tagged[0]
                row.update({name: value for name, value in appearance.items() if not name.startswith('_')})
                row.update(Concept=key[0], Start=key[1], End=key[2], DimCount=int(appearance['_dimensional']))
                if key[0].startswith(('Revenue', 'SalesRevenue')) and 'expense' in str(row.get('SourceSemanticType', '')).lower():
                    row.update(SourceClassificationConfidence=0.0, SourceAdmissionRule='REJECTED_REVENUE_AS_EXPENSE',
                               SourceSemanticType='Unclassified: revenue tagged as expense', SourceEvidenceStatus='REJECTED')
                if key[0] == 'OperatingIncomeLoss' and 'expense' in str(row.get('SourceSemanticType', '')).lower():
                    row.update(SourceClassificationConfidence=0.0, SourceAdmissionRule='REJECTED_OPERATING_INCOME_AS_EXPENSE',
                               SourceSemanticType='Unclassified: operating income tagged as expense', SourceEvidenceStatus='REJECTED')
            else:
                table_evidence = table_cell_evidence(root, row, prepared_tables)
                if table_evidence:
                    table_evidence['artifact_sha256'] = hashlib.sha256(document.encode() if isinstance(document, str) else document).hexdigest()
                    row.update(Start=table_evidence['start'], End=table_evidence['end'],
                               SourceCellIdentity=table_evidence['xpath'], SourceTableEvidence=json.dumps(table_evidence),
                               SourceExactUnitKind='USD', SourceAdmissionRule='html_table_period_column',
                               SourceClassificationConfidence=row.get('SourceConfidence', 0.95))
        try:
            raw_key = row.get('Concept'), row.get('Start'), row.get('End'), Decimal(str(row.get('Value')))
            occurrences = [item for item in raw_index.get(raw_key, []) if item['_dimensional'] == bool(row.get('DimCount'))]
            signatures = {(item['SourceExactUnitKind'], item['SourceEvidenceAxes'], item['SourceEvidenceMembers'], item['SourceConceptQName']) for item in occurrences}
            if len(signatures) == 1:
                row.update({name: value for name, value in occurrences[0].items() if not name.startswith('_')})
        except InvalidOperation:
            pass
        if row.get('Concept') == 'PaymentsForProceedsFromOtherInvestingActivities' and not row.get('DimCount'):
            amounts = other_facts.get((row.get('Start'), row.get('End')), set())
            if len(amounts) == 1:
                row['SourceReportedValue'] = str(next(iter(amounts)))
        if row.get('Category') == '1_Income_Statement' and row.get('Label') == 'Revenue' and not row.get('DimCount'):
            try:
                key = row.get('Concept'), row.get('Start'), row.get('End'), Decimal(str(row.get('Value')))
            except InvalidOperation:
                key = ()
            if key in evidence:
                row.update(evidence[key])
            elif (row.get('Start'), row.get('End')) in face_periods:
                row.update(Category='6_Disclosures', Label='Revenue excluding face-statement adjustments',
                           SourceAdmissionRule='SUPPLEMENTAL_REVENUE_NOT_FACE_TOTAL')
        result.append(row)
    return result


def prefer_cashflow_face_candidates(
    facts: pd.DataFrame, presented: Mapping[str, Any],
) -> pd.DataFrame:
    """A footnote component cannot substitute for the cash-flow face total."""
    face_concepts = {concept for concept, categories in presented.items()
                     if '3_Cash_Flow' in categories}
    cash = facts[facts.Category.eq('3_Cash_Flow')]
    proven = cash[cash.Concept.isin(face_concepts)][['Label', 'FY']].drop_duplicates()
    if proven.empty:
        return facts
    proven_keys = set(map(tuple, proven.itertuples(index=False, name=None)))
    reject = facts.Category.eq('3_Cash_Flow') & ~facts.Concept.isin(face_concepts)
    reject &= pd.Series([(label, year) in proven_keys for label, year in
                         zip(facts.Label, facts.FY)], index=facts.index)
    return facts.loc[~reject].copy()


def retain_derivation_inputs(frame: pd.DataFrame, source_facts: list[dict[str, Any]]) -> pd.DataFrame:
    """Persist the actual operands retained by the engine, with source dates."""
    audit = frame.attrs.get('fact_audit')
    if not isinstance(audit, pd.DataFrame):
        return frame
    audit = audit.copy()
    audit['SourceInputFacts'] = audit['SourceInputFacts'].astype(object) if 'SourceInputFacts' in audit else pd.Series(None, index=audit.index, dtype=object)
    source_index: dict[tuple[Any, Any], list[dict[str, Any]]] = defaultdict(list)
    for fact in source_facts:
        source_index[fact.get('Accession'), fact.get('Concept')].append(fact)
        if fact.get('SourceOriginalConcept') and fact.get('SourceOriginalConcept') != fact.get('Concept'):
            source_index[fact.get('Accession'), fact.get('SourceOriginalConcept')].append(fact)
    for index, row in audit.iterrows():
        if str(row.get('Concept', '')).startswith(('EarningsPerShare', 'WeightedAverage')):
            continue
        if pd.notna(row.get('SourceDerivationQ1Q2Q3SumValue')):
            audit.at[index, 'SourceInputFacts'] = None  # Rebuild, never reuse stale operands.
            year = str(row.get('Period')).split('-')[0]
            parts = []
            for quarter in (1, 2, 3):
                matches = audit[audit.Category.eq(row.get('Category')) & audit.Label.eq(row.get('Label'))
                                & audit.Period.eq(f'{year}-Q{quarter}')]
                if len(matches) != 1:
                    break
                part = matches.iloc[0].to_dict()
                if part.get('SourceEvidenceStatus') == 'REJECTED':
                    break
                parts.append(part)
            if len(parts) != 3 or sum(Decimal(str(part['Value'])) for part in parts) != Decimal(str(row['SourceDerivationQ1Q2Q3SumValue'])):
                continue
            annuals = [fact for fact in source_index.get((row.get('SourceDerivationAnnualAccession'), row.get('SourceDerivationAnnualConcept')), [])
                       if fact.get('End') == row.get('End') and fact.get('Label') == row.get('Label')
                       and Decimal(str(fact.get('Value'))) == Decimal(str(row.get('SourceDerivationAnnualValue')))
                       and fact.get('SourceEvidenceStatus') != 'REJECTED']
            if len({(fact.get('Start'), fact.get('End')) for fact in annuals}) != 1:
                continue
            inputs = []
            for part, coefficient in [(annuals[0], 1), *((part, -1) for part in parts)]:
                inputs.append({field: part.get(field) for field in ('Value', 'Accession', 'Concept', 'Start', 'End', 'Label',
                    'SourceReportedValue', 'SourceRawValue', 'SourceCellIdentity',
                    'SourceEvidenceAxes', 'SourceEvidenceMembers', 'SourceTableEvidence', 'SourceInputFacts')})
                inputs[-1].update(Coefficient=str(coefficient), Unit=part.get('SourceExactUnitKind') or 'USD')
            audit.at[index, 'SourceInputFacts'] = json.dumps(inputs, default=str)
            audit.at[index, 'SourceClassificationConfidence'] = row.get('SourceConfidence') if pd.notna(row.get('SourceConfidence')) else 0.95
            audit.at[index, 'SourceAdmissionRule'] = 'retained_annual_minus_three_reported_quarters'
            continue
        if isinstance(row.get('SourceInputFacts'), str) and row['SourceInputFacts'].strip():
            continue
        if pd.notna(row.get('SourceDerivationAnnualValue')):
            operands = [('Annual', 1), ('YTD9', -1)]
        elif pd.notna(row.get('SourceDerivationYTDValue')):
            operands = [('YTD', 1), ('Quarter' if pd.notna(row.get('SourceDerivationQuarterValue')) else 'Baseline', -1)]
        else:
            continue
        inputs = []
        target_start, target_end = str(row.get('Start')), str(row.get('End'))
        if target_start in {'None', 'nan'} or target_end in {'None', 'nan'}:
            continue
        fiscal_start = None
        before_target = (date.fromisoformat(target_start) - timedelta(days=1)).isoformat()
        inverse_q1 = str(row.get('SourceDerivation')) == 'exact_concept_ytd6_minus_q2'
        first_name = operands[0][0]
        first_prefix = 'SourceDerivation' + first_name
        first_candidates = [fact for fact in source_index.get((row.get(first_prefix + 'Accession'), row.get(first_prefix + 'Concept')), [])
                            if (pd.to_numeric(fact.get('Value'), errors='coerce') == row.get(first_prefix + 'Value')
                                or (row.get(first_prefix + 'Concept') == 'PaymentsForProceedsFromOtherInvestingActivities'
                                    and abs(pd.to_numeric(fact.get('Value'), errors='coerce')) == abs(row.get(first_prefix + 'Value'))))
                            and bool(fact.get('DimCount')) == bool(row.get('DimCount'))
                            and (not row.get('DimCount') or fact.get('Label') == row.get('Label'))]
        operand_end = target_end
        if inverse_q1:
            next_start = (date.fromisoformat(target_end) + timedelta(days=1)).isoformat()
            quarter_prefix = 'SourceDerivationQuarter'
            ends = {fact.get('End') for fact in source_index.get((row.get(quarter_prefix + 'Accession'), row.get(quarter_prefix + 'Concept')), [])
                    if fact.get('Start') == next_start and pd.to_numeric(fact.get('Value'), errors='coerce') == row.get(quarter_prefix + 'Value')}
            if len(ends) != 1:
                continue
            operand_end = str(next(iter(ends)))
            fiscal_start = target_start
        else:
            starts = {fact.get('Start') for fact in first_candidates if fact.get('End') == operand_end
                      and fact.get('Start') and str(fact['Start']) < target_start}
            if len(starts) != 1:
                continue
            fiscal_start = str(next(iter(starts)))
        for name, coefficient in operands:
            prefix = 'SourceDerivation' + name
            amount = row.get(prefix + 'Value')
            accession, concept = row.get(prefix + 'Accession'), row.get(prefix + 'Concept')
            signed_other = concept == 'PaymentsForProceedsFromOtherInvestingActivities'
            expected_dates = {
                'Annual': (fiscal_start, target_end), 'YTD9': (fiscal_start, before_target),
                'YTD': (fiscal_start, operand_end),
                'Quarter': (next_start, operand_end) if inverse_q1 else (fiscal_start, before_target),
                'Baseline': (fiscal_start, before_target)}[name]
            candidates = [fact for fact in source_index.get((accession, concept), []) if fact.get('Accession') == accession
                          and (fact.get('Concept') == concept or fact.get('SourceOriginalConcept') == concept)
                          and (fact.get('Start'), fact.get('End')) == expected_dates
                          and bool(fact.get('DimCount')) == bool(row.get('DimCount'))
                          and (not row.get('DimCount') or fact.get('Label') == row.get('Label'))
                          and (pd.to_numeric(fact.get('Value'), errors='coerce') == amount
                               or (signed_other and abs(pd.to_numeric(fact.get('Value'), errors='coerce')) == abs(amount)))]
            aspects = {(fact.get('Start'), fact.get('End')) for fact in candidates}
            if len(aspects) != 1:
                break
            if any(fact.get('SourceEvidenceStatus') == 'REJECTED' for fact in candidates):
                break
            start, end = next(iter(aspects))
            if signed_other:
                raw_values = {str(fact.get('SourceReportedValue')) for fact in candidates if pd.notna(fact.get('SourceReportedValue'))}
                if len(raw_values) != 1:
                    break
                amount = next(iter(raw_values))
            inputs.append({'Value': str(amount), 'Coefficient': str(coefficient),
                           'Accession': accession, 'Concept': candidates[0].get('SourceConceptQName') or candidates[0].get('Concept'),
                           'Start': start, 'End': end, 'Label': row.get('Label'),
                           'Unit': candidates[0].get('SourceExactUnitKind') or 'USD',
                           'SourceReportedValue': candidates[0].get('SourceReportedValue'),
                           'SourceRawValue': candidates[0].get('SourceRawValue'),
                           'SourceCellIdentity': candidates[0].get('SourceCellIdentity'),
                           'SourceEvidenceAxes': candidates[0].get('SourceEvidenceAxes') or candidates[0].get('SourceDimensionAxes'),
                           'SourceEvidenceMembers': candidates[0].get('SourceEvidenceMembers') or candidates[0].get('SourceDimensionMembers'),
                           'SourceTableEvidence': candidates[0].get('SourceTableEvidence')})
        if len(inputs) == 2:
            audit.at[index, 'SourceInputFacts'] = json.dumps(inputs)
            if row.get('SourceAdmissionRule') == 'MISSING_RECURSIVE_OPERANDS':
                audit.at[index, 'SourceClassificationConfidence'] = row.get('SourceConfidence') if pd.notna(row.get('SourceConfidence')) else 0.95
                audit.at[index, 'SourceAdmissionRule'] = 'explicit_retained_filing_operands'
            if row.get('Concept') == 'PaymentsForProceedsFromOtherInvestingActivities':
                total = sum(Decimal(item['Value']) * Decimal(item['Coefficient']) for item in inputs)
                frame.at[(row['Category'], row['Label']), row['Period']] = float(total)
                audit.at[index, 'Value'] = float(total)
                audit.at[index, 'SourceDerivationFormula'] = f"{inputs[0]['Value']} - {inputs[1]['Value']} = {total}"
    frame.attrs['fact_audit'] = audit
    return frame


def enrich_output_audit(
    frame: pd.DataFrame, calculation_parents: Mapping[str, Any], cash_roots: set[str],
) -> pd.DataFrame:
    audit = frame.attrs.get('fact_audit')
    if not isinstance(audit, pd.DataFrame) or audit.empty:
        return frame
    matrix = pd.DataFrame(frame.to_numpy(copy=False), index=frame.index, columns=frame.columns)
    audit = audit.copy()
    # Annual/YTD weighted averages sometimes leak into the legacy Q4 slot.
    # This is rejection evidence, not a quarter-length classification heuristic.
    concepts = audit.Concept.fillna('').astype(str)
    nonadditive = concepts.str.startswith(('WeightedAverage', 'EarningsPerShare'))
    quarters = audit.Period.fillna('').astype(str).str.contains(r'-Q[1-4]$', regex=True)
    days = (pd.to_datetime(audit.End, errors='coerce')
            - pd.to_datetime(audit.Start, errors='coerce')).dt.days
    mismatched = nonadditive & quarters & days.gt(150)
    audit.loc[mismatched, 'SourceClassificationConfidence'] = 0.0
    audit.loc[mismatched, 'SourceSemanticType'] = 'Non-additive period mismatch'
    audit.loc[mismatched, 'SourceAdmissionRule'] = 'REJECTED_NON_ADDITIVE_PERIOD_MISMATCH'
    # XBRL values retain their lexical direction. The filed calculation weight
    # records how a face line is displayed as a cash-flow contribution.
    for index, row in audit.iterrows():
        if row.get('Category') != '3_Cash_Flow':
            continue
        weights = {
            float(weight) for parent, weight in calculation_parents.get(str(row.get('Concept')), ())
            if parent in cash_roots and float(weight) in {-1.0, 1.0}
        }
        if len(weights) == 1:
            weight = weights.pop()
            audit.at[index, 'SourceDisplaySign'] = weight
            # The legacy bridge already applies the calculation weight to this
            # signed net line. Restore the retained SEC value at the export
            # boundary so its display weight is applied exactly once.
            raw = pd.to_numeric(row.get('SourceReportedValue'), errors='coerce')
            if row.get('Concept') == 'PaymentsForProceedsFromOtherInvestingActivities' and pd.notna(raw):
                key = row.get('Category'), row.get('Label')
                period = row.get('Period')
                if key in frame.index and period in frame.columns:
                    current = pd.to_numeric(frame.at[key, period], errors='coerce')
                    if current in (raw, raw * weight):
                        frame.at[key, period] = raw
                        audit.at[index, 'Value'] = raw

    income, cash = '1_Income_Statement', '3_Cash_Flow'
    recipes = [
        (income, 'Gross Profit', [(income, 'Revenue', 1), (income, 'Cost of Revenue', -1)]),
        (income, 'Total Operating Expenses', [(income, 'Gross Profit', 1), (income, 'Operating Income', -1)]),
        (cash, 'Net Cash Flow', [(cash, 'Operating Cash Flow', 1),
                               (cash, 'Investing Cash Flow', 1), (cash, 'Financing Cash Flow', 1)]),
        (cash, 'Purchases of Investments', [(cash, 'Purchases of Marketable Securities', 1),
                                           (cash, 'Purchases of Non-Marketable / Other Investments', 1)]),
        (cash, 'Proceeds from Investments', [(cash, 'Proceeds from Marketable Securities', 1),
                                           (cash, 'Proceeds from Non-Marketable / Other Investments', 1)]),
        (cash, 'Total Net Debt Issued (Repaid)', [(cash, 'Total Debt Issued', 1), (cash, 'Total Debt Repaid', -1)]),
    ]
    for output_category, output_label, recipe in recipes:
        output_key = output_category, output_label
        if output_key not in frame.index or not all((cat, label) in frame.index for cat, label, _ in recipe):
            continue
        additions = []
        for period in frame.columns:
            value = matrix.loc[output_key, period]
            if pd.isna(value):
                continue
            existing = audit[audit.Category.eq(output_key[0]) & audit.Label.eq(output_key[1])
                             & audit.Period.eq(period)]
            if not existing.empty and not str(existing.iloc[0].get('SourceKind', '')).startswith('calculated') and not str(existing.iloc[0].get('Concept', '')).startswith('Calculated'):
                continue
            inputs = []
            for category, label, coefficient in recipe:
                source_value = matrix.loc[(category, label), period]
                candidates = audit[audit.Category.eq(category) & audit.Label.eq(label)
                                   & audit.Period.eq(period)]
                matches = candidates[pd.to_numeric(candidates.Value, errors='coerce').eq(source_value)]
                if matches.empty or pd.isna(source_value):
                    break
                source = matches.iloc[0].to_dict()
                if not source.get('Accession') or not (source.get('Concept') or source.get('SourceInputFacts')):
                    break
                source['Coefficient'] = coefficient
                inputs.append(source)
            if len(inputs) != len(recipe) or float(value) != sum(
                    float(item['Value']) * item['Coefficient'] for item in inputs):
                continue
            if any(str(inputs[0].get(field, '')) != str(item.get(field, ''))
                   for item in inputs[1:]
                   for field in ('Start', 'End', 'SourceDimensionAxes', 'SourceDimensionMembers')):
                continue
            expression = ' '.join((('+' if coefficient == 1 else '-') + ' ' if index else '') + label
                                  for index, (_, label, coefficient) in enumerate(recipe))
            record = dict(inputs[0])
            record.update({
                'Category': output_key[0], 'Label': output_key[1], 'Value': value,
                'Concept': '', 'SourceKind': 'calculated', 'IsCalculated': True,
                'SourceRawValue': None, 'SourceSemanticType': output_label,
                'SourceAdmissionRule': 'audited_financial_identity',
                'SourceDerivationFormula': expression,
                'SourceDisplaySign': -1 if output_label == 'Purchases of Investments' else 1,
                'SourceInputAccessions': ';'.join(str(item['Accession']) for item in inputs),
                'SourceInputFacts': json.dumps([
                    {field: str(item.get(field, '')) for field in
                     ('Category', 'Label', 'Period', 'Value', 'Concept', 'Accession', 'Start', 'End',
                      'Coefficient', 'SourceDerivationFormula', 'SourceInputFacts', 'SourceEvidenceAxes',
                      'SourceEvidenceMembers', 'SourceTableEvidence', 'SourceExactUnitKind',
                      'SourceReportedValue', 'SourceRawValue', 'SourceCellIdentity')}
                    for item in inputs
                ], ensure_ascii=False),
            })
            if not existing.empty:
                audit = audit.drop(index=existing.index)
            additions.append(record)
        if additions:
            audit = pd.concat([audit, pd.DataFrame(additions)], ignore_index=True)
    frame.attrs['fact_audit'] = audit
    return frame


def require_complete_output_evidence(frame: pd.DataFrame) -> pd.DataFrame:
    """Unproved output stays in the audit ledger, never becomes a trusted zero.

    This applies to every company, including missing calculation operands and
    taxonomy-generated zeroes that are not actually reported in the artifact.
    """
    audit = frame.attrs.get('fact_audit')
    if not isinstance(audit, pd.DataFrame):
        return frame
    audit = audit.copy()
    def text_present(value: Any) -> bool:
        return str(value).strip() not in {'', 'nan', 'None'}

    def table_evidence_present(value: Any) -> bool:
        if not text_present(value):
            return False
        try:
            payload = json.loads(str(value))
        except (TypeError, ValueError, json.JSONDecodeError):
            return False
        return isinstance(payload, dict) and bool(
            payload.get('artifact_sha256')
            and (payload.get('xpath') or payload.get('physical_locator') or payload.get('table_index') is not None)
        )

    def input_is_proven(item: Any) -> bool:
        if not isinstance(item, dict):
            return False
        nested = item.get('SourceInputFacts')
        if text_present(nested):
            try:
                children = json.loads(str(nested))
            except (TypeError, ValueError, json.JSONDecodeError):
                return False
            return bool(children) and all(input_is_proven(child) for child in children)
        required = ('Value', 'Accession', 'Concept', 'Start', 'End', 'Unit')
        if not all(text_present(item.get(field)) for field in required):
            return False
        # A numeric operand description is not source evidence.  At least one
        # relocatable/lexical field must survive from the selected SEC fact.
        return (
            str(item.get('SourceCellIdentity', '')).startswith('/')
            or table_evidence_present(item.get('SourceTableEvidence'))
            or text_present(item.get('SourceReportedValue'))
            or text_present(item.get('SourceRawValue'))
        )

    invalid_cells: set[tuple[str, str, str, str]] = set()
    valid_cells: set[tuple[str, str, str, str]] = set()

    def value_key(value: Any) -> str:
        try:
            parsed = Decimal(str(value))
        except (InvalidOperation, TypeError, ValueError):
            return str(value)
        return str(parsed.normalize()) if parsed.is_finite() else str(value)
    for index, row in audit.iterrows():
        if not str(row.get('Category', '')).startswith(('1_', '2_', '3_', '4')):
            continue
        def present(field: str) -> bool:
            return text_present(row.get(field, ''))
        derived = present('SourceDerivation') or 'calculated' in str(row.get('SourceKind'))
        reason = None
        if derived:
            try:
                inputs = json.loads(str(row.get('SourceInputFacts'))) if present('SourceInputFacts') else []
            except (TypeError, ValueError, json.JSONDecodeError):
                inputs = []
            if not inputs or not all(input_is_proven(item) for item in inputs):
                reason = 'MISSING_RECURSIVE_OPERANDS'
        elif not derived and not str(row.get('SourceCellIdentity', '')).startswith('/'):
            if not table_evidence_present(row.get('SourceTableEvidence')):
                reason = 'MISSING_RELOCATABLE_REPORTED_FACT'
        local_concept = str(row.get('Concept', '')).split('}')[-1].split(':')[-1]
        if derived and re.search(r'-Q[1-4]$', str(row.get('Period', ''))) and local_concept.startswith(('EarningsPerShare', 'WeightedAverage')):
            reason = 'NON_ADDITIVE_QUARTER_RECONSTRUCTION'
        # A lease payment alone cannot establish an aggregate of debt repayments
        # or prove that debt issuance was zero.
        if str(row.get('Concept')) == 'CalculatedDebtFamilyAggregate' and 'Finance Lease Principal Repaid' in str(row.get('SourceDerivationConcepts')):
            reason = 'UNPROVEN_DEBT_FAMILY_SCOPE'
        if reason:
            audit.at[index, 'SourceClassificationConfidence'] = 0.0
            audit.at[index, 'SourceAdmissionRule'] = reason
            audit.at[index, 'SourceEvidenceStatus'] = 'NOT_RUN'
            invalid_cells.add((str(row.get('Category')), str(row.get('Label')),
                               str(row.get('Period')), value_key(row.get('Value'))))
        else:
            valid_cells.add((str(row.get('Category')), str(row.get('Label')),
                             str(row.get('Period')), value_key(row.get('Value'))))
    # Keep rejected candidates in the audit ledger, but do not publish their
    # numeric value.  A second, proven candidate for the same business cell wins.
    for category, label, period, rejected_value in invalid_cells - valid_cells:
        key = (category, label)
        if (key in frame.index and period in frame.columns
                and value_key(frame.at[key, period]) == rejected_value):
            frame.at[key, period] = float('nan')
    # Late repair passes occasionally create a matrix cell without creating a
    # selected-fact/derivation record at all.  Period-level filing fallback is
    # useful for diagnostics, but it is not proof for a numeric front cell.
    for category, label in frame.index:
        if not str(category).startswith(('1_', '2_', '3_', '4')):
            continue
        for period in frame.columns:
            value = frame.at[(category, label), period]
            if pd.isna(pd.to_numeric(value, errors='coerce')):
                continue
            business_key = (str(category), str(label), str(period), value_key(value))
            if business_key not in valid_cells:
                frame.at[(category, label), period] = float('nan')
    frame.attrs['fact_audit'] = audit
    return frame
