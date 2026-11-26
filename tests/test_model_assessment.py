import json
import logging

import pytest

from app.src.JBGModelAssessment import parse_assessment, summarize_assessments
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI


ASSESSMENT = dict(priority='significant', category='complex_expression',
                  confidence='high', meaning_risk='low', requires_rewrite=False)


def suggestor(enabled):
    ai = JBGLangImprovSuggestorAI('test', 'test', 'base', 0,
                                  logging.getLogger(__name__), assess_changes=enabled)
    ai.json_structured_document = {'type': 'docx', 'elements': [
        {'type': 'paragraph', 'element_id': 'paragraph_1', 'text': 'Du ska erlägga betalning.'}]}
    return ai


def raw():
    return dict(type='paragraph', element_id='paragraph_1', old='Du ska erlägga betalning.',
                new='Du ska betala.', motivation='Enklare uttryck.')


@pytest.mark.parametrize('value', [None, [], {}, {'priority': []},
    dict(ASSESSMENT, priority='unknown'), dict(ASSESSMENT, confidence='HIGH'),
    dict(ASSESSMENT, meaning_risk='unknown'), dict(ASSESSMENT, category='unknown'),
    dict(ASSESSMENT, requires_rewrite='false'), dict(ASSESSMENT, requires_rewrite=0)])
def test_malformed_assessment_is_advisory(value):
    assert parse_assessment({'model_assessment': value}) == (None, 'invalid')


@pytest.mark.parametrize('status,value', [('missing', None), ('invalid', {}),
    ('valid', ASSESSMENT), ('valid', dict(ASSESSMENT, priority='cosmetic', confidence='low', meaning_risk='high', requires_rewrite=True))])
def test_shadow_keeps_identical_filter_decisions(tmp_path, status, value):
    item = raw()
    if status != 'missing':
        item['model_assessment'] = value
    legacy, shadow = suggestor(False), suggestor(True)
    for ai in (legacy, shadow):
        ai.json_suggestions = ai._parse_model_response(json.dumps([item]))
        validated = ai._postprocess_suggestions(ai.json_suggestions)
        ai.validated_suggestions, ai.filtered_review = ai.quality_filter.filter(validated)
    assert len(legacy.validated_suggestions) == len(shadow.validated_suggestions) == 1
    s = shadow.validated_suggestions[0]
    assert s.assessment_status == status
    assert s.old != item['old']  # assessment survives span minimization
    expected = legacy._suggested_change_to_output_dict(legacy.validated_suggestions[0])
    actual = shadow._suggested_change_to_output_dict(s)
    assert actual.pop('assessment_status') == status
    assert actual.pop('model_assessment', None) == (value if status == 'valid' else None)
    assert actual == expected
    assert [r.accepted for r in legacy.filtered_review] == [r.accepted for r in shadow.filtered_review]
    for filename, save in [('suggestions.json', shadow.save_as_json), ('filter.json', shadow.save_filter_report)]:
        path = tmp_path / filename
        save(path)
        obj = json.loads(path.read_text(encoding='utf-8'))[0]
        assert obj.get('suggestion', obj)['assessment_status'] == status
    summary = summarize_assessments([item], shadow.validated_suggestions)
    assert summary['raw'][status] == summary['accepted'][status] == 1
    assert summary['mode'] == 'shadow'


def test_prompt_opt_in_and_no_safety_override():
    assert suggestor(False).policy_prompt == 'base'
    ai = suggestor(True)
    assert 'model_assessment' in ai.policy_prompt
    item = dict(raw(), old='not present', model_assessment=ASSESSMENT, safe_to_apply=True)
    assert ai._postprocess_suggestions([item]) == []


def test_assessment_prompt_uses_independent_risk_and_priority():
    from app.src.JBGModelAssessment import assessment_prompt
    prompt = assessment_prompt()
    assert 'oberoende av nivå' in prompt
    assert 'avgör inte i sig prioriteten' in prompt
    assert 'Bedöm risken separat' in prompt
    assert 'klassificera inte osäker tolkning som low' in prompt
