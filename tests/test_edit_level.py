import logging
from pathlib import Path
from unittest.mock import patch

import pytest
from fastapi.testclient import TestClient

from app import main
from app.src.JBGEditLevel import get_edit_profile, edit_level_config, build_local_prompt, PROMPT_PROFILES
from app.src.JBGLanguageImprover import JBGLanguageImprover


@pytest.mark.parametrize('level,key', [(1,'very_conservative'), (20,'very_conservative'),
    (21,'conservative'), (40,'conservative'), (41,'balanced'), (50,'balanced'),
    (60,'balanced'), (61,'thorough'), (80,'thorough'), (81,'very_thorough'), (100,'very_thorough')])
def test_profile_boundaries(level, key):
    assert get_edit_profile(level).key == key


@pytest.mark.parametrize('level', [0, 101, -1, True, 50.5, '50'])
def test_invalid_direct_level(level):
    with pytest.raises(ValueError):
        get_edit_profile(level)


def test_config():
    response = TestClient(main.app).get('/config')
    assert response.status_code == 200
    assert response.json()['edit_level'] == edit_level_config()
    assert edit_level_config()['default'] == 70


@pytest.mark.parametrize('level', ['0', '101', '-1', '50.5', 'abc'])
def test_invalid_upload(level):
    with patch.object(main, 'setup_run_logger') as logger:
        response = TestClient(main.app).post('/upload/', data={
            'api_key':'test', 'model':'test', 'edit_level':level},
            files={'file':('test.docx', b'test')})
    assert response.status_code == 422
    logger.assert_not_called()


@pytest.mark.parametrize('level,key', [(None,None), (50,'balanced'), (70,'thorough'), (81,'very_thorough')])
def test_upload_passes_level_and_builds_local_prompt(tmp_path, monkeypatch, level, key):
    for name in ('UPLOAD_DIR','RESULTS_DIR','LOG_DIR'):
        directory = tmp_path / name
        directory.mkdir()
        monkeypatch.setattr(main, name, str(directory))
    monkeypatch.setattr(main, 'jobs_lang', {})
    data = {'api_key':'test', 'model':'test', 'editable_prompt':'test prompt'}
    if level is not None:
        data['edit_level'] = str(level)
    result = tmp_path / 'mock-result.docx'
    result.write_bytes(b'test')
    with patch.object(main, 'JBGLanguageImprover') as mock:
        mock.return_value.run.return_value = str(result)
        response = TestClient(main.app).post('/upload/', data=data,
            files={'file':('test.docx', b'test')})
    assert response.status_code == 200
    assert main.jobs_lang[response.json()['job_id']]['error'] is None
    kwargs = mock.call_args.kwargs
    assert kwargs['edit_level'] == level
    _, before, after = main.load_prompt_parts(str(Path(main.BASE_DIR)/'policy'/'prompt_policy.md'))
    assert kwargs['prompt_policy'] == f'{before}\n\ntest prompt\n\n{after}'
    improver = JBGLanguageImprover(**kwargs)
    summary = improver.run_summary.to_dict()
    assert summary['edit_level'] == level
    assert summary['edit_profile'] == key
    assert summary['edit_level_applied'] is False
    assert improver.prompt_policy == build_local_prompt(kwargs['prompt_policy'], level)


@pytest.mark.parametrize('level', [1, 21, 41, 61, 81])
def test_local_prompt_profile(level):
    base = 'original policy\nJSON rules'
    result = build_local_prompt(base, level)
    assert result.startswith(base + '\n\n')
    assert PROMPT_PROFILES[get_edit_profile(level).key] in result
    assert 'JSON-format gäller oförändrat' in result
    assert 'Skyddsregler' in result or 'skyddsregler' in result


def test_legacy_prompt_is_exactly_unchanged():
    assert build_local_prompt(' original\n', None) == ' original\n'


@pytest.mark.parametrize('applied', [True, False])
def test_about_section_reports_level(applied):
    from app.src.JBGAboutSectionRenderer import build_about_section_blocks
    blocks = build_about_section_blocks({'edit_level': 73, 'edit_level_applied': applied}, False)
    text = '\n'.join(block.text for block in blocks)
    assert '73 av 100 – Noggrann' in text
    assert ('tillämpades inte' in text) is (not applied)
    assert ('Nivån styr instruktionen' in text) is applied


def test_about_section_legacy_has_no_level():
    from app.src.JBGAboutSectionRenderer import build_about_section_blocks
    assert not any('Bearbetningsnivå' in b.text for b in build_about_section_blocks({}, False))


@pytest.mark.parametrize('level,success,expected', [(73, True, True), (73, False, False), (None, True, False)])
def test_local_suggestor_receives_profile_and_reports_usage(tmp_path, level, success, expected):
    improver = JBGLanguageImprover(
        input_path=str(tmp_path / 'test.docx'), api_key='test', model='test',
        prompt_policy='base policy', temperature=0, include_motivations=False,
        logger=logging.getLogger(__name__), edit_level=level, save_intermediate_json=False)
    improver.structure = {}
    with patch('app.src.JBGLanguageImprover.JBGLangImprovSuggestorAI') as mock:
        ai = mock.return_value
        ai.json_suggestions = ai.filtered_review = ai.validated_suggestions = []
        def simulate():
            if success:
                improver.usage_tracker.record('local', 'test', None)
            else:
                improver.usage_tracker.record_failure('local', 'test')
        ai.suggest_changes_token_aware_batching.side_effect = simulate
        improver._generate_suggestions()
        assert mock.call_args.kwargs['prompt_policy'] == build_local_prompt('base policy', level)
    assert improver.run_summary.edit_level_applied is expected


def test_profiles_are_cumulative_and_share_assessment_definitions():
    from app.src.JBGEditLevel import PRIORITY_DEFINITIONS, PROFILES, REVIEW_GUARDRAILS
    from app.src.JBGModelAssessment import assessment_prompt
    previous = []
    for profile in PROFILES:
        criteria = PROMPT_PROFILES[profile.key].splitlines()
        assert criteria[:-1] == previous
        assert len(criteria) == len(previous) + 1
        previous = criteria
        assert REVIEW_GUARDRAILS in build_local_prompt('base', profile.max_level)
    for key in ('critical', 'significant', 'moderate', 'minor'):
        assert PRIORITY_DEFINITIONS[key] in PROMPT_PROFILES[PROFILES[-1].key]
    for definition in PRIORITY_DEFINITIONS.values():
        assert definition in assessment_prompt()
    assert PRIORITY_DEFINITIONS['cosmetic'] not in PROMPT_PROFILES[PROFILES[-1].key]


def test_default_profile_matches_static_gui():
    from app.src.JBGEditLevel import DEFAULT_LEVEL
    profile = get_edit_profile(DEFAULT_LEVEL)
    html = (Path(main.BASE_DIR) / 'templates' / 'index.html').read_text(encoding='utf-8')
    assert f'{DEFAULT_LEVEL} av 100 – {profile.name}' in html
    assert profile.description in html


def test_every_level_has_its_own_prompt_even_within_profile():
    prompts = [build_local_prompt('base', level) for level in range(1, 101)]
    assert len(set(prompts)) == 100
    for level, prompt in enumerate(prompts, start=1):
        assert f'Önskad ingreppsnivå: {level}/100' in prompt
        assert 'Använd det exakta värdet även inom profilen' in prompt
    assert get_edit_profile(10) == get_edit_profile(12)
    assert prompts[9] != prompts[11]


def test_plain_language_descriptions_and_report():
    from app.src.JBGEditLevel import PROFILES
    from app.src.JBGAboutSectionRenderer import build_about_section_blocks
    for profile in PROFILES:
        assert profile.description.startswith('Föreslår')
        assert 'ekonomi' not in profile.description
        assert 'ekonomi' not in build_local_prompt('base', profile.max_level)
        blocks = build_about_section_blocks({'edit_level': profile.max_level,
                                             'edit_level_applied': True}, False)
        assert profile.description in '\n'.join(block.text for block in blocks)
