import logging
from pathlib import Path
from unittest.mock import patch

import pytest
from fastapi.testclient import TestClient

from app import main
from app.src.JBGEditLevel import get_edit_profile, edit_level_config
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
    assert edit_level_config()['default'] == 50


@pytest.mark.parametrize('level', ['0', '101', '-1', '50.5', 'abc'])
def test_invalid_upload(level):
    with patch.object(main, 'setup_run_logger') as logger:
        response = TestClient(main.app).post('/upload/', data={
            'api_key':'test', 'model':'test', 'edit_level':level},
            files={'file':('test.docx', b'test')})
    assert response.status_code == 422
    logger.assert_not_called()


@pytest.mark.parametrize('level,key', [(None,None), (50,'balanced'), (81,'very_thorough')])
def test_upload_passes_level_without_changing_prompt(tmp_path, monkeypatch, level, key):
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
    assert improver.prompt_policy == kwargs['prompt_policy']
