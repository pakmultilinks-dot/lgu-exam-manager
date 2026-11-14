"""Keep desktop regression runs out of the operator's saved workspace."""
import pytest

@pytest.fixture(autouse=True)
def isolated_workspace(monkeypatch,tmp_path):
    from app import App
    monkeypatch.setattr(App,'_autosave_path',lambda self:str(tmp_path/'operator-workspace.lguexam'))

    import workbook_schema
    monkeypatch.setattr(workbook_schema,'default_profile_path',lambda:tmp_path/'import-profiles.json')
