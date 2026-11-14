def test_default_launch_stays_empty_with_saved_workspace(tmp_path,monkeypatch):
    from app import App
    from academic_engine import ResultDataset
    from test_cross_program_matrix import attempt
    import session_store
    path=tmp_path/'operator-workspace.lguexam'
    saved=ResultDataset([attempt()],[],[],'prior-import.xlsx')
    session_store.save(path,'course',None,None,{}, {},saved)
    def forbidden(self):
        raise AssertionError('A normal startup must not restore saved records')
    monkeypatch.setattr(App,'_restore_autosave',forbidden)
    app=App()
    try:
        app.withdraw();app.update()
        assert app.academic_dataset is None and app.raw_df is None and app.df is None
        assert app.source_path is None and not app.job_busy
        assert session_store.restore(path)[1].courses==saved.courses
    finally:
        for job in app.tk.call('after','info'):app.tk.call('after','cancel',job)
        app.destroy()
