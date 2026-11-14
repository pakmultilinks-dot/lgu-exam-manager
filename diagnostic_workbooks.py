"""Opt-in checks only. These fictional records are never startup data."""
from pathlib import Path
import pandas as pd


def check_unseen_workbooks(app,root,settle,passed):
    from academic_engine import load_course_workbook
    from calculation_policy import calculation_attempts
    from workbook_schema import load_profiles
    from gazette_options import resolve_template_groups
    import workbook_schema
    import import_review
    original=workbook_schema.default_profile_path
    original_ui=import_review.default_profile_path
    profile=root/'isolated-profiles.json'
    workbook_schema.default_profile_path=lambda:profile
    import_review.default_profile_path=lambda:profile
    try:
        old=app.academic_dataset
        source=root/'FICTIONAL-unfamiliar-columns.xlsx'
        pd.DataFrame([['R-31','NA','SU29',2,'B',3,'A'],['R-32','Person Q','SU29',4,'F','','']],columns=['Identity','Person','When','Units x','Outcome x','Units y','Outcome y']).to_excel(source,index=False)
        importer=app.frames['import']
        importer._smart_load(str(source));settle()
        wizard=importer.import_review
        assert app.academic_dataset is old and wizard.winfo_exists()
        wizard.destroy();app.update()
        assert app.academic_dataset is old
        passed('unrecognized workbook opens mapping; cancelling preserves active data')
        importer._smart_load(str(source));settle();wizard=importer.import_review
        wizard.mode.set('Courses across columns');wizard.change_mode('Courses across columns')
        for role,column in {'registration':'Identity','name':'Person','term':'When'}.items():wizard.controls[role].set(column)
        wizard.fixed['degree'].insert(0,'New Unseen Program')
        wizard.fixed['department'].insert(0,'Computer Science')
        for code,title,credits,grade in [('NEW-91','Unseen First','Units x','Outcome x'),('NEW-92','Unseen Second','Units y','Outcome y')]:
            wizard.add_block();code_entry,title_entry,menus,fixed_credit=wizard.blocks[-1]
            code_entry.insert(0,code);title_entry.insert(0,title)
            menus['credits'].set(credits);menus['grade'].set(grade)
        wizard.store();wizard.apply();settle()
        ds=app.academic_dataset
        assert len(ds.courses)==3 and len(ds.source_rows)==2 and ds.courses[0].name=='NA'
        assert [c.source_excel_row for c in ds.courses]==[2,2,3]
        assert len(resolve_template_groups('auto',ds.courses))==1
        record=next(iter(load_profiles(profile).values()))
        assert record['requires_review'] and record['defaults']=={}
        assert all(not block['defaults'] for block in record['course_blocks'])
        assert 'Unseen' not in profile.read_text()
        passed('desktop wide-course mapping; original row audit; unseen program uses supplied department')
        passed('remembered structures exclude entered programs, terms, credits and course metadata')
        legacy=Path(__file__).with_name('diagnostic_assets')/'FICTIONAL-reader-check.xls'
        older=load_course_workbook(legacy)
        assert len(older.courses)==1 and calculation_attempts(older)[0].quality_points=='6'
        passed('packaged XLS reader and explicit source points per credit')
    finally:
        workbook_schema.default_profile_path=original
        import_review.default_profile_path=original_ui
