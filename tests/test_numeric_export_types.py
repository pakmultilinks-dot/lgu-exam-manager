def test_subject_counts_and_validation_rows_remain_excel_numbers(tmp_path):
    from test_academic_engine import _df
    from academic_engine import load_course_workbook
    from course_reports import export_report
    from openpyxl import load_workbook
    source=tmp_path/'source.xlsx';_df().to_excel(source,index=False)
    dataset=load_course_workbook(source)
    target=tmp_path/'totals.xlsx'
    export_report('cgpa_list','excel',dataset,dataset.courses,target)
    book=load_workbook(target,read_only=True,data_only=True)
    cell=book.active['K4']
    assert cell.value==2 and cell.data_type=='n'
    # GPA strings intentionally retain two displayed decimals, rounded half-up
    # like the official gazettes (2.825 rounds to 2.83, not truncated to 2.82).
    assert book.active['H4'].value=='2.83'
    book.close()
