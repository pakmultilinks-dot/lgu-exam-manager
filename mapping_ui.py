"""Open the workbook-structure review used by the import pipeline."""
def open_mapping(app):
    app.frames['import'].review_workbook()
