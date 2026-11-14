"""Operator-facing errors; diagnostic exception classes stay in the support log."""
def friendly_error(error):
    kind,_,detail=str(error).partition(': ')
    if kind=='ValueError' and detail:
        if 'not in list' not in detail and '__mapped' not in detail:return detail
    if kind in ('PermissionError','FileExistsError'):return 'The file cannot be written. Close it in other applications or choose a different destination.'
    if kind=='FileNotFoundError':return 'The selected file is no longer available. Select it again from its current location.'
    if kind in ('BadZipFile','InvalidFileException'):return 'This workbook or saved session could not be read. Check that the file opens correctly in Excel, then try a fresh copy.'
    if kind in ('ImportError','ModuleNotFoundError'):return 'An application component is missing. Extract and run the complete application package.'
    return 'The operation could not complete. Your previous saved files are preserved. Review the workbook structure or try another destination. If it continues, use the support log.'
