"""Active-workspace transitions; import settings never inherit a prior file's scope."""
from pathlib import Path
import shutil


def import_settings(dataset, previous):
    settings=dict(previous)
    for field in ('gazette_term','cover_department','cover_period'):settings.pop(field,None)
    settings['gazette_template']='auto'
    if dataset is not None:
        from gazette_options import current_term
        try:settings['gazette_term']=current_term(dataset,settings)
        except ValueError:pass
    return settings


def persist_import(path, source, dataset, settings, raw=None, mapping=None):
    import session_store
    path=Path(path)
    if path.is_file():shutil.copy2(path,path.with_name(path.stem+'.previous'+path.suffix))
    session_store.save(path,'course' if dataset is not None else 'wide',source,raw,mapping or {},settings,dataset)
