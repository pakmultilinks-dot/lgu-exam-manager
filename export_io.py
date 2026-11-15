"""Commit an export only when its writer has completed successfully."""
from contextlib import contextmanager
from pathlib import Path
import os
import tempfile


def protect_source(dataset,path):
    target=Path(path).resolve()
    paths=getattr(dataset,'source_paths',[])
    if paths:
        # A case-only rename is the same file on Windows (case-insensitive),
        # so reject it even though the path does not exist yet and samefile
        # cannot see it. A matching name in another directory stays allowed.
        protected=any(target==Path(source).resolve() or (target.exists() and Path(source).exists() and os.path.samefile(target,source)) or (target.parent==Path(source).resolve().parent and target.name.casefold()==Path(source).name.casefold()) for source in paths)
    else:
        protected=any(Path(c.source_file).name.casefold()==target.name.casefold() for c in dataset.courses)
    if protected:raise ValueError('Output must not overwrite the original workbook.')


@contextmanager
def atomic_output(path):
    target=Path(path)
    handle,name=tempfile.mkstemp(prefix='.lgu-export-',suffix=target.suffix,dir=target.parent)
    os.close(handle)
    temporary=Path(name)
    try:
        yield temporary
        if not temporary.stat().st_size:raise ValueError('Export produced an empty file.')
        os.replace(temporary,target)
    finally:
        temporary.unlink(missing_ok=True)
