"""Gazette fonts: bundled DejaVu first, then matplotlib, then system fonts.

A gazette export must never fail because a font library is missing, so the
app ships its own DejaVu fonts in gazette_assets/fonts and only falls back
to matplotlib or system fonts when a style file is not bundled.
"""
from pathlib import Path

def _font_file(name):
    """Find one font file without depending on any single library."""
    candidates=[Path(__file__).with_name('gazette_assets')/'fonts'/name]
    try:
        import matplotlib
        candidates.append(Path(matplotlib.get_data_path())/'fonts'/'ttf'/name)
    except Exception:
        pass
    candidates+=[
        Path('/usr/share/fonts/truetype/dejavu')/name,
        Path('/usr/local/share/fonts/truetype/dejavu')/name,
        Path('C:/Windows/Fonts')/name,
    ]
    for path in candidates:
        if path.is_file():
            return str(path)
    return None

def add_font_family(pdf, family, styles):
    """Add a font family; a missing style file falls back to the regular file."""
    regular=_font_file(styles[''])
    if regular is None:
        raise ValueError('Gazette fonts are missing. Reinstall the app so the gazette_assets/fonts folder is present.')
    for style, name in styles.items():
        pdf.add_font(family, style, _font_file(name) or regular)

def add_fonts(pdf):
    add_font_family(pdf, 'LGU', {'': 'DejaVuSans.ttf', 'B': 'DejaVuSans-Bold.ttf', 'I': 'DejaVuSans-Oblique.ttf'})
