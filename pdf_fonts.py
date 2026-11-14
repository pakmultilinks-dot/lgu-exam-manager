from pathlib import Path

def add_fonts(pdf):
    import matplotlib
    root=Path(matplotlib.get_data_path())/'fonts'/'ttf'
    for style,file in [('', 'DejaVuSans.ttf'),('B','DejaVuSans-Bold.ttf'),('I','DejaVuSans-Oblique.ttf')]:
        pdf.add_font('LGU',style,str(root/file))
