"""Editable draft formatting only. Does not authorize official publication."""
import json
from pathlib import Path

def load_template():
    path=Path(__file__).with_name('draft_template.json')
    defaults={'institution':'LAHORE GARRISON UNIVERSITY - EXAMINATION BRANCH','title':'DRAFT RESULT REGISTER - NOT FOR PUBLICATION','statement_title':'DRAFT RESULT STATEMENT - NOT FOR PUBLICATION','include_contents':True}
    if path.is_file():defaults.update(json.loads(path.read_text(encoding='utf-8')))
    for key in ('institution','title','statement_title'):
        if not isinstance(defaults[key],str) or len(defaults[key])>180:raise ValueError('Invalid draft template text: '+key)
    # The mandatory draft notice is rendered separately, even for customized headings.
    return defaults
