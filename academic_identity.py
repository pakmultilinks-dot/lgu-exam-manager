"""Program identity ignores typographic case, spacing and punctuation only."""
from functools import lru_cache
import unicodedata

@lru_cache(maxsize=8192)
def program_key(value):
    return ''.join(char for char in unicodedata.normalize('NFKC',str(value)).upper() if char.isalnum())

@lru_cache(maxsize=8192)
def registration_key(value):
    return unicodedata.normalize('NFKC',str(value)).strip().upper()

def student_count(rows):
    return len({registration_key(course.registration) for course in rows})
