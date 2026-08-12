orginal = "Adjustable Elbow"
words = orginal.split()
thing = " ".join(words).lower()
print(f'{thing}')


def _clean(text):
    return ' '.join(text.split()).lower()


def _clean_old(text):
    return text.strip().lower()


new = "the HappY cow   came _home  "

cleaned1 = _clean(new)
cleaned2 = _clean_old(new)
print(cleaned1)
print(cleaned2)
print(repr(cleaned1))
