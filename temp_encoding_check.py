import pathlib
root = pathlib.Path('.')
exts = {'.html', '.js', '.css', '.txt'}
bad = []
for p in root.rglob('*'):
    if 'node_modules' in p.parts:
        continue
    if p.suffix.lower() in exts and p.is_file():
        data = p.read_bytes()
        try:
            text = data.decode('utf-8')
        except Exception as e:
            bad.append((str(p), 'Invalid UTF-8', str(e)))
            continue
        if any(ord(ch) > 127 for ch in text):
            bad.append((str(p), 'Non-ASCII chars', None))
if not bad:
    print('OK: Checked project files are valid UTF-8 and contain only ASCII characters.')
else:
    for f, issue, err in bad:
        print(f + ' | ' + issue + ('' if err is None else ' | ' + err))
