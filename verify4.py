import json, pathlib
path = pathlib.Path('faculty-allocations-2025-10-28 (6).json')
data = json.loads(path.read_text(encoding='utf-8'))
items = [s for s in data['subjects'] if s.startswith('7TEC')]
for s in items:
    print(s.encode('unicode_escape').decode())
