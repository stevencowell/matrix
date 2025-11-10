import json, pathlib
path = pathlib.Path('faculty-allocations-2025-10-28 (6).json')
data = json.loads(path.read_text(encoding='utf-8'))
wed_subjects = [s for s in data['subjects'] if '_Wed' in s]
print('Wed subjects', wed_subjects)
