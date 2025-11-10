import json, pathlib
path = pathlib.Path('faculty-allocations-2025-10-28 (6).json')
data = json.loads(path.read_text(encoding='utf-8'))
result = [s for s in data['subjects'] if '7TEC' in s]
print('subjects:', len(result))
print('list:')
for item in result:
    print(item.encode('unicode_escape').decode())
