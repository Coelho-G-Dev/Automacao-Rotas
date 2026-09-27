import json, sys
sys.stdout.reconfigure(encoding='utf-8')

path = r'C:\Users\gabri\.gemini\antigravity-ide\brain\489795f7-c92d-4d6e-aba8-32fe401e1f5a\.system_generated\logs\transcript.jsonl'
with open(path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for idx, line in enumerate(lines):
    data = json.loads(line)
    typ = data.get('type')
    cnt = str(data.get('content', ''))
    if typ == 'USER_INPUT':
        print(f"[{idx}] USER:")
        print(cnt)
        print("-" * 50)
