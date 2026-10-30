# Shared helpers for the test-pack generators: output folder and the test registry
# (every test gets an ID, the file, where to look, what is tested and the expected result).
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, 'pylibs'))  # optional: pip install --target scripts/test-pack/pylibs ...
OUT = os.path.join(HERE, 'out')
ASSETS = os.path.join(HERE, 'assets')
os.makedirs(OUT, exist_ok=True)
os.makedirs(ASSETS, exist_ok=True)
FONTS = 'C:/Windows/Fonts'

class Registry:
    """Collects tests for one generator run; saved as tests-<name>.json for the checklist."""
    def __init__(self, name):
        self.name = name
        self.tests = []

    def add(self, test_id, file, where, what, expected, status=''):
        self.tests.append({'id': test_id, 'file': file, 'where': where, 'what': what, 'expected': expected, 'status': status})
        return f'{test_id} · {what}'

    def save(self):
        with open(os.path.join(HERE, f'tests-{self.name}.json'), 'w', encoding='utf8') as f:
            json.dump(self.tests, f, indent=1, ensure_ascii=False)
        print(f'{self.name}: {len(self.tests)} tests')
