"""
Runs the offline tests (no PowerWorld needed): every `test_*` function of tests/test_*.py.
    python tests/run.py            all tests
    python tests/run.py makeup     only the tests whose name contains "makeup"
Stops with an error when two files define a test of the same name (the second would shadow the first).
"""
import importlib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def collect():
    """[(name, function)] of every test in tests/test_*.py, files in name order; exits when a name appears twice."""
    sys.path.insert(0, str(HERE))
    tests, where = [], {}
    for fp in sorted(HERE.glob('test_*.py')):
        module = importlib.import_module(fp.stem)
        for name, f in vars(module).items():
            if name.startswith('test_') and callable(f) and getattr(f, '__module__', None) == module.__name__:
                if name in where:
                    sys.exit(f'Duplicate test name {name}: in {where[name]} and {fp.name}.')
                where[name] = fp.name
                tests.append((name, f))
    return tests


if __name__ == '__main__':
    tests = [(name, f) for name, f in collect() if not sys.argv[1:] or sys.argv[1] in name]
    for _, test in tests:
        test()
    print(f'{len(tests)} tests passed.')
