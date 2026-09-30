import sys, traceback, importlib.util, pathlib
root = pathlib.Path(__file__).parent
for f in sorted(root.glob("test_*.py")):
    spec = importlib.util.spec_from_file_location(f.stem, f)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    passed = failed = 0
    for name in dir(mod):
        if name.startswith("test_"):
            fn = getattr(mod, name)
            if callable(fn):
                try:
                    fn()
                    passed += 1
                except Exception:
                    failed += 1
                    print(f"FAIL {f.name}::{name}")
                    traceback.print_exc()
    print(f"{f.name}: {passed} passed, {failed} failed")
