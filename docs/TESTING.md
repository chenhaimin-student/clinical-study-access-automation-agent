# Testing

Run checks sequentially:

```bash
python -m pip install -r requirements-dev.txt
ruff check --select E9,F63,F7,F82 .
python -m compileall -q scripts tests ops
python tests/self_test.py
python scripts/setup_study_from_excel.py --excel-file assets/application-template.xlsx --project-id DEMO001 --parse-only
```

`tests/self_test.py` uses temporary files and mocked responses. It must not contact a production mailbox or API.

`ops/gateway_monitor.py` is macOS-specific because it uses `fcntl` and
`launchctl`. Syntax can be checked on Windows, but runtime behavior belongs in
an isolated macOS environment. The core Python workflow is separately covered
by the Ubuntu GitHub Actions matrix.
