# Doc Generator Tool

Standalone document generator files separated from the main landing analyzer.

## Local files

- `doc_generator.py` - main PySide6 app.
- `AdsMailImporter.gs` - Apps Script template loaded by the app.
- `octo_scan_log.json` - persistent scan history used by the Octo widget.
- `assets/docgen.ico` - DocGenerator executable icon.
- `DocGenerator.spec` - PyInstaller spec for this tool.
- `scripts/build_docgen_exe.ps1` - one-file EXE build script.
- `scripts/generate_docgen_icon.py` - icon generator used by the build script.
- `archive/` - old DocGenerator backup.
- `build/` and `dist/` - DocGenerator build artifacts.

## Shared project dependency

The tool still imports the main project `../core` package for site profiling and
landing analysis. This keeps the shared analyzer code in one place while the
DocGenerator-specific files stay isolated here.

## Python dependencies

Runtime packages used by `doc_generator.py`:

- `PySide6`
- `python-docx`
- `fpdf`
- `openai`
- `Pillow`
- `requests`
- `beautifulsoup4`

Optional imports detected in specific features:

- `trafilatura`
- `langdetect`
- `cairosvg`
- `se_simulator` - referenced by the fast parser, but not present in this workspace.

## Run

From the project root:

```powershell
.\.venv\Scripts\python.exe .\doc_generator_tool\doc_generator.py
```

## Build

From the project root:

```powershell
.\doc_generator_tool\scripts\build_docgen_exe.ps1
```

The output is written under `doc_generator_tool\dist`.
