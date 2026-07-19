"""
setup.py
Compiles LinkHarvest's backend logic modules to C extensions (.pyd on
Windows) via Cython. app.py is intentionally excluded — it's the UI
entry point and PyInstaller's Analysis needs a plain .py to start from.

Usage:
    python setup.py build_ext --inplace

This produces one .pyd per module next to the .py source, plus .c
intermediate files and a build/ folder (both safe to delete/gitignore).
After building, the .py sources are no longer needed at runtime —
Python imports the .pyd transparently as if it were the .py file.
"""
from setuptools import setup
from Cython.Build import cythonize

MODULES = [
    "ai_assistant.py",
    "pdf_extractor.py",
    "dashboard_builder.py",
    "data_profiler.py",
    "donut_chart.py",
    "downloader.py",
    "sheet_editor.py",
    "donation.py",
    "updater.py",
]

setup(
    name="LinkHarvest",
    ext_modules=cythonize(
        MODULES,
        compiler_directives={
            "language_level": "3",
            "embedsignature": False,  # don't leak docstrings/signatures into the binary
        },
        # Strips out the .c files' own docstring-based introspection helpers
        annotate=False,
    ),
    zip_safe=False,
)
